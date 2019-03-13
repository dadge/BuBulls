using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace TestDocx
{
    public class Subject
    {
        public string Name { get; set; }
        public int ChoiceCount { get; set; }
        public List<string> Acquisitions {get;set;}

    }

    public class SchoolReportTemplate
    {
        public List<Subject> Subjects { get; set; }
    }

    public class ExcelParser
    {
        public ExcelParser()
        {
        }

        public static SchoolReportTemplate FromExcel(string filename)
        {
            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(filename, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }
            var toRet = new SchoolReportTemplate();
            toRet.Subjects = new List<Subject>();
            ISheet sheet = hssfwb.GetSheet("Matières");
            Subject currentSubject = null;
            for (int row = 1; row <= sheet.LastRowNum; row++) // skip first line
            {
                if (sheet.GetRow(row) != null) //null is when the row only contains empty cells 
                {
                    var firstCol = sheet.GetRow(row).GetCell(0);
                    var secondCol = sheet.GetRow(row).GetCell(1);
                    if(firstCol != null && !String.IsNullOrEmpty(firstCol.StringCellValue))
                    {
                        currentSubject = new Subject()
                        {
                            ChoiceCount = 0,
                            Name = firstCol.StringCellValue,
                            Acquisitions = new List<string>()
                        };
                        toRet.Subjects.Add(currentSubject);
                    }
                    if (secondCol != null && !String.IsNullOrEmpty(secondCol.StringCellValue))
                    {
                        currentSubject.Acquisitions.Add(secondCol.StringCellValue);
                        Console.WriteLine(string.Format("adding {0} --> {1}", currentSubject.Name,secondCol.StringCellValue ));
                    }


                }
            }
            return toRet;
        }

        public class ParsingError
        {
            public String Eleve { get; set; }
            public string Compétence { get; set; }
            public int RowIndex { get; set; }
            public int ColumnIndex { get; set; }
            public string ColumnName
            {
                get
                {
                    return NPOI.SS.Util.CellReference.ConvertNumToColString(ColumnIndex);
                }
            }
            public string Message { get; set; }
        }

        public class ParsingResult
        {
            public List<PersonalizedSchoolReportData> Data { get; set; }
            public List<ParsingError> Errors { get; set; }
        }

        public class PersonalizedSchoolReportData {
            public string FirstName { get; set; }
            public string LastName { get; set; }

            public Dictionary<string, int> Acquisitions { get; set; }
            public Dictionary<string, string> Comments { get; set; }
        }

        public static ParsingResult ParseAcquisitions(string filename) 
        {
            ParsingResult result = new ParsingResult();
            result.Data = new List<PersonalizedSchoolReportData>();
            result.Errors = new List<ParsingError>();
            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(filename, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }

            var template = FromExcel(filename); // so what, perf importants here??
            var aquisitions = template.Subjects.SelectMany(p => p.Acquisitions);

            var toRet = new List<PersonalizedSchoolReportData>();


            ISheet sheet = hssfwb.GetSheet("Acquisitions");

            var aquisitionRow = sheet.GetRow(1);


            for (int row = 3; row <= sheet.LastRowNum; row++) // skip first line
            {
                if (sheet.GetRow(row) != null) //null is when the row only contains empty cells 
                {
                    var firstCol = sheet.GetRow(row).GetCell(0);
                    var secondCol = sheet.GetRow(row).GetCell(1);
                    string firstname = null;
                    string lastname = null;
                    try
                    {
                        if (firstCol != null && !String.IsNullOrEmpty(firstCol.StringCellValue) && firstCol.StringCellValue != "0")
                        {
                            firstname = firstCol.StringCellValue;
                        }
                        if (secondCol != null && !String.IsNullOrEmpty(secondCol.StringCellValue) && secondCol.StringCellValue != "0")
                        {
                            lastname = secondCol.StringCellValue;
                        }
                    }
                    catch(Exception ee)
                    {                        
                    }

                    if(!string.IsNullOrEmpty(firstname) && !string.IsNullOrEmpty(lastname))
                    {
                        PersonalizedSchoolReportData data = new PersonalizedSchoolReportData()
                        {
                            FirstName = firstname,
                            LastName = lastname,
                            Acquisitions = new Dictionary<string, int>(),
                            Comments = new Dictionary<string, string>()

                        };
                        toRet.Add(data);

                        for (int c = 2; c < aquisitionRow.LastCellNum; c++)
                        {
                            var aquisitionCell = aquisitionRow.GetCell(c);
                            if(aquisitionCell != null && aquisitionCell.CellType != CellType.Numeric)
                            {
                                string acText = null;

                                try{
                                    acText = aquisitionCell.StringCellValue;
                                }catch { }

                                if (acText != null)
                                {
                                    var cell = sheet.GetRow(row).GetCell(c);
                                    if (cell != null && cell.CellType != CellType.Blank)
                                    {
                                        if (!data.Acquisitions.ContainsKey(aquisitionCell.StringCellValue))
                                            try
                                            {
                                                data.Acquisitions.Add(aquisitionCell.StringCellValue, (int)cell.NumericCellValue);
                                            } catch (Exception eee)
                                            {
                                                if(string.IsNullOrWhiteSpace(cell.StringCellValue))
                                                    data.Acquisitions.Add(aquisitionCell.StringCellValue, -1);
                                                else
                                                    result.Errors.Add(new ParsingError()
                                                    {
                                                        Eleve = firstname + " " + lastname,
                                                        Compétence = acText,
                                                        ColumnIndex = c,
                                                        RowIndex = row+1,
                                                        Message = "La valeur doit être numérique mais elle contient: '" + cell.StringCellValue + "'. " + eee.Message
                                                    });
                                            }
                                            
                                    }                                        
                                    else
                                        if (!data.Acquisitions.ContainsKey(aquisitionCell.StringCellValue))
                                            data.Acquisitions.Add(aquisitionCell.StringCellValue, -1);

                                }                            }
                        }
                    }
                }
            }



            // now comments
            /*
            sheet = hssfwb.GetSheet("Commentaires");

            var subjectRow = sheet.GetRow(1);


            for (int row = 3; row <= sheet.LastRowNum; row++) // skip first line
            {
                if (sheet.GetRow(row) != null) //null is when the row only contains empty cells 
                {
                    var firstCol = sheet.GetRow(row).GetCell(0);
                    var secondCol = sheet.GetRow(row).GetCell(1);
                    string firstname = null;
                    string lastname = null;
                    try
                    {
                        if (firstCol != null && !String.IsNullOrEmpty(firstCol.StringCellValue) && firstCol.StringCellValue != "0")
                        {
                            firstname = firstCol.StringCellValue;
                        }
                        if (secondCol != null && !String.IsNullOrEmpty(secondCol.StringCellValue) && secondCol.StringCellValue != "0")
                        {
                            lastname = secondCol.StringCellValue;
                        }
                    }
                    catch { }

                    if (!string.IsNullOrEmpty(firstname) && !string.IsNullOrEmpty(lastname))
                    {
                        var data = toRet.Single(p => p.FirstName == firstname && p.LastName == lastname);

                        for (int c = 2; c < subjectRow.LastCellNum; c++)
                        {
                            var subjectCell = subjectRow.GetCell(c);
                            if (subjectCell != null && subjectCell.CellType != CellType.Numeric)
                            {
                                string acText = null;

                                try
                                {
                                    acText = subjectCell.StringCellValue;
                                    if (acText != null)
                                    {
                                        var cell = sheet.GetRow(row).GetCell(c);
                                        if (cell != null)
                                            data.Comments.Add(acText, cell.StringCellValue);

                                    }
                                }
                                catch { }


                            }
                        }
                    }
                }
            }
            */
            result.Data = toRet;
            return result;
        }
    }
}

