using System;
using System.Linq;
using static TestDocx.ExcelParser;
using Xceed.Words.NET;

namespace TestDocx
{
    public class PersonalizedSchoolReportGenerator
    {
        public PersonalizedSchoolReportGenerator()
        {
        }

        public static void GeneratePersonalizedReport(PersonalizedSchoolReportData data, SchoolReportTemplate GlobalModel, string outputFolder, string intermediateTemplate){
            string outputFile = System.IO.Path.Combine(outputFolder, data.FirstName + "_" + data.LastName + ".docx");

            using (DocX document = DocX.Load(intermediateTemplate))
            {
                document.ReplaceText("{{fn}}", data.FirstName);
                document.ReplaceText("{{ln}}", data.LastName);

                // Check if all the replace patterns are used in the loaded document.
                foreach (var aTable in document.Tables)
                {
                    
                    int rowIndex = 0;
                    foreach (var aRow in aTable.Rows)
                    {
                        
                        var firstCell = aRow.Cells[0];

                        foreach (var aP in firstCell.Paragraphs)
                        {
                            string text = aP.Text;

                            var trimmed = data.Acquisitions.Keys.Select(p => p.Trim().ToLower().Replace(" ","")).ToList();
                            if (data.Acquisitions.ContainsKey(text.Trim().ToLower().Replace(" ", "")) || trimmed.Contains(text.Trim().ToLower().Replace(" ", "")) && !string.IsNullOrEmpty(text) && !string.IsNullOrWhiteSpace(text))
                            {
                                bool isSet = false;
                                int acquisitionValue = -1;
                                foreach(var a in data.Acquisitions)
                                    if(a.Key.Trim().ToLower().Replace(" ", "") == text.Trim().ToLower().Replace(" ", ""))
                                    {
                                        acquisitionValue = a.Value;
                                    }
                                //int acquisitionValue = data.Acquisitions[text];
                                for (int c = 1; c < aRow.Cells.Count; c++){
                                    if (acquisitionValue == c-1)
                                    {
                                        aRow.Cells[c].ReplaceText("{X}", "X");
                                        isSet = true;
                                    }                                        
                                    else
                                        aRow.Cells[c].ReplaceText("{X}", "");
                                }
                                if(!isSet)
                                {
                                    for (int c = 1; c < aRow.Cells.Count; c++)
                                    {
                                        aRow.Cells[c].FillColor = System.Drawing.Color.LightGray;
                                    }
                                }
                            }

                        }

                        var commentTest = firstCell.FindAll("commentaires}}");
                        if(commentTest.Count > 0)
                        {
                            firstCell.ReplaceText("&&place_pour_les_commentaires&&", "");
                            //firstCell.ReplaceText("\r", "");
                            //firstCell.ReplaceText("\n", "");

                        }
                    }
                }


                foreach(var aSubject in GlobalModel.Subjects) {
                    if (data.Comments.ContainsKey(aSubject.Name))
                        document.ReplaceText("{{" + aSubject.Name + "_commentaires" + "}}", ""); // data.Comments[aSubject.Name]);
                    else
                        document.ReplaceText("{{" + aSubject.Name + "_commentaires" + "}}","");
                }

                //if (document.fin FindUniqueByPattern(@"<[\w \=]{4,}>", RegexOptions.IgnoreCase).Count == _replacePatterns.Count)
                {
                    /*
                    // Do the replacement
                    for (int i = 0; i < _replacePatterns.Count; ++i)
                    {
                        document.ReplaceText("<(.*?)>", DocumentSample.ReplaceFunc, false, RegexOptions.IgnoreCase, null, new Formatting());
                    }
                    */
                    // Save this document to disk.
                    document.SaveAs(outputFile);

                }
            }
        }
    }
}
