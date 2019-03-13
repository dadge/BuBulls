using System;
using System.Linq;
using Xceed.Words.NET;

namespace TestDocx
{
    public class GlobalTemplateGenerator
    {
        public GlobalTemplateGenerator()
        {
        }

        public static void GenerateTemplateFile(string rootTemplatePath, SchoolReportTemplate GlobalModel, string intermediateTemplatePath)
        {
            string blablaComment = "&&place_pour_les_commentaires&&";
            using (DocX document = DocX.Load(rootTemplatePath))
            {

                // Check if all the replace patterns are used in the loaded document.
                foreach (var aTable in document.Tables)
                {
                    Console.WriteLine("table " + aTable.Index + " with " + aTable.RowCount + "x" + aTable.ColumnCount);
                    int rowIndex = 0;
                    foreach (var aRow in aTable.Rows)
                    {
                        Console.WriteLine("row at " + aTable.Index + " with " + aRow.ColumnCount);
                        var firstCell = aRow.Cells[0];

                        foreach (var aP in firstCell.Paragraphs)
                        {
                            string text = aP.Text;


                            if (text.StartsWith("{{") && text.EndsWith("}}"))
                            {
                                string tText = aP.Text.Replace("{{", "").Replace("}}", "");
                                var subject = GlobalModel.Subjects.SingleOrDefault(p => p.Name == tText);
                                if (subject != null)
                                {
                                    subject.Acquisitions.Reverse();
                                    foreach (var aC in subject.Acquisitions)
                                    {
                                        var nRow = aTable.InsertRow(aRow, rowIndex);
                                        // nRow.InsertParagraph(0, "hey bitchy", false);
                                        nRow.ReplaceText(text, aC);
                                        /*
                                        if(aRow.ColumnCount-1 > 0)
                                            for (int i = 1; i < aRow.ColumnCount;i++){
                                                if(aRow.Cells[i].Paragraphs != null && aRow.Cells[i].Paragraphs.Count() > 0)
                                                {
                                                   // foreach (var pa in aRow.Cells[i].Paragraphs)
                                                   //     aRow.Cells[i].RemoveParagraph(pa);
                                                }
                                                aRow.Cells[i].InsertParagraph(0, "{{X}}", false);
                                            }*/
                                    }
                                    aRow.Remove();
                                }

                                subject = null;
                                subject = GlobalModel.Subjects.SingleOrDefault(p => p.Name + "_commentaires" == tText);
                                if (subject != null)
                                    aP.ReplaceText("{{" + subject.Name + "_commentaires" + "}}", "{{" + subject.Name + "_commentaires" + "}}\r\n"  + blablaComment + "\r\n" + blablaComment + "\r\n" + blablaComment + "\r\n" + blablaComment);
                            }
                        }
                        rowIndex++;

                    }
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
                    document.SaveAs(intermediateTemplatePath);

                }
            }
        }
    }
}
