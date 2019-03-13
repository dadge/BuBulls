using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace BuBulls
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        /*
         *  Console.WriteLine("Hello World!");

            string ExcelNoteBook = "/Users/fcasado/Documents/Workbook1.xlsx";
            string rootTemplate = "/Users/fcasado/Documents/template.docx";
            string intermediateTemplate = "/Users/fcasado/Documents/intermediateTemplate.docx";
            string outputFolder = "/Users/fcasado/Documents/bulletins";

            if (!System.IO.Directory.Exists(outputFolder))
                System.IO.Directory.CreateDirectory(outputFolder);


            var GlobalModel = ExcelParser.FromExcel(ExcelNoteBook);
            GlobalTemplateGenerator.GenerateTemplateFile(rootTemplate,GlobalModel,intermediateTemplate);
            var personalizedAcquisitions = ExcelParser.ParseAcquisitions(ExcelNoteBook);

            foreach (var aReport in personalizedAcquisitions)
                PersonalizedSchoolReportGenerator.GeneratePersonalizedReport(aReport, GlobalModel, outputFolder, intermediateTemplate);

         */
    }
}
