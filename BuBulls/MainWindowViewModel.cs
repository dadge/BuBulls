using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using TestDocx;

namespace BuBulls
{
    public class MainWindowViewModel:  INotifyPropertyChanged
    {
        #region private properties
        private string _PathExcel;       
        private string _PathBaseTemplate;        
        private string _PathTemplate;        
        private string _PathOutpuFolder;

        private Visibility _ExcelVisible = Visibility.Hidden;
        private Visibility _TemplateVisible = Visibility.Hidden;
        private Visibility _GenerateVisible = Visibility.Hidden;
        private Visibility _EditTemplateButtonVisible = Visibility.Hidden;
        private Visibility _GeneratingVisible = Visibility.Hidden;
        private Visibility _ExcelLoadingVisible = Visibility.Hidden;
        private Visibility _GenerateButtonVisible = Visibility.Visible;


        private string _BaseTemplateCaption = "1. Dropper le template de base ici";
        private string _ExcelCaption = "2. Dropper l'Excel ici";
        private string _GenerateCaption = "4. Générer les bulletins";
        private string _TemplateCaption = "3. Adapter la mise en page.";

        private string _BaseTemplateHelp = "Le template de base est le modèle de base des bulletins. Il n'est constitué que d'une liste de tableau de matière et il est anonyme.";
        private string _ExcelHelp = "L'excel doit être l'excel qui vous a été fourni pour encoder la liste des élèves, des matières et compétences et qui permet l'encodage des résultats.";
        private string _GenerateHelp = "Tout est prêt, cliquez ici pour générer les bulletins.";
        private string _TemplateHelp = "Word s'est ouvert avec le template de bulletin adapté aux compétences.";

        private BackgroundWorker _excelBGLoader = null;
        private BackgroundWorker _genetingBGWorker = null;


        public class ExcelException: Exception
        {
            int Row { get; set; }
            int Column { get; set; }
            string Description { get; set; }

        }

        public MainWindowViewModel()
        {
            _excelBGLoader = new BackgroundWorker();
            _excelBGLoader.DoWork += _excelBGLoader_DoWork;
            _excelBGLoader.RunWorkerCompleted += _excelBGLoader_RunWorkerCompleted;

            _genetingBGWorker = new BackgroundWorker();
            _genetingBGWorker.DoWork += _genetingBGWorker_DoWork;
            _genetingBGWorker.RunWorkerCompleted += _genetingBGWorker_RunWorkerCompleted;
        }

      




        #endregion
        #region public properties
        public string PathExcel
        {
            get
            {
                return _PathExcel;
            }

            set
            {
                _PathExcel = value;
                NotifyPropertyChanged();
            }
        }
        public bool HasExcel
        {
            get { return !String.IsNullOrEmpty(PathExcel); }
        }

        public string PathBaseTemplate
        {
            get
            {
                return _PathBaseTemplate;
            }

            set
            {
                _PathBaseTemplate = value;
                NotifyPropertyChanged();
            }
        }
        public bool HasBaseTemplate
        {
            get { return !String.IsNullOrEmpty(PathBaseTemplate); }
        }

        public string PathTemplate
        {
            get
            {
                return _PathTemplate;
            }

            set
            {
                _PathTemplate = value;
                NotifyPropertyChanged();
            }
        }
        public bool HasTemplate
        {
            get { return !String.IsNullOrEmpty(PathTemplate); }
        }

        public string PathOutpuFolder
        {
            get
            {
                return _PathOutpuFolder;
            }

            set
            {
                _PathOutpuFolder = value;
                NotifyPropertyChanged();
            }
        }
        public bool HasOutpuFolder
        {
            get { return !String.IsNullOrEmpty(PathOutpuFolder); }
        }


        
        public string BaseTemplateCaption
        {
               get { return _BaseTemplateCaption;  }
            set
            {
                _BaseTemplateCaption = value;
                NotifyPropertyChanged();
            }
        }

        
        public string ExcelCaption
        {
            get { return _ExcelCaption; }
            set
            {
                _ExcelCaption = value;
                NotifyPropertyChanged();
            }
        }

        
        public string TemplateCaption
        {
            get { return _TemplateCaption; }
            set
            {
                _TemplateCaption = value;
                NotifyPropertyChanged();
            }
        }

        
        public string GenerateCaption
        {
            get { return _GenerateCaption; }
            set
            {
                _GenerateCaption = value;
                NotifyPropertyChanged();
            }
        }

        public Visibility ExcelVisible
        {
            get
            {
                return _ExcelVisible;
            }

            set
            {
                _ExcelVisible = value;
                NotifyPropertyChanged();
            }
        }

        public Visibility TemplateVisible
        {
            get
            {
                return _TemplateVisible;
            }

            set
            {
                _TemplateVisible = value;
                NotifyPropertyChanged();
            }
        }

        public Visibility GenerateVisible
        {
            get
            {
                return _GenerateVisible;
            }

            set
            {
                _GenerateVisible = value;
                NotifyPropertyChanged();
            }
        }

        public string TemplateHelp
        {
            get
            {
                return _TemplateHelp;
            }

            set
            {
                _TemplateHelp = value;
                NotifyPropertyChanged();
            }
        }

        public string GenerateHelp
        {
            get
            {
                return _GenerateHelp;
            }

            set
            {
                _GenerateHelp = value;
                NotifyPropertyChanged();
            }
        }

        public string ExcelHelp
        {
            get
            {
                return _ExcelHelp;
            }

            set
            {
                _ExcelHelp = value;
                NotifyPropertyChanged();
            }
        }

        public string BaseTemplateHelp
        {
            get
            {
                return _BaseTemplateHelp;
            }

            set
            {
                _BaseTemplateHelp = value;
                NotifyPropertyChanged();
            }
        }

        public Visibility EditTemplateButtonVisible
        {
            get
            {
                return _EditTemplateButtonVisible;
            }

            set
            {
                _EditTemplateButtonVisible = value;
                NotifyPropertyChanged();
            }
        }

        public Visibility GeneratingVisible
        {
            get
            {
                return _GeneratingVisible;
            }

            set
            {
                _GeneratingVisible = value;
                NotifyPropertyChanged();
            }
        }

        public Visibility ExcelLoadingVisible
        {
            get
            {
                return _ExcelLoadingVisible;
            }

            set
            {
                _ExcelLoadingVisible = value;
                NotifyPropertyChanged();
            }
        }

        public Visibility GenerateButtonVisible
        {
            get
            {
                return _GenerateButtonVisible;
            }

            set
            {
                _GenerateButtonVisible = value;
                NotifyPropertyChanged();
            }
        }

        #endregion

        #region business logic
        public void DropBaseTemplate(string file)
        {
            // validate file
            if(System.IO.Path.GetExtension(file) != ".docx")
            {
                System.Windows.MessageBox.Show("Le fichier doit être un fichier Word (.docx).", "Générateur de bulletins", MessageBoxButton.OK);
                return;
            }

            var isOpened = false;
            try
            {
                Stream s = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.None);

                s.Close();

                isOpened = false;
            }
            catch (Exception)
            {
                isOpened = true;
            }

            if (isOpened)
            {
                System.Windows.MessageBox.Show("Le fichier Word (template de base: \"" + System.IO.Path.GetFileName(file) +"\") doit être fermé pour que cela fonctionne!\r\nFermez le fichier word et recommencez.", "Générateur de bulletins", MessageBoxButton.OK);
                return;
            }

            this.PathBaseTemplate = file;
            NotifyPropertyChanged("HasBaseTemplate");

            // change caption
            this.BaseTemplateCaption = "Template de base Ok!";
            this.BaseTemplateHelp = "Fichier selectionné: " + System.IO.Path.GetFileName(file);
            
            // set excel visible
            this.ExcelVisible = Visibility.Visible;
        }

        public List<ExcelParser.ParsingError> Errors = null;
        List<ExcelParser.PersonalizedSchoolReportData> ParsedData = null;
        SchoolReportTemplate GlobalModel = null;
        public async void DropExcel(string file)
        {
            // validate file
            if (System.IO.Path.GetExtension(file) != ".xlsx" && System.IO.Path.GetExtension(file) != ".xlsm")
            {
                System.Windows.MessageBox.Show("Le fichier doit être un fichier Excel (.xslx) ou un fichier Excel avec macro (.xslm)", "Générateur de bulletins", MessageBoxButton.OK);
                return;
            }
            var isOpened = false;
            try
            {
                Stream s = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.None);

                s.Close();

                isOpened = false;
            }
            catch (Exception)
            {
                isOpened = true;
            }

            if (isOpened)
            {
                System.Windows.MessageBox.Show("Le fichier Excel (\"" + System.IO.Path.GetFileName(file) + "\") doit être fermé pour que cela fonctionne!\r\nFermez le fichier Excel et recommencez.", "Générateur de bulletins", MessageBoxButton.OK);
                return;
            }
            this.PathExcel = file;
            NotifyPropertyChanged("HasExcel");
            this.ExcelCaption = "";
            // change caption
            this.ExcelHelp =  "";
            // set excel visible
            this.ExcelLoadingVisible = Visibility.Visible;

            // Generates template file
            _excelBGLoader.RunWorkerAsync();                                    
        }

        private void _excelBGLoader_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if(e.Error != null)
            {
                System.Windows.MessageBox.Show("Désolé, une erreure est apparue dans le fichier. Notez le message suivant et transférer le au developpeur ainsi que le fichier concerné.\r\n" + e.Error.Message + "\r\n" + e.Error.StackTrace, "Générateur de bulletins", MessageBoxButton.OK);
            }
            else
            {
                if(Errors != null && Errors.Count > 0)
                {
                    this.ShowErrors();
                    this.ExcelLoadingVisible = Visibility.Hidden;
                    return;
                }
                
                this.ExcelCaption = "Excel OK!";
                // change caption
                this.ExcelHelp = "Fichier selectionné: " + System.IO.Path.GetFileName(PathExcel);
                // set excel visible
                this.TemplateVisible = Visibility.Visible;
                this.ExcelLoadingVisible = Visibility.Hidden;

                // open word to edit file
                OpenTemplate();
            }

        }

        private void _excelBGLoader_DoWork(object sender, DoWorkEventArgs e)
        {
            string intermediateTemplate = System.IO.Path.ChangeExtension(PathExcel, ".docx");
            GlobalModel = ExcelParser.FromExcel(this.PathExcel);
            GlobalTemplateGenerator.GenerateTemplateFile(PathBaseTemplate, GlobalModel, intermediateTemplate);
            var parseResult = ExcelParser.ParseAcquisitions(PathExcel);
            Errors = parseResult.Errors;
            ParsedData = parseResult.Data;
            if (Errors != null && Errors.Count > 0)
            {
                this.ExcelHelp +=  Errors.Count + " problème(s) dans le fichier Excel\r\nIl faut corriger et dropper ici à nouveau.\r\nCliquez ici pour revoir le détail des erreurs.";
                this.PathTemplate = intermediateTemplate;
            }
            else
            {
                this.ExcelHelp += "\r\n" + ParsedData.Count + " bulletins prêts à générer.";
                this.PathTemplate = intermediateTemplate;
            }
        }

        public void ShowErrors()
        {
            if (Errors != null && Errors.Count > 0)
            {
                MessageBox mb = new MessageBox();
                mb.Show(Errors);
                this.ExcelLoadingVisible = Visibility.Hidden;                
            }
        }

        public void OpenTemplate()
        {
            System.Diagnostics.Process word = new System.Diagnostics.Process();
            word.StartInfo.FileName =this.PathTemplate;
            word.Exited += Word_Exited;
            word.EnableRaisingEvents = true;
            word.Start();
        }

        private void Word_Exited(object sender, EventArgs e)
        {
            this.TemplateCaption = "Template vérifié !";
            this.TemplateHelp = "Fichier selectionné: " + System.IO.Path.GetFileName(PathTemplate) + "\r\n" + "Cliquez ici pour le réouvrir";
            this.GenerateVisible = Visibility.Visible;
            this.EditTemplateButtonVisible = Visibility.Visible;
            

        }

        private string findOutputFolderName(string folder, int i = 0)
        {
            string target = null;
            if(i > 0)
                target = System.IO.Path.Combine(folder, "bulletins_"+i).ToString();
            else
                target = System.IO.Path.Combine(folder, "bulletins");

            if (System.IO.Directory.Exists(target))
                return findOutputFolderName(folder, i + 1);
            else
            {
                System.IO.Directory.CreateDirectory(target);
                return target;
            }
                
        }
        private string outputFolder;
        public void GeneratesTemplate()
        {
            this.GeneratingVisible = Visibility.Visible;
            GenerateButtonVisible = Visibility.Hidden;
            this.GenerateCaption = "";
            this.GenerateHelp = "";
            _genetingBGWorker.RunWorkerAsync();            
        }


        private void _genetingBGWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.GeneratingVisible = Visibility.Hidden;
            GenerateButtonVisible = Visibility.Visible;

            this.GenerateCaption = "Terminé";
            this.GenerateHelp = "Les bulletins ont été générés dans " + outputFolder;
        }

        private void _genetingBGWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            string basePath = System.IO.Path.GetDirectoryName(PathExcel);
            outputFolder = findOutputFolderName(basePath);
            foreach (var aReport in ParsedData)
                PersonalizedSchoolReportGenerator.GeneratePersonalizedReport(aReport, GlobalModel, outputFolder, PathTemplate);
            System.Diagnostics.Process.Start("explorer.exe", outputFolder);
        }

        #endregion


        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            var handler = PropertyChanged;
            if (handler != null)
                handler(this, new PropertyChangedEventArgs(propertyName));
        }

    }
}
