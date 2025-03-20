using Dalton_Trapper.Model.Projektering_tab;
using Dalton_Trapper.Utilities;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Input;
using ImportExport = Dalton_Trapper.Model.Projektering_tab.ImportExport;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Windows.Controls;
using Microsoft.Office.Interop.Excel;
using Drawing = Dalton_Trapper.Model.Projektering_tab.Drawing;

namespace Dalton_Trapper.ViewModel
{
    class ProjekteringVM : Utilities.ViewModelBase
    {
        public ProjekteringVM()
        {
            // Initialize project info class
            ProjectInfo = new ProjectInfo();

            // Data lists
            KonsekvensklasseListe = new ObservableCollection<string> { "", "CC2", "CC3" };
            MiljøklasseListe = new ObservableCollection<string> { "", "Passiv", "Moderat", "Aggressiv", "Ekstra Aggressiv" };
            SelectedDrawings = new ObservableCollection<Drawing>();
            SelectedCalculations = new ObservableCollection<Calculation>();

            // Commands           
            OpenDrawingCommand = new RelayCommand(OpenDrawing);
            OpenSelectedDrawingsCommand = new RelayCommand(_ => ExecuteOpenSelectedDrawings(), _ => CanExecuteOpenSelectedDrawings());
            OpenSelectedCalculationsCommand = new RelayCommand(_ => ExecuteOpenSelectedCalculations(), _ => CanExecuteOpenSelectedCalculations());
            CommandBeregningsark = new RelayCommand(OpenBeregningsark);
            CommandDorneBrand = new RelayCommand(OpenDorneBrand);
            CommandCRHDokumentation = new RelayCommand(OpenCRHDokumentation);
            CommandIngeniørEkstern = new RelayCommand(_ => OpenIngeniørEkstern(), _ => ProjectNumberEntered());
            CommandTjekliste = new RelayCommand(_ => OpenCreateTjekliste(), _ => ProjectNumberEntered());
            CommandBeregningsmappe = new RelayCommand(_ => OpenCalculationFolder(), _ => ProjectNumberEntered());
            CommandTegningsmappe = new RelayCommand(_ => OpenDrawingFolder(), _ => ProjectNumberEntered());
            AddFoldersCommand = new RelayCommand(_ => AddFolders(), _ => CanAddFolders());
            GoUpDrawingCommand = new RelayCommand(_ => GoUpDrawings(), _ => !string.IsNullOrEmpty(CurrentDrawingFolder));
            GoUpCalculationCommand = new RelayCommand(_ => GoUpCalculations(), _ => !string.IsNullOrEmpty(CurrentCalculationFolder));
            CopyDrawingsCommand = new RelayCommand(_ => CopySelectedFiles(), _ => CanCopyFiles());

            OpretLigeløbCommand = new RelayCommand(_ => OpretBeregning("Ligeløb"), _ => CanAddFiles(Ligeløb));
            OpretKnækløbCommand = new RelayCommand(_ => OpretBeregning("Knækløb"), _ => CanAddFiles(Knækløb));
            OpretReposCommand = new RelayCommand(_ => OpretBeregning("Repos"), _ => CanAddFiles(Repos));
            OpretSvingløbCommand = new RelayCommand(_ => OpretBeregning("Svingløb"), _ => CanAddFiles(Svingløb));
            OpretDetaljeCommand = new RelayCommand(_ => OpretBeregning("Detalje"), _ => CanAddFiles(Detalje));
            OpretPladeCommand = new RelayCommand(_ => OpretBeregning("Plade"), _ => CanAddFiles(Plade));
        }

        #region Commands
        public ICommand OpenDrawingCommand { get; set; }
        public ICommand OpenSelectedDrawingsCommand { get; set; }
        public ICommand OpenSelectedCalculationsCommand { get; set; }
        public ICommand CommandBeregningsark { get; set; }
        public ICommand CommandDorneBrand { get; set; }
        public ICommand CommandCRHDokumentation { get; set; }
        public ICommand CommandIngeniørEkstern { get; set; }
        public ICommand CommandTjekliste { get; set; }
        public ICommand CommandBeregningsmappe { get; set; }
        public ICommand CommandTegningsmappe { get; set; }
        public ICommand GoUpDrawingCommand { get; set; }
        public ICommand GoUpCalculationCommand { get; set; }
        public ICommand AddFoldersCommand { get; set; }
        public ICommand OpretLigeløbCommand { get; set; }
        public ICommand OpretKnækløbCommand { get; set; }
        public ICommand OpretReposCommand { get; set; }
        public ICommand OpretSvingløbCommand { get; set; }
        public ICommand OpretDetaljeCommand { get; set; }
        public ICommand OpretPladeCommand { get; set; }
        public ICommand CopyDrawingsCommand { get; set; }

        private void ExecuteOpenSelectedDrawings()
        {
            foreach (var item in SelectedDrawings)
            {
                    OpenDrawing(item);
            }
        }

        private bool CanExecuteOpenSelectedDrawings() => SelectedDrawings?.Count > 0;

        private void ExecuteOpenSelectedCalculations()
        {
            foreach (var item in SelectedCalculations)
            {
                OpenDrawing(item);
            }
        }

        private bool CanExecuteOpenSelectedCalculations() => SelectedCalculations?.Count > 0;

        private void OpenDrawing(object parameter)
        {
            if (parameter is Drawing file)
            {
                if (file.IsFolder)
                {
                    LoadFolderContents(file.FilePath, "Drawing");
                }
                else
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = file.FilePath,
                        UseShellExecute = true
                    });
                }
            }

            else if (parameter is Calculation file2)
            {
                if (file2.IsFolder)
                {
                    LoadFolderContents(file2.FilePath, "Calculation");
                }
                else
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = file2.FilePath,
                        UseShellExecute = true
                    });
                }

            }
        }

        private void OpenBeregningsark(object parameter)
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
            {
                FileName = "N:\\Opslag\\Statik\\Dalton statikark",
                UseShellExecute = true,
                Verb = "open"
            });
        }

        private void OpenDorneBrand(object parameter)
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
            {
                FileName = "N:\\Projektering\\Tilst\\Trappedorne_Brand",
                UseShellExecute = true,
                Verb = "open"
            });
        }

        private void OpenCRHDokumentation(object parameter)
        {
            string user = Environment.UserName.ToLower();

            try
            {
                string FileName = $"C:\\Users\\{user}\\Desktop\\CRHDokumentation.appref-ms";

                // CRH Dokumentation ligger på skrivebordet, men stien dertil er forskellig alt efter om OneDrive gemmer skrivebordets indhold eller ej.
                if (File.Exists(FileName)) // Ikke i OneDrive-mappen.
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                    {
                        FileName = FileName,
                        UseShellExecute = true,
                        Verb = "open"
                    });
                }

                //I OneDrive-mappen hvor skrivebordet er kaldet Desktop
                else if (File.Exists($"C:\\Users\\{user}\\OneDrive - CRH\\Desktop\\CRHDokumentation.appref-ms"))
                    {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                    {
                        FileName = $"C:\\Users\\{user}\\OneDrive - CRH\\Desktop\\CRHDokumentation.appref-ms",
                        UseShellExecute = true,
                        Verb = "open"
                    });
                }

                else //I OneDrive-mappen hvor skrivebordet er kaldet Skrivebord
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                    {
                        FileName = $"C:\\Users\\{user}\\OneDrive - CRH\\Skrivebord\\CRHDokumentation.appref-ms",
                        UseShellExecute = true,
                        Verb = "open"
                    });
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Kan ikke finde åbne CRHDokumentation.\n" +
                                e.Message, "Fejl!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenIngeniørEkstern()
        {
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                {
                    FileName = $"S:\\{ProjectInfo.ProjectNumber.Substring(0, 3)}\\{ProjectInfo.ProjectNumber}\\09.Beregninger\\Ingeniør ekstern",
                    UseShellExecute = true,
                    Verb = "open"
                });
            }
            catch (Exception e)
            {
                MessageBox.Show("Kan ikke finde mappen: \"Ingeniør ekstern\". Det kan skyldes en gammel mappestruktur på sagen.\n\n" +
                                e.Message, "Fejl!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenCreateTjekliste()
        {
            string file = $"S:\\{ProjectInfo.ProjectNumber.Substring(0, 3)}\\{ProjectInfo.ProjectNumber}\\09.Beregninger\\Elementstatik\\15 Trapper-Reposer og Skakte\\Tjekliste.docx";

            if (!File.Exists(file))
            {
                File.Copy("Model\\Projektering tab\\Tjekliste.docx", file);
            }

            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
            {
                FileName = file,
                UseShellExecute = true,
                Verb = "open"
            });
        }

        private void OpenCalculationFolder()
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
            {
                FileName = CurrentCalculationFolder,
                UseShellExecute = true,
                Verb = "open"
            });
        }

        private void OpenDrawingFolder()
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
            {
                FileName = CurrentDrawingFolder,
                UseShellExecute = true,
                Verb = "open"
            });
        }

        private bool ProjectNumberEntered()
        {
            return !string.IsNullOrWhiteSpace(ProjectInfo.ProjectNumber);
        }

        private void GoUpDrawings()
        {
            if (!string.IsNullOrEmpty(CurrentDrawingFolder))
            {
                string parentFolder = Path.GetDirectoryName(CurrentDrawingFolder);
                if (!string.IsNullOrEmpty(parentFolder) && Directory.Exists(parentFolder))
                {
                    LoadFolderContents(parentFolder, "Drawing");
                }
                else
                {
                    MessageBox.Show("No parent folder exists.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
        }

        private void GoUpCalculations()
        {
            if (!string.IsNullOrEmpty(CurrentCalculationFolder))
            {
                string parentFolder = Path.GetDirectoryName(CurrentCalculationFolder);
                if (!string.IsNullOrEmpty(parentFolder) && Directory.Exists(parentFolder))
                {
                    LoadFolderContents(parentFolder, "Calculation");
                }
                else
                {
                    MessageBox.Show("No parent folder exists.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
        }

        private void AddFolders()
        {
            string calculation_Path = $"S:\\{ProjectInfo.ProjectNumber.Substring(0, 3)}\\{ProjectInfo.ProjectNumber}\\09.Beregninger\\Elementstatik\\15 Trapper-Reposer og Skakte";

            if (string.IsNullOrEmpty(calculation_Path) || !Directory.Exists(calculation_Path))
            {
                MessageBox.Show("Ugyldig mappe sti. Kan ikke oprette mapper.", "Fejl!", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var folderNames = new[] { "01 Styrkeberegninger", "02 Rettede tegninger", "03 Statisk dokumentation", "04 Diverse" };
            bool allFoldersExists = folderNames.All(folderName => Directory.Exists(Path.Combine(calculation_Path, folderName)));

            foreach (var folderName in folderNames)
            {
                string newFolderPath = Path.Combine(calculation_Path, folderName);
                try
                {
                    if (!Directory.Exists(newFolderPath))
                    {
                        Directory.CreateDirectory(newFolderPath);
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show($"Kunne ikke oprette mapper '{folderName}': {e.Message}", "Fejl!", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

            LoadFolderContents(calculation_Path, "Calculation");
        }

        private bool CanAddFolders()
        {
            if (string.IsNullOrEmpty(ProjectInfo.ProjectNumber))
            {
                return false;
            }

            string calculation_Path = $"S:\\{ProjectInfo.ProjectNumber.Substring(0, 3)}\\{ProjectInfo.ProjectNumber}\\09.Beregninger\\Elementstatik\\15 Trapper-Reposer og Skakte";

            if (string.IsNullOrEmpty(calculation_Path) || !Directory.Exists(calculation_Path))
            {
                return false;
            }

            var folderNames = new[] { "01 Styrkeberegninger", "02 Rettede tegninger", "03 Statisk dokumentation", "04 Diverse" };
            return folderNames.Any(folderName => !Directory.Exists(Path.Combine(calculation_Path, folderName)));
        }

        private bool CanAddFiles(string type)
        {
            if (string.IsNullOrEmpty(ProjectInfo.ProjectNumber) || string.IsNullOrEmpty(type))
            {
                return false;
            }

            string calculation_Path = $"S:\\{ProjectInfo.ProjectNumber.Substring(0, 3)}\\{ProjectInfo.ProjectNumber}\\09.Beregninger\\Elementstatik\\15 Trapper-Reposer og Skakte";

            if (string.IsNullOrEmpty(calculation_Path) || !Directory.Exists(calculation_Path))
            {
                return false;
            }

            return true;
        }

        private void OpretBeregning(string type)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            string environmentClass = "";
            if (ProjectInfo.Miljøklasse == "Moderat")
            {
                environmentClass = "M";
            }
            else if (ProjectInfo.Miljøklasse == "Aggressiv")
            {
                environmentClass = "A";
            }
            else if (ProjectInfo.Miljøklasse == "Ekstra Aggressiv")
            {
                environmentClass = "E";
            }
            else
            {
                environmentClass = "P";
            }

            if (type == "Ligeløb")
            {
                string sheet = "N:\\Opslag\\Statik\\Dalton statikark\\LobEC_190701.xlsm";

                string newFile = System.IO.Path.Combine(CurrentCalculationFolder, $"{Ligeløb}.xlsm");

                if (File.Exists(newFile))
                {
                    MessageBox.Show($"Filen '{Ligeløb}.xlsm' eksisterer i forvejen. Vælg et andet navn.", "Fejl!", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                File.Copy(sheet, newFile);

                // Updates the sheet with projectinformation
                excelApp = new Excel.Application();
                workbook = excelApp.Workbooks.Open(newFile);
                worksheet = (Excel.Worksheet)workbook.Sheets[1];
                worksheet.Cells[2, 3] = ProjectInfo.ProjectName;
                worksheet.Cells[6, 3] = Ligeløb;
                worksheet.Cells[9, 9] = ProjectInfo.Konsekvensklasse == "" ? 2 : ProjectInfo.Konsekvensklasse.Substring(ProjectInfo.Konsekvensklasse.Length - 1, 1); // Consequence class
                worksheet.Cells[35, 7] = ProjectInfo.Liveload == "5" ? "C" : "A"; // Liveload
                worksheet.Cells[11, 9] = environmentClass;
                worksheet.Cells[2, 10] = ProjectInfo.ProjectNumber;

                Ligeløb = "";
            }
            else if (type == "Repos")
            {
                string sheet = "Model\\Projektering tab\\Beregningsark\\reposEC_v2010141 rev 17012025.xlsm";

                string newFile = System.IO.Path.Combine(CurrentCalculationFolder, $"{Repos}.xlsm");

                if (File.Exists(newFile))
                {
                    MessageBox.Show($"Filen '{Repos}.xlsm' eksisterer i forvejen. Vælg et andet navn.", "Fejl!", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                File.Copy(sheet, newFile);

                // Updates the sheet with projectinformation
                excelApp = new Excel.Application();
                workbook = excelApp.Workbooks.Open(newFile);
                worksheet = (Excel.Worksheet)workbook.Sheets[1];
                worksheet.Cells[2, 3] = ProjectInfo.ProjectName;
                worksheet.Cells[6, 3] = Repos;
                worksheet.Cells[9, 9] = ProjectInfo.Konsekvensklasse == "" ? 2 : ProjectInfo.Konsekvensklasse.Substring(ProjectInfo.Konsekvensklasse.Length - 1, 1); // Consequence class
                worksheet.Cells[44, 7] = ProjectInfo.Liveload == "5" ? "C" : "A"; // Liveload
                worksheet.Cells[11, 9] = environmentClass;
                worksheet.Cells[2, 10] = ProjectInfo.ProjectNumber;

                Repos = "";
            }
            else if (type == "Knækløb")
            {
                string sheet = "Model\\Projektering tab\\Beregningsark\\LobrepEC_190701 rev 17012025.xlsm";

                string newFile = System.IO.Path.Combine(CurrentCalculationFolder, $"{Knækløb}.xlsm");

                if (File.Exists(newFile))
                {
                    MessageBox.Show($"Filen '{Knækløb}.xlsm' eksisterer i forvejen. Vælg et andet navn.", "Fejl!", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                File.Copy(sheet, newFile);

                // Updates the sheet with projectinformation
                excelApp = new Excel.Application();
                workbook = excelApp.Workbooks.Open(newFile);
                worksheet = (Excel.Worksheet)workbook.Sheets[1];
                worksheet.Cells[2, 3] = ProjectInfo.ProjectName;
                worksheet.Cells[6, 3] = Knækløb;
                worksheet.Cells[9, 9] = ProjectInfo.Konsekvensklasse == "" ? 2 : ProjectInfo.Konsekvensklasse.Substring(ProjectInfo.Konsekvensklasse.Length - 1, 1); // Consequence class
                worksheet.Cells[50, 7] = ProjectInfo.Liveload == "5" ? "C" : "A"; // Liveload
                worksheet.Cells[11, 9] = environmentClass;
                worksheet.Cells[2, 10] = ProjectInfo.ProjectNumber;

                Knækløb = "";
            }
            else if (type == "Svingløb")
            {
                string sheet = "Model\\Projektering tab\\Beregningsark\\Svingløb_v28-10-2022.xlsm";

                string newFile = System.IO.Path.Combine(CurrentCalculationFolder, $"{Svingløb}.xlsm");

                if (File.Exists(newFile))
                {
                    MessageBox.Show($"Filen '{Svingløb}.xlsm' eksisterer i forvejen. Vælg et andet navn.", "Fejl!", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                File.Copy(sheet, newFile);

                // Updates the sheet with projectinformation
                excelApp = new Excel.Application();
                workbook = excelApp.Workbooks.Open(newFile);
                worksheet = (Excel.Worksheet)workbook.Sheets[1];
                worksheet.Cells[2, 4] = ProjectInfo.ProjectName;
                worksheet.Cells[6, 4] = Svingløb;
                worksheet.Cells[77, 7] = ProjectInfo.Konsekvensklasse == "" ? 2 : ProjectInfo.Konsekvensklasse.Substring(ProjectInfo.Konsekvensklasse.Length - 1, 1); // Consequence class
                worksheet.Cells[102, 11] = ProjectInfo.Liveload == "5" ? "Kategori C" : "Kategori A"; // Liveload
                worksheet.Cells[79, 7] = environmentClass;
                worksheet.Cells[2, 15] = ProjectInfo.ProjectNumber;

                Svingløb = "";
            }
            else if (type == "Detalje")
            {
                string sheet = "N:\\Opslag\\Statik\\Dalton statikark\\TrdetEC_191017 rev 09012023.xlsm";

                string newFile = System.IO.Path.Combine(CurrentCalculationFolder, $"{Detalje}.xlsm");

                if (File.Exists(newFile))
                {
                    MessageBox.Show($"Filen '{Detalje}.xlsm' eksisterer i forvejen. Vælg et andet navn.", "Fejl!", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                File.Copy(sheet, newFile);

                // Updates the sheet with projectinformation
                excelApp = new Excel.Application();
                workbook = excelApp.Workbooks.Open(newFile);
                worksheet = (Excel.Worksheet)workbook.Sheets[1];
                worksheet.Cells[2, 3] = ProjectInfo.ProjectName;
                worksheet.Cells[6, 3] = Detalje;
                worksheet.Cells[9, 9] = ProjectInfo.Konsekvensklasse == "" ? 2 : ProjectInfo.Konsekvensklasse.Substring(ProjectInfo.Konsekvensklasse.Length - 1, 1); // Consequence class
                worksheet.Cells[11, 9] = environmentClass;
                worksheet.Cells[2, 10] = ProjectInfo.ProjectNumber;

                Detalje = "";
            }
            else if (type == "Plade")
            {
                string sheet = "N:\\Opslag\\Statik\\Dalton statikark\\RApladEC_190701.xlsm";

                string newFile = System.IO.Path.Combine(CurrentCalculationFolder, $"{Plade}.xlsm");

                if (File.Exists(newFile))
                {
                    MessageBox.Show($"Filen '{Plade}.xlsm' eksisterer i forvejen. Vælg et andet navn.", "Fejl!", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                File.Copy(sheet, newFile);

                // Updates the sheet with projectinformation
                excelApp = new Excel.Application();
                workbook = excelApp.Workbooks.Open(newFile);
                worksheet = (Excel.Worksheet)workbook.Sheets[1];
                worksheet.Cells[2, 3] = ProjectInfo.ProjectName;
                worksheet.Cells[6, 3] = Plade;
                worksheet.Cells[29, 9] = ProjectInfo.Konsekvensklasse == "" ? 2 : ProjectInfo.Konsekvensklasse.Substring(ProjectInfo.Konsekvensklasse.Length - 1, 1); // Consequence class
                worksheet.Cells[52, 7] = ProjectInfo.Liveload == "5" ? "C" : "A"; // Liveload
                worksheet.Cells[31, 9] = environmentClass;
                worksheet.Cells[2, 10] = ProjectInfo.ProjectNumber;

                Plade = "";
            }

            // Save, close and release objects
            workbook.Save();
            workbook.Close();
            excelApp.Quit();
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);

            LoadFolderContents(CurrentCalculationFolder, "Calculation");
        }
        #endregion

        #region Collections for user input
        public ObservableCollection<string> KonsekvensklasseListe { get; set; }

        public ObservableCollection<string> MiljøklasseListe { get; set; }


        private ProjectInfo _projectInfo;

        public ProjectInfo ProjectInfo
        {
            get => _projectInfo;
            set
            {
                if (_projectInfo != value)
                {
                    _projectInfo = value;
                    OnPropertyChanged(nameof(ProjectInfo));
                    SubscribeToPropertyChanges();
                }
            }
        }

        private ObservableCollection<Drawing> drawings;
        public ObservableCollection<Drawing> Drawings
        {
            get { return drawings ?? (drawings = new ObservableCollection<Drawing>()); }

            set { drawings = value; }
        }

        private ObservableCollection<Drawing> selectedDrawings;
        public ObservableCollection<Drawing> SelectedDrawings
        {
            get { return selectedDrawings ?? (selectedDrawings = new ObservableCollection<Drawing>()); }

            set { selectedDrawings = value; }
        }

        private ObservableCollection<Calculation> selectedCalculations;
        public ObservableCollection<Calculation> SelectedCalculations
        {
            get { return selectedCalculations ?? (selectedCalculations = new ObservableCollection<Calculation>()); }

            set { selectedCalculations = value; }
        }

        private ObservableCollection<Calculation> calculations;
        public ObservableCollection<Calculation> Calculations
        {
            get { return calculations ?? (calculations = new ObservableCollection<Calculation>()); }

            set { calculations = value; }
        }

        private String _currentDrawingFolder;
        public String CurrentDrawingFolder
        {
            get => _currentDrawingFolder;
            set
            {
                if (_currentDrawingFolder != value)
                {
                    _currentDrawingFolder = value;
                    OnPropertyChanged(nameof(CurrentDrawingFolder));
                }
            }
        }

        private String _currentCalculationFolder;
        public String CurrentCalculationFolder
        {
            get => _currentCalculationFolder;
            set
            {
                if (_currentCalculationFolder != value)
                {
                    _currentCalculationFolder = value;
                    OnPropertyChanged(nameof(CurrentCalculationFolder));
                    CommandManager.InvalidateRequerySuggested();
                }
            }
        }

        private String _ligeløb;
        public String Ligeløb
        {
            get => _ligeløb;
            set
            {
                if (_ligeløb != value)
                {
                    _ligeløb = value;
                    OnPropertyChanged(nameof(Ligeløb));
                }
            }
        }

        private String _knækløb;
        public String Knækløb
        {
            get => _knækløb;
            set
            {
                if (_knækløb != value)
                {
                    _knækløb = value;
                    OnPropertyChanged(nameof(Knækløb));
                }
            }
        }

        private String _repos;
        public String Repos
        {
            get => _repos;
            set
            {
                if (_repos != value)
                {
                    _repos = value;
                    OnPropertyChanged(nameof(Repos));
                }
            }
        }

        private String _svingløb;
        public String Svingløb
        {
            get => _svingløb;
            set
            {
                if (_svingløb != value)
                {
                    _svingløb = value;
                    OnPropertyChanged(nameof(Svingløb));
                }
            }
        }

        private String _detalje;
        public String Detalje
        {
            get => _detalje;
            set
            {
                if (_detalje != value)
                {
                    _detalje = value;
                    OnPropertyChanged(nameof(Detalje));
                }
            }
        }
        private String _plade;
        public String Plade
        {
            get => _plade;
            set
            {
                if (_plade != value)
                {
                    _plade = value;
                    OnPropertyChanged(nameof(Plade));
                }
            }
        }
        #endregion

        #region Save and import functions
        private void SubscribeToPropertyChanges()
        {
            ProjectInfo.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(ProjectInfo.ProjectNumber))
                {
                    LoadProjectInfo();
                }
                else
                {
                    SaveProjectInfo();
                }
            };
        }

        private void SaveProjectInfo()
        {
            if (!string.IsNullOrWhiteSpace(ProjectInfo.ProjectNumber) &&
                !string.IsNullOrWhiteSpace(ProjectInfo.ProjectName))
            {
                string folderPath = $"S:\\{ProjectInfo.ProjectNumber.Substring(0, 3)}\\{ProjectInfo.ProjectNumber}\\09.Beregninger\\Elementstatik\\15 Trapper-Reposer og Skakte";
                string fileName = $"{ProjectInfo.ProjectNumber} TrappeApp Data.json";
                string fullPath = Path.Combine(folderPath, fileName);

                var data = new ProjectInfo
                {
                    ProjectNumber = ProjectInfo.ProjectNumber,
                    ProjectName = ProjectInfo.ProjectName,
                    Konsekvensklasse = ProjectInfo.Konsekvensklasse,
                    Miljøklasse = ProjectInfo.Miljøklasse,
                    Liveload = ProjectInfo.Liveload,
                };

                ImportExport.SaveToJson(data, fullPath);
            }
        }

        private bool _isLoadingData = false;

        private async void LoadProjectInfo()
        {
            if (_isLoadingData) return;

            _isLoadingData = true;

            string folderPath = $"S:\\{ProjectInfo.ProjectNumber.Substring(0, 3)}\\{ProjectInfo.ProjectNumber}\\09.Beregninger\\Elementstatik\\15 Trapper-Reposer og Skakte";
            string fileName = $"{ProjectInfo.ProjectNumber} TrappeApp Data.json";
            string fullPath = Path.Combine(folderPath, fileName);
            string drawingFolderPath = $"S:\\{ProjectInfo.ProjectNumber.Substring(0, 3)}\\{ProjectInfo.ProjectNumber}\\07.Tegninger\\CRH-Projekt\\15 Trapper-Reposer og Skakte\\4 Færdige PDF tegninger";
            string drawingFolderPath2 = $"S:\\{ProjectInfo.ProjectNumber.Substring(0, 3)}\\{ProjectInfo.ProjectNumber}\\07.Tegninger\\1 - CRH-Projekt\\15 Trapper-Reposer og Skakte\\4 Færdige PDF tegninger";
            string calculationFolderPath = $"S:\\{ProjectInfo.ProjectNumber.Substring(0, 3)}\\{ProjectInfo.ProjectNumber}\\09.Beregninger\\Elementstatik\\15 Trapper-Reposer og Skakte\\01 Styrkeberegninger";
            string calculationFolderPath2 = $"S:\\{ProjectInfo.ProjectNumber.Substring(0, 3)}\\{ProjectInfo.ProjectNumber}\\09.Beregninger\\Elementstatik\\15 Trapper-Reposer og Skakte";

            CurrentDrawingFolder = drawingFolderPath;
            CurrentCalculationFolder = calculationFolderPath;

            if (!string.IsNullOrWhiteSpace(ProjectInfo.ProjectNumber))
            {
                // Load data if anything is saved
                if (File.Exists(fullPath))
                {
                    Drawings.Clear();
                    Calculations.Clear();
                    var data = await Task.Run(() => ImportExport.LoadFromJson(fullPath));

                    // Assign project info to app
                    ProjectInfo.ProjectNumber = data.ProjectNumber;
                    ProjectInfo.ProjectName = data.ProjectName;
                    ProjectInfo.Konsekvensklasse = data.Konsekvensklasse;
                    ProjectInfo.Miljøklasse = data.Miljøklasse;
                    ProjectInfo.Liveload = data.Liveload;
                }
                else
                {
                    ProjectInfo.ProjectName = "";
                    ProjectInfo.Konsekvensklasse = "";
                    ProjectInfo.Miljøklasse = "";
                    ProjectInfo.Liveload = "";
                    Drawings.Clear();
                    Calculations.Clear();
                }
                _isLoadingData = false;
            }

            // Load drawings //
            try
            {
                string[] directories = Directory.GetDirectories(drawingFolderPath, "*", SearchOption.TopDirectoryOnly);
                foreach (string directory in directories)
                {
                    var folder = new Drawing { FileName = Path.GetFileName(directory), FilePath = directory, IsFolder = true };
                    Drawings.Add(folder);
                }

                string[] pdfFiles = Directory.GetFiles(drawingFolderPath, "*pdf", SearchOption.TopDirectoryOnly);
                foreach (string file in pdfFiles)
                {
                    var drawing = new Drawing { FileName = Path.GetFileName(file), FilePath = file, FileExtension = Path.GetExtension(file), IsFolder = false };
                    Drawings.Add(drawing);
                }
            }

            catch (Exception)
            {
                CurrentDrawingFolder = drawingFolderPath2;
                try
                {
                    string[] directories = Directory.GetDirectories(drawingFolderPath2, "*", SearchOption.TopDirectoryOnly);
                    foreach (string directory in directories)
                    {
                        var folder = new Drawing { FileName = Path.GetFileName(directory), FilePath = directory, IsFolder = true };
                        Drawings.Add(folder);
                    }

                    string[] pdfFiles = Directory.GetFiles(drawingFolderPath2, "*pdf", SearchOption.TopDirectoryOnly);
                    foreach (string file in pdfFiles)
                    {
                        var drawing = new Drawing { FileName = Path.GetFileName(file), FilePath = file, FileExtension = Path.GetExtension(file), IsFolder = false };
                        Drawings.Add(drawing);
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("Kan ikke finde mappen: \"4 Færdige PDF tegninger\". Det kan skyldes en gammel mappestruktur på sagen.\n\n" + e.Message, 
                        "Fejl!", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }


            // Load Calculations //
            try // Styrkeberegninger folder
            {
                // Load folders
                string[] directories = Directory.GetDirectories(calculationFolderPath, "*", SearchOption.TopDirectoryOnly);
                foreach (string directory in directories)
                {
                    var folder = new Calculation
                    {
                        FileName = Path.GetFileName(directory),
                        FilePath = directory,
                        IsFolder = true
                    };
                    Calculations.Add(folder);
                }

                string[] files = Directory.GetFiles(calculationFolderPath, "", SearchOption.TopDirectoryOnly);
                foreach (string file in files)
                {
                    var calculation = new Calculation
                    {
                        FileName = Path.GetFileName(file),
                        FilePath = file,
                        FileExtension = Path.GetExtension(file),
                        IsFolder = false,
                    };
                    Calculations.Add(calculation);
                }
            }
            catch (Exception)
            {
                CurrentCalculationFolder = calculationFolderPath2;
                try
                {
                    // Load folders
                    string[] directories = Directory.GetDirectories(calculationFolderPath2, "*", SearchOption.TopDirectoryOnly);
                    foreach (string directory in directories)
                    {
                        var folder = new Calculation
                        {
                            FileName = Path.GetFileName(directory),
                            FilePath = directory,
                            IsFolder = true
                        };
                        Calculations.Add(folder);
                    }

                    string[] files = Directory.GetFiles(calculationFolderPath2, "", SearchOption.TopDirectoryOnly);
                    foreach (string file in files)
                    {
                        var calculation = new Calculation
                        {
                            FileName = Path.GetFileName(file),
                            FilePath = file,
                            FileExtension = Path.GetExtension(file),
                            IsFolder = false,
                        };
                        Calculations.Add(calculation);
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Kan ikke finde mappe med beregninger. Det kan skyldes en gammel mappestruktur på sagen.\n\n", "Fejl!", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void LoadFolderContents(string folderPath, string type)
        {
            if (type == "Drawing")
            {
                Drawings.Clear();
                CurrentDrawingFolder = folderPath;
                try
                {
                    // Load folders
                    string[] directories = Directory.GetDirectories(folderPath, "*", SearchOption.TopDirectoryOnly);
                    foreach (string directory in directories)
                    {
                        var folder = new Drawing
                        {
                            FileName = Path.GetFileName(directory),
                            FilePath = directory,
                            IsFolder = true
                        };
                        Drawings.Add(folder);
                    }

                    // Load files
                    string[] files = Directory.GetFiles(folderPath, "*.pdf", SearchOption.TopDirectoryOnly);
                    foreach (string file in files)
                    {
                        var drawing = new Drawing
                        {
                            FileName = Path.GetFileName(file),
                            FilePath = file,
                            FileExtension = Path.GetExtension(file),
                            IsFolder = false
                        };
                        Drawings.Add(drawing);
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show($"Error loading folder contents: {e.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                Calculations.Clear();
                CurrentCalculationFolder = folderPath;
                try
                {
                    // Load folders
                    string[] directories = Directory.GetDirectories(folderPath, "*", SearchOption.TopDirectoryOnly);
                    foreach (string directory in directories)
                    {
                        var folder = new Calculation
                        {
                            FileName = Path.GetFileName(directory),
                            FilePath = directory,
                            IsFolder = true
                        };
                        Calculations.Add(folder);
                    }

                    // Load files
                    string[] files = Directory.GetFiles(folderPath, "*", SearchOption.TopDirectoryOnly);
                    foreach (string file in files)
                    {
                        var calculation = new Calculation
                        {
                            FileName = Path.GetFileName(file),
                            FilePath = file,
                            FileExtension = Path.GetExtension(file),
                            IsFolder = false
                        };
                        Calculations.Add(calculation);
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show($"Error loading folder contents: {e.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
        #endregion

        private bool CanCopyFiles()
        {
            return SelectedDrawings.Any();
        }

        private void CopySelectedFiles()
        {
            if (!Directory.Exists(CurrentCalculationFolder))
            {
                Directory.CreateDirectory(CurrentCalculationFolder);
            }

            foreach (var drawing in SelectedDrawings)
            {
                try
                {
                    var fileName = drawing.FileName;
                    var destFilePath = Path.Combine(CurrentCalculationFolder, fileName);

                    // Copy the file
                    File.Copy(drawing.FilePath, destFilePath, overwrite: false);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Kunne ikke kopiere {drawing}: {ex.Message}");
                }
            }

            LoadFolderContents(CurrentCalculationFolder, "Calculation");
        }
    }
}
