using Dalton_Trapper.Model.Projektering_tab;
using Dalton_Trapper.Utilities;
using Dalton_Trapper.View;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Shapes;
using static Dalton_Trapper.ViewModel.FemdesignVM;
using Drawing = Dalton_Trapper.Model.Projektering_tab.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using FontSize = DocumentFormat.OpenXml.Wordprocessing.FontSize;
using ImportExport = Dalton_Trapper.Model.Projektering_tab.ImportExport;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Path = System.IO.Path;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using Tag = DocumentFormat.OpenXml.Wordprocessing.Tag;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using Word = Microsoft.Office.Interop.Word;


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
            GelænderListe = new ObservableCollection<string> { "Type A", "Type A u. ben", "Type A m. balusterben", "Type X" };
            MiljøklasseListe = new ObservableCollection<string> { "", "Passiv", "Moderat", "Aggressiv", "Ekstra Aggressiv" };
            BetonListe = new ObservableCollection<string> { "", "Beton", "Terrazzo" };
            SelectedDrawings = new ObservableCollection<Drawing>();
            SelectedCalculations = new ObservableCollection<Calculation>();
            TegningerHeader = "Tegninger";
            BeregningerHeader = "Beregninger";
            Gelænder = new Gelænder();


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
            CopyCorrectedDrawingsCommand = new RelayCommand(_ => CopySelectedCalculationFiles(), _ => IsCalculationItemSelected());
            SwitchDrawingFolderCommand = new RelayCommand(_ => SwitchDrawingFolder(), _ => CanSwitchDrawingFolder());
            DeleteCalculationItemCommand = new RelayCommand(_ => DeleteCalculationItem(), _ => IsCalculationItemSelected());
            CreateDocsCommand = new RelayCommand(_ => CreateDocs(), _ => ProjectNumberEntered() && IsTjeklisteCreated());
            DornUdnyttelseTjeklisteCommand = new RelayCommand(_ => DornUdnyttelseTjekliste(), _ => ProjectNumberEntered() && IsTjeklisteCreated());
            PrintPdfCommand = new RelayCommand(_ => PrintPdf(), _ => IsCalculationItemSelected());
            CreateSheetsCommand = new RelayCommand(_ => CreateSheets(), _ => ProjectNumberEntered() && IsTjeklisteCreated());

            OpretLigeløbCommand = new RelayCommand(_ => OpretBeregning("Ligeløb"), _ => CanAddFiles(Ligeløb));
            OpretKnækløbCommand = new RelayCommand(_ => OpretBeregning("Knækløb"), _ => CanAddFiles(Knækløb));
            OpretReposCommand = new RelayCommand(_ => OpretBeregning("Repos"), _ => CanAddFiles(Repos));
            OpretSvingløbCommand = new RelayCommand(_ => OpretBeregning("Svingløb"), _ => CanAddFiles(Svingløb));
            OpretDetaljeCommand = new RelayCommand(_ => OpretBeregning("Detalje"), _ => CanAddFiles(Detalje));
            OpretPladeCommand = new RelayCommand(_ => OpretBeregning("Plade"), _ => CanAddFiles(Plade));
            OpretVibrationskomfortCommand = new RelayCommand(_ => OpretBeregning("Vibrationskomfort"), _ => CanAddFiles(Vibrationskomfort));
            OpretRytmiskPersonlastCommand = new RelayCommand(_ => OpretBeregning("RytmiskPersonlast"), _ => CanAddFiles(RytmiskPersonlast));
            OpretReaktionArkCommand = new RelayCommand(_ => OpretBeregning("ReaktionArk"), _ => CanAddFiles(ReaktionArk));
            OpretGelænderCommand = new RelayCommand(_ => OpretBeregning("Gelænder"), _ => CanAddFiles(Gelænder.Name) && CanAddFiles(Gelænder.Type));
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
        public ICommand OpretVibrationskomfortCommand { get; set; }
        public ICommand OpretRytmiskPersonlastCommand { get; set; }
        public ICommand OpretReaktionArkCommand { get; set; }
        public ICommand OpretGelænderCommand { get; set; }
        public ICommand CopyDrawingsCommand { get; set; }
        public ICommand CopyCorrectedDrawingsCommand { get; set; }
        public ICommand SwitchDrawingFolderCommand { get; set; }
        public ICommand DeleteCalculationItemCommand { get; set; }
        public ICommand CreateDocsCommand { get; set; }
        public ICommand DornUdnyttelseTjeklisteCommand { get; set; }
        public ICommand PrintPdfCommand { get; set; }
        public ICommand CreateSheetsCommand { get; set; }

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
                    SetDrawingHeader(CurrentDrawingFolder);
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
                    SetCalculationHeader(CurrentCalculationFolder);
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

                    // Opret undermappen "Overslag" under "01 Styrkeberegninger"
                    if (folderName == "01 Styrkeberegninger")
                    {
                        string subFolderPath = Path.Combine(newFolderPath, "Overslag");

                        if (!Directory.Exists(subFolderPath))
                        {
                            Directory.CreateDirectory(subFolderPath);
                        }
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
            string fck = "";
            if (ProjectInfo.Miljøklasse == "Passiv" & ProjectInfo.Betontype == "Terrazzo")
            {
                environmentClass = "P";
                fck = "30";
            }
            else if (ProjectInfo.Miljøklasse == "Passiv" & ProjectInfo.Betontype == "Beton")
            {
                environmentClass = "P";
                fck = "35";
            }
            else if (ProjectInfo.Miljøklasse == "Moderat")
            {
                environmentClass = "M";
                fck = "35";
            }
            else if (ProjectInfo.Miljøklasse == "Aggressiv")
            {
                environmentClass = "A";
                fck = "35";
            }
            else if (ProjectInfo.Miljøklasse == "Ekstra Aggressiv")
            {
                environmentClass = "E";
                fck = "45";
            }
            else
            {
                fck = "30";
            }

            if (type == "Ligeløb")
            {
                string sheet = "N:\\Opslag\\Statik\\Dalton statikark\\LobEC_190701 rev15042026.xlsm";

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
                worksheet.Cells[6, 15] = ProjectInfo.Betontype == "Terrazzo" ? "ja" : "nej";

                Ligeløb = "";
            }
            else if (type == "Repos")
            {
                string sheet = "N:\\Opslag\\Statik\\Dalton statikark\\reposEC_v2010141 rev 15042026.xlsm";

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
                worksheet.Cells[7, 18] = ProjectInfo.Betontype == "Terrazzo" ? "ja" : "nej";

                Repos = "";
            }
            else if (type == "Knækløb")
            {
                string sheet = "N:\\Opslag\\Statik\\Dalton statikark\\LobrepEC_190701 rev 15042026.xlsm";

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
                worksheet.Cells[20, 14] = ProjectInfo.Betontype == "Terrazzo" ? "ja" : "nej";

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
                worksheet.Cells[80, 23] = ProjectInfo.Betontype == "Terrazzo" ? "Ja" : "Nej";

                Svingløb = "";
            }
            else if (type == "Detalje")
            {
                string sheet = "N:\\Opslag\\Statik\\Dalton statikark\\TrdetEC_191017 rev 15042026.xlsm";

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
                worksheet.Cells[6, 18] = ProjectInfo.Betontype == "Terrazzo" ? "ja" : "nej";

                Detalje = "";
            }
            else if (type == "Plade")
            {
                string sheet = "N:\\Opslag\\Statik\\Dalton statikark\\RApladEC_190701 rev 15042026.xlsm";

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
                worksheet.Cells[52, 7] = ProjectInfo.Liveload == "5" ? "C" : "A"; // Liveload category
                worksheet.Cells[52, 8] = ProjectInfo.Liveload == "5" ? "5" : "3"; // Liveload intensity
                worksheet.Cells[31, 9] = environmentClass;
                worksheet.Cells[2, 10] = ProjectInfo.ProjectNumber;
                worksheet.Cells[7, 17] = ProjectInfo.Betontype == "Terrazzo" ? "ja" : "nej";

                Plade = "";
            }
            else if (type == "Vibrationskomfort")
            {
                string sheet = "Model\\Projektering tab\\Beregningsark\\Vibrationskomfort_ganglast_bef.xlsm";

                string newFile = System.IO.Path.Combine(CurrentCalculationFolder, $"{Vibrationskomfort}.xlsm");

                if (File.Exists(newFile))
                {
                    MessageBox.Show($"Filen '{Vibrationskomfort}.xlsm' eksisterer i forvejen. Vælg et andet navn.", "Fejl!", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                File.Copy(sheet, newFile);

                // Updates the sheet with projectinformation
                excelApp = new Excel.Application();
                workbook = excelApp.Workbooks.Open(newFile);
                for (int i = 3; i < 5; i++)
                {
                    worksheet = (Excel.Worksheet)workbook.Sheets[i];
                    worksheet.Cells[4, 3] = ProjectInfo.ProjectName;
                    worksheet.Cells[5, 3] = Vibrationskomfort;
                    worksheet.Cells[4, 9] = ProjectInfo.ProjectNumber;
                    worksheet.Cells[6, 3] = Environment.UserName.ToUpper();
                }

                Vibrationskomfort = "";
            }
            else if (type == "RytmiskPersonlast")
            {
                string sheet = "Model\\Projektering tab\\Beregningsark\\Rytmisk_personlast_bef.xlsm";

                string newFile = System.IO.Path.Combine(CurrentCalculationFolder, $"{RytmiskPersonlast}.xlsm");

                if (File.Exists(newFile))
                {
                    MessageBox.Show($"Filen '{RytmiskPersonlast}.xlsm' eksisterer i forvejen. Vælg et andet navn.", "Fejl!", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                File.Copy(sheet, newFile);

                // Updates the sheet with projectinformation
                excelApp = new Excel.Application();
                workbook = excelApp.Workbooks.Open(newFile);

                worksheet = (Excel.Worksheet)workbook.Sheets[2];
                worksheet.Cells[4, 3] = ProjectInfo.ProjectName;
                worksheet.Cells[5, 3] = RytmiskPersonlast;
                worksheet.Cells[4, 9] = ProjectInfo.ProjectNumber;
                worksheet.Cells[6, 3] = Environment.UserName.ToUpper();


                RytmiskPersonlast = "";
            }
            else if (type == "ReaktionArk")
            {
                string sheet = "Model\\Projektering tab\\Beregningsark\\Overslag på reaktioner V03-11-2025.xlsx";

                string newFile = System.IO.Path.Combine(CurrentCalculationFolder, $"{ReaktionArk}.xlsx");

                if (File.Exists(newFile))
                {
                    MessageBox.Show($"Filen '{ReaktionArk}.xlsx' eksisterer i forvejen. Vælg et andet navn.", "Fejl!", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                File.Copy(sheet, newFile);

                ReaktionArk = "";
                LoadFolderContents(CurrentCalculationFolder, "Calculation");
                return;
            }
            else if (type == "Gelænder")
            {
                string newFile = System.IO.Path.Combine(CurrentCalculationFolder, $"{Gelænder.Name}.xlsm");
                int fckCellRow= 0;
                int fckCellColumn = 0;

                if (Gelænder.Type == "Type A")
                {
                    string sheet = "Model\\Projektering tab\\Beregningsark\\Gelænder Type A v 1.0 _2025 rev 3.xlsm";
                    fckCellRow = 27;
                    fckCellColumn = 27;

                    if (File.Exists(newFile))
                    {
                        MessageBox.Show($"Filen '{Gelænder.Name}.xlsm' eksisterer i forvejen. Vælg et andet navn.", "Fejl!", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    File.Copy(sheet, newFile);     
                }
                else if (Gelænder.Type == "Type A m. balusterben")
                {
                    string sheet = "N:\\Projektering\\Tilst\\Gelændere\\Nye regneark_September 2023\\Regneark gelænder Type A med balusterben _v 1.0_2025_rev 2.xlsm";
                    fckCellRow = 27;
                    fckCellColumn = 26;

                    if (File.Exists(newFile))
                    {
                        MessageBox.Show($"Filen '{Gelænder.Name}.xlsm' eksisterer i forvejen. Vælg et andet navn.", "Fejl!", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    File.Copy(sheet, newFile);
                }
                else if (Gelænder.Type == "Type A u. ben")
                {
                    string sheet = "N:\\Projektering\\Tilst\\Gelændere\\Nye regneark_September 2023\\Regneark gelænder Type A uden ben _v_ 1.0_2025_rev 2.xlsm";
                    fckCellRow = 25;
                    fckCellColumn = 26;

                    if (File.Exists(newFile))
                    {
                        MessageBox.Show($"Filen '{Gelænder.Name}.xlsm' eksisterer i forvejen. Vælg et andet navn.", "Fejl!", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    File.Copy(sheet, newFile);
                }
                else if (Gelænder.Type == "Type X")
                {
                    string sheet = "N:\\Projektering\\Tilst\\Gelændere\\Nye regneark_September 2023\\Regneark gelænder Type X_rev 04_09_2025.xlsm";
                    fckCellRow = 26;
                    fckCellColumn = 26;

                    if (File.Exists(newFile))
                    {
                        MessageBox.Show($"Filen '{Gelænder.Name}.xlsm' eksisterer i forvejen. Vælg et andet navn.", "Fejl!", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    File.Copy(sheet, newFile);
                }

                // Updates the sheet with projectinformation
                excelApp = new Excel.Application();
                workbook = excelApp.Workbooks.Open(newFile);
                worksheet = (Excel.Worksheet)workbook.Sheets[1];
                worksheet.Cells[4, 7] = ProjectInfo.ProjectName;
                worksheet.Cells[5, 7] = Gelænder.Name;
                worksheet.Cells[5, 13] = ProjectInfo.ProjectNumber;
                worksheet.Cells[9, 13] = Environment.UserName.ToUpper();
                worksheet.Cells[fckCellColumn, fckCellRow] = fck;
                Gelænder.Name = "";
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

        private bool IsCalculationItemSelected()
        {
            return SelectedCalculations.Any();
        }

        private void CopySelectedCalculationFiles()
        {
            foreach (var file in SelectedCalculations)
            {
                try
                {
                    var fileName = file.FileName;
                    var destFilePath = Path.Combine(CurrentDrawingFolder, fileName);

                    // Copy the file
                    File.Copy(file.FilePath, destFilePath, overwrite: true);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Kunne ikke kopiere {file}: {ex.Message}");
                }
            }

            LoadFolderContents(CurrentDrawingFolder, "Drawing");
        }

        private void SwitchDrawingFolder()
        {
            if (CurrentDrawingFolder == $"S:\\{ProjectInfo.ProjectNumber.Substring(0, 3)}\\{ProjectInfo.ProjectNumber}\\09.Beregninger\\Elementstatik\\15 Trapper-Reposer og Skakte\\02 Rettede tegninger")
            {
                LoadFolderContents(BaseDrawingFolder, "Drawing");
            }
            else 
            {
                LoadFolderContents($"S:\\{ProjectInfo.ProjectNumber.Substring(0, 3)}\\{ProjectInfo.ProjectNumber}\\09.Beregninger\\Elementstatik\\15 Trapper-Reposer og Skakte\\02 Rettede tegninger", "Drawing");
            }
        }

        private bool CanSwitchDrawingFolder()
        {
            if (ProjectNumberEntered() && Directory.Exists($"S:\\{ProjectInfo.ProjectNumber.Substring(0, 3)}\\{ProjectInfo.ProjectNumber}\\09.Beregninger\\Elementstatik\\15 Trapper-Reposer og Skakte\\02 Rettede tegninger"))
            {
                return true;
            }

            return false;
        }

        private void DeleteCalculationItem()
        {
            foreach (var file in SelectedCalculations)
            {
                try
                {
                    // Copy the file
                    File.Delete(file.FilePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Kunne ikke slette {file}: {ex.Message}");
                }
            }

            LoadFolderContents(CurrentCalculationFolder, "Calculation");
        }
        private bool IsTjeklisteCreated()
        {
            string file = $"S:\\{ProjectInfo.ProjectNumber.Substring(0, 3)}\\{ProjectInfo.ProjectNumber}\\09.Beregninger\\Elementstatik\\15 Trapper-Reposer og Skakte\\Tjekliste.docx";
            return File.Exists(file);
        }

        public static bool IsFileLocked(string filePath)
        {
            try
            {
                using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    // Hvis vi kan åbne med FileShare.None, er filen IKKE låst
                    return false;
                }
            }
            catch (IOException)
            {
                // Filen er i brug (låst)
                return true;
            }
        }
        private void CreateDocs()
        {
            string filePath = $"S:\\{ProjectInfo.ProjectNumber.Substring(0, 3)}\\{ProjectInfo.ProjectNumber}\\09.Beregninger\\Elementstatik\\15 Trapper-Reposer og Skakte\\Tjekliste.docx";

            var tjekliste = new Dictionary<string, string>();
            var elementinddeling = new List<string>();

            // Tjek om filen er låst
            if (IsFileLocked(filePath))
            {
                MessageBox.Show("Tjekliste.docx er åbent.\nLuk dokumentet før du fortsætter.", "Fil i brug",
                    MessageBoxButton.OK, MessageBoxImage.Warning);

                return;
            }

            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
            {
                var body = doc.MainDocumentPart.Document.Body;
                var tables = body.Elements<Table>();
                var tables_information = body.Elements<Table>().Take(2);
                var table_elements = body.Elements<Table>().Skip(4).FirstOrDefault();

                // Indlæser tjekliste data
                foreach (var table in tables_information)
                {
                    // Skips the header row
                    var rows = table.Elements<TableRow>().Skip(1);

                    foreach (var row in rows)
                    {
                        var cells = row.Elements<TableCell>().Take(2).ToList();

                        string key = cells[0].InnerText.Trim();
                        string value = cells[1].InnerText.Trim();

                        if (string.IsNullOrEmpty(value) || value == "")
                        {
                            MessageBox.Show($"Tjekliste indeholder tomme felter.\nUdfyld alle felter før du fortsætter.", "Fejl i tjekliste",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                        }

                        if (!string.IsNullOrEmpty(key))
                        {
                            tjekliste[key] = value;
                        }
                    }
                }

                // Indlæser elementinddeling
                var table_rows = table_elements.Elements<TableRow>().Skip(2);

                foreach (var row in table_rows)
                {
                    var cell = row.Elements<TableCell>().Skip(1).FirstOrDefault().InnerText;
                        
                    if (!string.IsNullOrEmpty(cell) || !string.IsNullOrWhiteSpace(cell))
                    {
                        elementinddeling.Add(cell);
                    }
                }

                if (elementinddeling.Count > 0)
                {
                    elementinddeling.RemoveAt(elementinddeling.Count - 1);
                }
            }

            string udarbejder = tjekliste["Ingeniør"] switch
            {
                "KDL" => "Kristian Damgaard Laugesen",
                "FST" => "Fie Sandberg Poulsen",
                "SOK" => "Sofi Karzoun",
                "NTO" => "Nils Tholstrup",
                _ => ""
            };

            string kontrollant = tjekliste["Kontrollant"] switch
            {
                "KDL" => "Kristian Damgaard Laugesen",
                "FST" => "Fie Sandberg Poulsen",
                "SOK" => "Sofi Karzoun",
                "NTO" => "Nils Tholstrup",
                _ => ""
            };

            string kfi = tjekliste["Konsekvensklasse"] switch
            {
                "CC2" => "1,0",
                "CC3" => "1,1",
                _ => ""
            };

            string kontrolniveauer = tjekliste["Konstruktionsklasse"] switch
            {
                "KK2" => "Udv. 10%",
                "KK3" => "Udv. 25%",
                "KK4" => "Maks.",
                _ => ""
            };

            if (udarbejder == "" || kontrollant == "")
            {
                MessageBox.Show($"Tjekliste indeholder ugyldige initialer for udarbejder og/eller kontrollant.\nBrug KDL, FST, SOK eller NTO.", "Fejl i tjekliste",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (kfi == "" || kontrolniveauer == "")
            {
                MessageBox.Show($"Tjekliste indeholder ugyldige værdier for konsekvensklasse og/eller konstruktionsklasse.\nBrug CC2 eller CC3 for konsekvensklasse og KK2, KK3 eller KK4 for konstruktionsklasse.", "Fejl i tjekliste",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string filePathA1 = "Model\\Projektering tab\\Statisk dokumentation\\A1.2.docx";
            string filenameA1 = $"A1.2.{tjekliste["Konstruktionsafsnitsnummer"]} Konstruktionsgrundlag - {tjekliste["Konstruktionsafsnitsnavn"]}.docx";
            string filePathA2 = "Model\\Projektering tab\\Statisk dokumentation\\A2.2.docx";
            string filenameA2 = $"A2.2.{tjekliste["Konstruktionsafsnitsnummer"]} Statiske beregninger - {tjekliste["Konstruktionsafsnitsnavn"]}.docx";
            string filePathA3 = "Model\\Projektering tab\\Statisk dokumentation\\A3.2.docx";
            string filenameA3 = $"A3.2.{tjekliste["Konstruktionsafsnitsnummer"]} Konstruktionstegninger - {tjekliste["Konstruktionsafsnitsnavn"]}.docx";
            string filePathA4 = "Model\\Projektering tab\\Statisk dokumentation\\A4.2.docx";
            string filenameA4 = $"A4.2.{tjekliste["Konstruktionsafsnitsnummer"]} Konstruktionsændringer - {tjekliste["Konstruktionsafsnitsnavn"]}.docx";
            string filePathA5 = "Model\\Projektering tab\\Statisk dokumentation\\A5.docx";
            string filenameA5 = $"A5.{tjekliste["Konstruktionsafsnitsnummer"]} Konstruktion som udført - {tjekliste["Konstruktionsafsnitsnavn"]}.docx";
            string filePathB1 = "Model\\Projektering tab\\Statisk dokumentation\\B1.2.docx";
            string filenameB1 = $"B1.2.{tjekliste["Konstruktionsafsnitsnummer"]} Statisk projektredegørelse - {tjekliste["Konstruktionsafsnitsnavn"]}.docx";
            string filePathB2 = "Model\\Projektering tab\\Statisk dokumentation\\B2.1.2.docx";
            string filenameB2 = $"B2.1.2.{tjekliste["Konstruktionsafsnitsnummer"]} Statisk kontrolplan projektering - {tjekliste["Konstruktionsafsnitsnavn"]}.docx";
            string filePathB3 = "Model\\Projektering tab\\Statisk dokumentation\\B3.1.2.docx";
            string filenameB3 = $"B3.1.2.{tjekliste["Konstruktionsafsnitsnummer"]} Statisk kontrolrapport projektering - {tjekliste["Konstruktionsafsnitsnavn"]}.docx";
            string filePathB32 = "Model\\Projektering tab\\Statisk dokumentation\\B3.2.2.docx";
            string filenameB32 = $"B3.2.2.CRH Statisk kontrolrapport udførelse (fremstilling).docx";


            var documents = new List<(string templatePath, string outputFileName)>
            {
                (filePathA1, filenameA1),
                (filePathA2, filenameA2),
                (filePathA3, filenameA3),
                (filePathA4, filenameA4),
                (filePathA5, filenameA5),
                (filePathB1, filenameB1),
                (filePathB2, filenameB2),
                (filePathB3, filenameB3),
                (filePathB32, filenameB32)
            };
            string targetFolder = $"S:\\{ProjectInfo.ProjectNumber.Substring(0, 3)}\\{ProjectInfo.ProjectNumber}\\09.Beregninger\\Elementstatik\\15 Trapper-Reposer og Skakte\\03 Statisk dokumentation";

            // Opret mappe hvis ikke findes
            if (!Directory.Exists(targetFolder))
            {
                Directory.CreateDirectory(targetFolder);
            }

            // Shared dictionary
            var tagValueMap = new Dictionary<string, string>
            {
                ["Konstruktionsafsnitsnummer"] = tjekliste["Konstruktionsafsnitsnummer"],
                ["Konstruktionsafsnitsnavn"] = tjekliste["Konstruktionsafsnitsnavn"],
                ["Projektnavn"] = tjekliste["Sagsnavn"],
                ["Adresse"] = tjekliste["Adresse"],
                ["Sagsnummer"] = tjekliste["Sagsnummer"],
                ["Udarbejdelsesdato"] = DateTime.Today.ToString("d"),
                ["Udarbejdet_navn"] = udarbejder,
                ["Udarbejdet_initialer"] = tjekliste["Ingeniør"],
                ["Kontrollant_navn"] = kontrollant,
                ["Kontrollant_initialer"] = tjekliste["Kontrollant"],
                ["Godkendt_navn"] = udarbejder,
                ["Godkendt_initialer"] = tjekliste["Ingeniør"],
                ["A1.1_bygværk_revision"] = tjekliste["A1.1 Konstruktionsgrundlag, Revision"],
                ["A1.1_bygværk_dato"] = tjekliste["A1.1 Konstruktionsgrundlag, Dato"],
                ["Opstalter"] = tjekliste["Nummerering af opstalter"],
                ["Afstand_til_væg"] = tjekliste["Afstand mellem elementer og væg"],
                ["Konsekvensklasse"] = tjekliste["Konsekvensklasse"],
                ["Konstruktionsklasse"] = tjekliste["Konstruktionsklasse"],
                ["KFI"] = kfi,
                ["Udførelsesklasse"] = tjekliste["Udførelsesklasse"],
                ["Miljøpåvirkning"] = tjekliste["Miljøpåvirkning"],
                ["Anvendelse"] = tjekliste["Anvendelse"],
                ["Brandtid"] = tjekliste["Brandkrav"].Replace("R", ""),
                ["Nyttelastkategori"] = tjekliste["Nyttelastkategori"],
                ["Nyttelast"] = tjekliste["Nyttelast"].Replace("kN/m2", ""),
                ["Beregn_A1.2_Udarbejdet_navn"] = udarbejder,
                ["Beregn_A2.2_Udarbejdet_navn"] = udarbejder,
                ["Beregn_A3.2_Udarbejdet_navn"] = udarbejder,
                ["Beregn_A4.2_Udarbejdet_navn"] = udarbejder,
                ["Beregn_B1.2_Udarbejdet_navn"] = udarbejder,
                ["Beregn_B2.1.2_Udarbejdet_navn"] = udarbejder,
                ["Beregn_B3.1.2_Udarbejdet_navn"] = kontrollant,
                ["Beregn_A1.2_Udarbejdet_initialer"] = tjekliste["Ingeniør"],
                ["Beregn_A2.2_Udarbejdet_initialer"] = tjekliste["Ingeniør"],
                ["Beregn_A3.2_Udarbejdet_initialer"] = tjekliste["Ingeniør"],
                ["Beregn_A4.2_Udarbejdet_initialer"] = tjekliste["Ingeniør"],
                ["Beregn_B1.2_Udarbejdet_initialer"] = tjekliste["Ingeniør"],
                ["Beregn_B2.1.2_Udarbejdet_initialer"] = tjekliste["Ingeniør"],
                ["Beregn_B3.1.2_Udarbejdet_initialer"] = tjekliste["Kontrollant"],
                ["Beregn_A1.2_Dato"] = DateTime.Today.ToString("d"),
                ["Beregn_A2.2_Dato"] = DateTime.Today.ToString("d"),
                ["Beregn_A3.2_Dato"] = DateTime.Today.ToString("d"),
                ["Beregn_A4.2_Dato"] = DateTime.Today.ToString("d"),
                ["Beregn_B2.1.2_Dato"] = DateTime.Today.ToString("d"),
                ["Beregn_B3.1.2_Dato"] = DateTime.Today.ToString("d"),
                ["model_A113"] = tjekliste["A113 model"].Replace("4LK / 5", "5"),
                ["Beregn_A1.2_Kontrollant_navn"] = kontrollant,
                ["Beregn_A2.2_Kontrollant_navn"] = kontrollant,
                ["Beregn_A3.2_Kontrollant_navn"] = kontrollant,
                ["Beregn_A4.2_Kontrollant_navn"] = kontrollant,
                ["Beregn_B1.2_Kontrollant_navn"] = kontrollant,
                ["Beregn_A1.2_Kontrollant_initialer"] = tjekliste["Kontrollant"],
                ["Beregn_A2.2_Kontrollant_initialer"] = tjekliste["Kontrollant"],
                ["Beregn_A3.2_Kontrollant_initialer"] = tjekliste["Kontrollant"],
                ["Beregn_A4.2_Kontrollant_initialer"] = tjekliste["Kontrollant"],
                ["Beregn_B1.2_Kontrollant_initialer"] = tjekliste["Kontrollant"],
                ["Beregn_B2.1.2_Kontrollant_initialer"] = tjekliste["Kontrollant"],
                ["Beregn_A2.2_kontrolniveau"] = kontrolniveauer,
                ["Beregn_A3.2_kontrolniveau"] = kontrolniveauer,
                ["Beregn_A4.2_kontrolniveau"] = kontrolniveauer,
                ["Beregn_B1.2_kontrolniveau"] = kontrolniveauer,
            };

            Word.Application wordApp = new Word.Application();

            // Loop igennem alle dokumenter
            foreach (var (templatePath, outputFileName) in documents)
            {
                string newFile = Path.Combine(targetFolder, outputFileName);

                if (File.Exists(newFile))
                {
                    MessageBox.Show($"Filen '{outputFileName}' eksisterer allerede. Springer over.",
                        "Advarsel", MessageBoxButton.OK, MessageBoxImage.Warning);
                    continue; // hopper videre i stedet for at stoppe alt
                }

                File.Copy(templatePath, newFile, true);

                // Lokal kopi af tagValueMap
                var localTagValueMap = new Dictionary<string, string>(tagValueMap);

                // Specialregel for B3.1.2
                if (templatePath == filePathB3)
                {
                    localTagValueMap["Udarbejdet_navn"] = kontrollant;
                    localTagValueMap["Udarbejdet_initialer"] = tjekliste["Kontrollant"];
                }

                Word.Document doc = null;

                try
                {
                    doc = wordApp.Documents.Open(
                        newFile,
                        ReadOnly: false, 
                        Visible: false
                    );

                    foreach (Word.ContentControl cc in doc.ContentControls)
                    {
                        if (localTagValueMap.TryGetValue(cc.Title, out string newValue))
                        {
                            cc.Range.Text = newValue;
                        }
                    }

                    doc.Save();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    if (doc != null)
                    {
                        doc.Close();
                        Marshal.ReleaseComObject(doc);
                    }
                }
            }

            // Indsætter elementinddelinger som overskrifter i A2.2
            var A2 = wordApp.Documents.Open(Path.Combine(targetFolder, filenameA2));

            // Gå til starten af side 11
            Word.Range range = A2.GoTo(
                What: Word.WdGoToItem.wdGoToPage,
                Which: Word.WdGoToDirection.wdGoToAbsolute,
                Count: 10
            );

            range.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            // Indsæt overskrifter
            foreach (var tekst in elementinddeling)
            {
                range.Text = tekst;
                range.set_Style(Word.WdBuiltinStyle.wdStyleHeading2);
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.InsertParagraphAfter();
                range.InsertBreak(Word.WdBreakType.wdPageBreak);
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }

            // Opdater indholdsfortegnelse
            foreach (Word.TableOfContents toc in A2.TablesOfContents)
            {
                toc.Update();
                toc.UpdatePageNumbers(); 
            }

            A2.Save();
            A2.Close();
            wordApp.Quit();
            Marshal.ReleaseComObject(wordApp);

            LoadFolderContents(CurrentCalculationFolder, "Calculation");

            MessageBox.Show("Alle dokumenter er oprettet!", "Succes!", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        #region Udnyttelse på dorne til tjekliste

        public class DornData
        {
            public string Elementnavn { get; set; }
            public string Elementype { get; set; }
            public string Dorntype { get; set; }
            public string Udnyttelsesgrad { get; set; }
            public double Dornlængde { get; set; }
            public double Ankerstørrelse { get; set; }
        }

        // Hjælpefunktion til sikker konvertering
        private double ToDouble(object value)
        {
            return Math.Round(double.TryParse(value?.ToString(), out double result) ? result : 0);
        }

        // Hjælpefunktion til at bestemme dorntype
        private void BestemDorntype(DornData data)
        {
            // Først: check for "Ingen dorn"
            if (string.Equals(data.Udnyttelsesgrad, "-", StringComparison.OrdinalIgnoreCase))
            {
                data.Dorntype = "-";
                return;
            }

            // Dernæst: match på længde og anker
            if (data.Dornlængde == 300)
            {
                if (data.Ankerstørrelse == 14)
                    data.Dorntype = "TR61";
                else if (data.Ankerstørrelse == 16)
                    data.Dorntype = "TR65";
                else
                    data.Dorntype = "Specialdorn";
            }
            else if (data.Dornlængde == 400)
            {
                if (data.Ankerstørrelse == 14)
                    data.Dorntype = "TR63";
                else if (data.Ankerstørrelse == 16)
                    data.Dorntype = "TR66";
                else if (data.Ankerstørrelse == 20)
                    data.Dorntype = "TR67";
                else
                    data.Dorntype = "Specialdorn";
            }
            else
            {
                data.Dorntype = "Specialdorn";
            }
        }

        // Hjælpefunktion til at sætte tekst i en tabelcelle
        private void SetCellText(TableCell cell, string text)
        {
            var paragraph = new Paragraph(
                new ParagraphProperties(
                    new Justification() { Val = JustificationValues.Center } // ← alignment
                ),
                new Run(
                    new RunProperties(
                        new RunFonts() { Ascii = "Trebuchet MS", HighAnsi = "Trebuchet MS" },
                        new FontSize() { Val = "20" } // ← størrelse (20 = 10 pt)
                    ),
                    new Text(text ?? "")
                )
            );

            cell.RemoveAllChildren<Paragraph>();
            cell.Append(paragraph);
        }


        private List<DornData> ReadData(string folderPath, List<string> elementinddeling)
        {
            var result = new List<DornData>();

            var excelApp = new Excel.Application
            {
                Visible = false,
                DisplayAlerts = false
            };

            try
            {
                foreach (var file in Directory.GetFiles(folderPath, "*.xlsm"))
                {
                    string fileName = Path.GetFileNameWithoutExtension(file);

                    // Case-insensitive match
                    if (!elementinddeling.Contains(fileName, StringComparer.OrdinalIgnoreCase))
                        continue;

                    Excel.Workbook workbook = excelApp.Workbooks.Open(file);
                    Excel.Worksheet sheet = workbook.Sheets[1];

                    try
                    {
                        string type = (sheet.Range["B8"].Value2 ?? "").ToString();

                        var data = new DornData
                        {
                            Elementnavn = fileName,
                            Elementype = type
                        };

                        if (type == "REPOSE")
                        {
                            data.Udnyttelsesgrad = ToDouble(sheet.Range["O118"].Value2 * 100) > 100 ? "-" : ToDouble(sheet.Range["O118"].Value2 * 100).ToString();
                            data.Dornlængde = ToDouble(sheet.Range["F106"].Value2);
                            data.Ankerstørrelse = ToDouble(sheet.Range["J110"].Value2);

                            result.Add(data);
                        }
                        else if (type == "LØB MED REPOSE I ENDE(R)")
                        {
                            data.Udnyttelsesgrad = ToDouble(sheet.Range["P129"].Value2 * 100) > 100 ? "-" : ToDouble(sheet.Range["P129"].Value2 * 100).ToString();
                            data.Dornlængde = ToDouble(sheet.Range["F118"].Value2);
                            data.Ankerstørrelse = ToDouble(sheet.Range["J122"].Value2);

                            result.Add(data);
                        }
                    }
                    finally
                    {
                        workbook.Close(false);
                    }
                }
            }
            finally
            {
                excelApp.Quit();
            }

            return result;
        }

        private void DornUdnyttelseTjekliste()
        {
            string filePathTjekliste = $"S:\\{ProjectInfo.ProjectNumber.Substring(0, 3)}\\{ProjectInfo.ProjectNumber}\\09.Beregninger\\Elementstatik\\15 Trapper-Reposer og Skakte\\Tjekliste.docx";

            var elementinddeling = new List<string>();

            // Tjek om Word-fil er låst
            if (IsFileLocked(filePathTjekliste))
            {
                MessageBox.Show("Tjekliste.docx er åbent.\nLuk dokumentet før du fortsætter.",
                    "Fil i brug",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            // Læs elementer fra Word
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePathTjekliste, false))
            {
                var body = doc.MainDocumentPart.Document.Body;
                var table_elements = body.Elements<Table>().Skip(4).FirstOrDefault();

                if (table_elements == null)
                    return;

                var table_rows = table_elements.Elements<TableRow>().Skip(2);

                foreach (var row in table_rows)
                {
                    var cell = row.Elements<TableCell>().Skip(1).FirstOrDefault()?.InnerText;

                    if (!string.IsNullOrWhiteSpace(cell))
                    {
                        elementinddeling.Add(cell.Trim());
                    }
                }

                if (elementinddeling.Count > 0)
                {
                    elementinddeling.RemoveAt(elementinddeling.Count - 1);
                }
            }

            // Kald Excel-læsning
            var dornDataListe = ReadData(CurrentCalculationFolder, elementinddeling);

            foreach (var element in dornDataListe)
                BestemDorntype(element);


            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePathTjekliste, true))
            {
                var body = doc.MainDocumentPart.Document.Body;
                var table = body.Elements<Table>().Skip(4).FirstOrDefault();

                if (table == null)
                    return;

                // Fjerne de to første overskriftslinjer og den sidste linje for brand
                var rows = table.Elements<TableRow>().Skip(2).ToList();
                rows.RemoveAt(rows.Count - 1);

                foreach (var row in rows)
                {
                    var cells = row.Elements<TableCell>().ToList();

                    //if (cells.Count < 4)
                    //    continue;

                    // Elementnavn står i kolonne 2 (index 1)
                    string elementNavn = cells[1].InnerText?.Trim();

                    if (string.IsNullOrWhiteSpace(elementNavn))
                        continue;

                    // Find match i din liste
                    var data = dornDataListe.FirstOrDefault(d =>
                        d.Elementnavn.Equals(elementNavn, StringComparison.OrdinalIgnoreCase));

                    if (data == null)
                        continue;

                    // Opdater celler
                    SetCellText(cells[3], data.Dorntype);
                    SetCellText(cells[4], data.Udnyttelsesgrad);
                }

                doc.MainDocumentPart.Document.Save();
            }

            var sb = new System.Text.StringBuilder();
            sb.AppendLine("Tjeklisten er udfyldt med følgende data:\n");

            foreach (var d in dornDataListe)
            {
                sb.AppendLine($"Element {d.Elementnavn} : {d.Dorntype} : {d.Udnyttelsesgrad}%");
            }

            MessageBox.Show(sb.ToString(), "Dorn Data",
                MessageBoxButton.OK,
                MessageBoxImage.Information);

        }
        #endregion

        private void PrintPdf()
        {
            Word.Application wordApp = null;

            try
            {
                wordApp = new Word.Application
                {
                    Visible = false,
                    ScreenUpdating = false
                };

                foreach (var file in SelectedCalculations)
                {
                    Word.Document doc = null;

                    try
                    {
                        doc = wordApp.Documents.Open(file.FilePath, ReadOnly: true, Visible: false);

                        foreach (Word.TableOfContents toc in doc.TablesOfContents)
                        {
                            toc.Update();
                            Marshal.ReleaseComObject(toc);
                        }

                        string pdfPath = Path.ChangeExtension(file.FilePath, ".pdf");

                        doc.ExportAsFixedFormat(
                            pdfPath,
                            Word.WdExportFormat.wdExportFormatPDF);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Kunne ikke eksportere {file.FilePath}: {ex.Message}");
                    }
                    finally
                    {
                        if (doc != null)
                        {
                            doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                            Marshal.ReleaseComObject(doc);
                        }
                    }
                }
            }
            finally
            {
                if (wordApp != null)
                {
                    wordApp.Quit();
                    Marshal.ReleaseComObject(wordApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            LoadFolderContents(CurrentCalculationFolder, "Calculation");
        }

        private void CreateSheets()
        {
            string filePath = $"S:\\{ProjectInfo.ProjectNumber.Substring(0, 3)}\\{ProjectInfo.ProjectNumber}\\09.Beregninger\\Elementstatik\\15 Trapper-Reposer og Skakte\\Tjekliste.docx";

            var elementinddeling = new Dictionary<string, string>();

            // Tjek om filen er låst
            if (IsFileLocked(filePath))
            {
                MessageBox.Show("Tjekliste.docx er åbent.\nLuk dokumentet før du fortsætter.", "Fil i brug",
                    MessageBoxButton.OK, MessageBoxImage.Warning);

                return;
            }

            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
            {
                var body = doc.MainDocumentPart.Document.Body;
                var table_elements = body.Elements<Table>().Skip(4).FirstOrDefault();

                if (table_elements == null)
                    return;

                var table_rows = table_elements.Elements<TableRow>().Skip(2).SkipLast(1);

                foreach (var row in table_rows)
                {
                    var cells = row.Elements<TableCell>().ToList();

                    if (cells.Count < 3)
                        continue;

                    string elementNr = cells[1].InnerText?.Trim();
                    string type = cells[2].InnerText?.Trim();

                    if (!string.IsNullOrEmpty(elementNr))
                    {
                        elementinddeling[elementNr] = type;
                    }
                }
            }

            Excel.Application excelApp = null;


            string environmentClass = ProjectInfo.Miljøklasse switch
            {
                "Passiv" => "P",
                "Moderat" => "M",
                "Aggressiv" => "A",
                "Ekstra Aggressiv" => "E",
                _ => ""
            };

            try
            {
                excelApp = new Excel.Application();

                foreach (var element in elementinddeling)
                {
                    string elementNr = element.Key;
                    string type = element.Value;

                    Excel.Workbook workbook = null;
                    Excel.Worksheet worksheet = null;

                    string newFile = System.IO.Path.Combine(CurrentCalculationFolder, $"{elementNr}.xlsm");

                    if (File.Exists(newFile))
                    {
                        MessageBox.Show($"Filen '{elementNr}.xlsm' eksisterer i forvejen. Vælg et andet navn.", "Fejl!", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    try
                    {
                        if (type == "L")
                        {
                            string sheet = "N:\\Opslag\\Statik\\Dalton statikark\\LobEC_190701 rev15042026.xlsm";
                            File.Copy(sheet, newFile);

                            workbook = excelApp.Workbooks.Open(newFile);
                            worksheet = (Excel.Worksheet)workbook.Sheets[1];

                            worksheet.Cells[2, 3] = ProjectInfo.ProjectName;
                            worksheet.Cells[6, 3] = elementNr;
                            worksheet.Cells[9, 9] = string.IsNullOrEmpty(ProjectInfo.Konsekvensklasse) ? 2 : ProjectInfo.Konsekvensklasse[^1].ToString();
                            worksheet.Cells[35, 7] = ProjectInfo.Liveload == "5" ? "C" : "A";
                            worksheet.Cells[11, 9] = environmentClass;
                            worksheet.Cells[2, 10] = ProjectInfo.ProjectNumber;
                            worksheet.Cells[6, 15] = ProjectInfo.Betontype == "Terrazzo" ? "ja" : "nej";

                            workbook.Save();
                        }
                        else if (type == "LR")
                        {
                            string sheet = "N:\\Opslag\\Statik\\Dalton statikark\\LobrepEC_190701 rev 15042026.xlsm";                            
                            File.Copy(sheet, newFile);

                            // Updates the sheet with projectinformation
                            workbook = excelApp.Workbooks.Open(newFile);
                            worksheet = (Excel.Worksheet)workbook.Sheets[1];
                            worksheet.Cells[2, 3] = ProjectInfo.ProjectName;
                            worksheet.Cells[6, 3] = elementNr;
                            worksheet.Cells[9, 9] = ProjectInfo.Konsekvensklasse == "" ? 2 : ProjectInfo.Konsekvensklasse.Substring(ProjectInfo.Konsekvensklasse.Length - 1, 1); // Consequence class
                            worksheet.Cells[50, 7] = ProjectInfo.Liveload == "5" ? "C" : "A"; // Liveload
                            worksheet.Cells[11, 9] = environmentClass;
                            worksheet.Cells[2, 10] = ProjectInfo.ProjectNumber;
                            worksheet.Cells[20, 14] = ProjectInfo.Betontype == "Terrazzo" ? "ja" : "nej";
                            workbook.Save();
                        }
                        else if (type == "RE")
                        {
                            string sheet = "N:\\Opslag\\Statik\\Dalton statikark\\reposEC_v2010141 rev 15042026.xlsm";
                            File.Copy(sheet, newFile);

                            // Updates the sheet with projectinformation
                            workbook = excelApp.Workbooks.Open(newFile);
                            worksheet = (Excel.Worksheet)workbook.Sheets[1];
                            worksheet.Cells[2, 3] = ProjectInfo.ProjectName;
                            worksheet.Cells[6, 3] = elementNr;
                            worksheet.Cells[9, 9] = ProjectInfo.Konsekvensklasse == "" ? 2 : ProjectInfo.Konsekvensklasse.Substring(ProjectInfo.Konsekvensklasse.Length - 1, 1); // Consequence class
                            worksheet.Cells[44, 7] = ProjectInfo.Liveload == "5" ? "C" : "A"; // Liveload
                            worksheet.Cells[11, 9] = environmentClass;
                            worksheet.Cells[2, 10] = ProjectInfo.ProjectNumber;
                            worksheet.Cells[7, 18] = ProjectInfo.Betontype == "Terrazzo" ? "ja" : "nej";
                            workbook.Save();
                        }
                        else if (type == "PL")
                        {
                            string sheet = "N:\\Opslag\\Statik\\Dalton statikark\\RApladEC_190701 rev 15042026.xlsm";
                            File.Copy(sheet, newFile);

                            // Updates the sheet with projectinformation
                            workbook = excelApp.Workbooks.Open(newFile);
                            worksheet = (Excel.Worksheet)workbook.Sheets[1];
                            worksheet.Cells[2, 3] = ProjectInfo.ProjectName;
                            worksheet.Cells[6, 3] = elementNr;
                            worksheet.Cells[29, 9] = ProjectInfo.Konsekvensklasse == "" ? 2 : ProjectInfo.Konsekvensklasse.Substring(ProjectInfo.Konsekvensklasse.Length - 1, 1); // Consequence class
                            worksheet.Cells[52, 7] = ProjectInfo.Liveload == "5" ? "C" : "A"; // Liveload category
                            worksheet.Cells[52, 8] = ProjectInfo.Liveload == "5" ? "5" : "3"; // Liveload intensity
                            worksheet.Cells[31, 9] = environmentClass;
                            worksheet.Cells[2, 10] = ProjectInfo.ProjectNumber;
                            worksheet.Cells[7, 17] = ProjectInfo.Betontype == "Terrazzo" ? "ja" : "nej";
                            workbook.Save();
                        }
                        else if (type == "SV")
                        {
                            string sheet = "Model\\Projektering tab\\Beregningsark\\Vibrationskomfort_ganglast_bef.xlsm";
                            File.Copy(sheet, newFile);

                            // Updates the sheet with projectinformation
                            workbook = excelApp.Workbooks.Open(newFile);
                            for (int i = 3; i < 5; i++)
                            {
                                worksheet = (Excel.Worksheet)workbook.Sheets[i];
                                worksheet.Cells[4, 3] = ProjectInfo.ProjectName;
                                worksheet.Cells[5, 3] = elementNr;
                                worksheet.Cells[4, 9] = ProjectInfo.ProjectNumber;
                                worksheet.Cells[6, 3] = Environment.UserName.ToUpper();
                            }
                            workbook.Save();
                        }
                    }
                    finally
                    {
                        if (workbook != null)
                        {
                            workbook.Close(false);
                            Marshal.ReleaseComObject(workbook);
                        }

                        if (worksheet != null)
                            Marshal.ReleaseComObject(worksheet);
                    }    
                }
            }

            finally
            {
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
            }

            LoadFolderContents(CurrentCalculationFolder, "Calculation");
        }
        #endregion

        #region Collections for user input
        public ObservableCollection<string> KonsekvensklasseListe { get; set; }

        public ObservableCollection<string> MiljøklasseListe { get; set; }
        public ObservableCollection<string> BetonListe { get; set; }
        public ObservableCollection<string> GelænderListe { get; set; }


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

        private String _baseDrawingFolder;
        public String BaseDrawingFolder
        {
            get => _baseDrawingFolder;
            set
            {
                if (_baseDrawingFolder != value)
                {
                    _baseDrawingFolder = value;
                    OnPropertyChanged(nameof(BaseDrawingFolder));
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

        private String _vibrationskomfort;
        public String Vibrationskomfort
        {
            get => _vibrationskomfort;
            set
            {
                if (_vibrationskomfort != value)
                {
                    _vibrationskomfort = value;
                    OnPropertyChanged(nameof(Vibrationskomfort));
                }
            }
        }

        private String _rytmiskPersonlast;
        public String RytmiskPersonlast
        {
            get => _rytmiskPersonlast;
            set
            {
                if (_rytmiskPersonlast != value)
                {
                    _rytmiskPersonlast = value;
                    OnPropertyChanged(nameof(RytmiskPersonlast));
                }
            }
        }
        private String _reaktionArk;
        public String ReaktionArk
        {
            get => _reaktionArk;
            set
            {
                if (_reaktionArk != value)
                {
                    _reaktionArk = value;
                    OnPropertyChanged(nameof(ReaktionArk));
                }
            }
        }

        private Gelænder _gelænder;
        public Gelænder Gelænder
        {
            get => _gelænder;
            set
            {
                if (_gelænder != value)
                {
                    _gelænder = value;
                    OnPropertyChanged(nameof(Gelænder));
                }
            }
        }

        private String _tegningerHeader;
        public String TegningerHeader
        {
            get => _tegningerHeader;
            set
            {
                if (_tegningerHeader != value)
                {
                    _tegningerHeader = value;
                    OnPropertyChanged(nameof(TegningerHeader));
                }
            }
        }

        private String _beregningerHeader;
        public String BeregningerHeader
        {
            get => _beregningerHeader;
            set
            {
                if (_beregningerHeader != value)
                {
                    _beregningerHeader = value;
                    OnPropertyChanged(nameof(BeregningerHeader));
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
                    Betontype = ProjectInfo.Betontype,
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
                    var data = await System.Threading.Tasks.Task.Run(() => ImportExport.LoadFromJson(fullPath));

                    // Assign project info to app
                    ProjectInfo.ProjectNumber = data.ProjectNumber;
                    ProjectInfo.ProjectName = data.ProjectName;
                    ProjectInfo.Konsekvensklasse = data.Konsekvensklasse;
                    ProjectInfo.Miljøklasse = data.Miljøklasse;
                    ProjectInfo.Liveload = data.Liveload;
                    ProjectInfo.Betontype = data.Betontype;
                }
                else
                {
                    ProjectInfo.ProjectName = "";
                    ProjectInfo.Konsekvensklasse = "";
                    ProjectInfo.Miljøklasse = "";
                    ProjectInfo.Liveload = "";
                    ProjectInfo.Betontype = "";
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
                BaseDrawingFolder = drawingFolderPath;
                SetDrawingHeader(drawingFolderPath);
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
                    BaseDrawingFolder = drawingFolderPath2;
                    SetDrawingHeader(drawingFolderPath2);
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
                SetCalculationHeader(calculationFolderPath);
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
                    SetCalculationHeader(calculationFolderPath2);
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

                    SetDrawingHeader(CurrentDrawingFolder);
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
                    SetCalculationHeader(CurrentCalculationFolder);
                }
                catch (Exception e)
                {
                    MessageBox.Show($"Error loading folder contents: {e.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void SetDrawingHeader(string folderPath)
        {
            string directory = new DirectoryInfo(folderPath).Name;
            TegningerHeader = "Tegninger: " + directory;
        }
        private void SetCalculationHeader(string folderPath)
        {
            string directory = new DirectoryInfo(folderPath).Name;
            BeregningerHeader = "Beregninger: " + directory;
        }
        #endregion
    }
}
