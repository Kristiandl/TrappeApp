using Dalton_Trapper.Model.Projektering_tab;
using Dalton_Trapper.Utilities;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Input;
using ImportExport = Dalton_Trapper.Model.Projektering_tab.ImportExport;

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

            MiljøklasseListe = new ObservableCollection<string> { "", "Passiv", "Moderat", "Agressiv", "Ekstra Agressiv" };

            // Commands
            OpenDrawingCommand = new RelayCommand(OpenDrawing);
            CommandBeregningsark = new RelayCommand(OpenBeregningsark);
            CommandDorneBrand = new RelayCommand(OpenDorneBrand);
            CommandIngeniørEkstern = new RelayCommand(OpenIngeniørEkstern);
        }

        #region Commands
        public ICommand OpenDrawingCommand { get; set; }

        public ICommand CommandBeregningsark { get; set; }
        public ICommand CommandDorneBrand { get; set; }
        public ICommand CommandIngeniørEkstern { get; set; }

        private void OpenDrawing(object parameter)
        {
            if (parameter is Drawing file)
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = file.FilePath,
                    UseShellExecute = true
                });
            }

            else if (parameter is Calculation file2)
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = file2.FilePath,
                    UseShellExecute = true
                });
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

        private void OpenIngeniørEkstern(object parameter)
        {
            if (!string.IsNullOrWhiteSpace(ProjectInfo.ProjectNumber))
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
            else 
            {
                MessageBox.Show("Sagsnummer er ikke indtastet!", "Fejl!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
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

        private ObservableCollection<Calculation> calculations;
        public ObservableCollection<Calculation> Calculations
        {
            get { return calculations ?? (calculations = new ObservableCollection<Calculation>()); }

            set { calculations = value; }
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
            string calculationFolderPath = $"S:\\{ProjectInfo.ProjectNumber.Substring(0, 3)}\\{ProjectInfo.ProjectNumber}\\09.Beregninger\\Elementstatik\\15 Trapper-Reposer og Skakte\\01 Styrkeberegninger";
            string calculationFolderPath2 = $"S:\\{ProjectInfo.ProjectNumber.Substring(0, 3)}\\{ProjectInfo.ProjectNumber}\\09.Beregninger\\Elementstatik\\15 Trapper-Reposer og Skakte";

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
                string[] pdfFiles = Directory.GetFiles(drawingFolderPath, "*pdf", SearchOption.TopDirectoryOnly);
                foreach (string file in pdfFiles)
                {
                    var drawing = new Drawing { FileName = Path.GetFileName(file), FilePath = file };
                    Drawings.Add(drawing);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Kan ikke finde mappen: \"4 Færdige PDF tegninger\". Det kan skyldes en gammel mappestruktur på sagen.\n\n" +
                                e.Message, "Fejl!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            
            
            // Load Calculations //
            try // Styrkeberegninger folder
            {
                string[] files = Directory.GetFiles(calculationFolderPath, "", SearchOption.TopDirectoryOnly);
                foreach (string file in files)
                {
                    var calculation = new Calculation { FileName = Path.GetFileName(file), FilePath = file };
                    Calculations.Add(calculation);
                }
            }
            catch (Exception)
            {
                try
                {
                    string[] files = Directory.GetFiles(calculationFolderPath2, "", SearchOption.TopDirectoryOnly);
                    foreach (string file in files)
                    {
                        var calculation = new Calculation { FileName = Path.GetFileName(file), FilePath = file };
                        Calculations.Add(calculation);
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Kan ikke finde mappe med beregninger. Det kan skyldes en gammel mappestruktur på sagen.\n\n", "Fejl!", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
        #endregion
    }
}
