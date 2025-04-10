using Dalton_Trapper.Model;
using Dalton_Trapper.Model.Draw;
using Dalton_Trapper.Model.ImportExport;
using Dalton_Trapper.Model.TestPage;
using Dalton_Trapper.Utilities;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Labels = Dalton_Trapper.Model.Labels;

namespace Dalton_Trapper.ViewModel
{
    class TestPageVM : Utilities.ViewModelBase
    {
        public List<string> FormOptions { get; } = new List<string> { "Konkav", "Konveks" };
        public List<string> KvalitetListe { get; } = new List<string> { "K", "Y", "P", "Z", "N", "R" };
        public ObservableCollection<TransferOptions> Options { get; set; }
        public bool IsDocumentationTemplateChecked => Options[3].IsYesChecked;

        private void GridView_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is DataGrid dataGrid)
                dataGrid.SelectedItem = null;
        }




        public ObservableCollection<PointViewModel> Points { get; set; } = new();

        private PointViewModel _draggingPoint;

        public ICommand StartDragCommand => new RelayCommand2<PointViewModel>(point =>
        {
                _draggingPoint = point;
                Mouse.AddMouseMoveHandler(Application.Current.MainWindow, OnMouseMove);
                Mouse.AddMouseUpHandler(Application.Current.MainWindow, OnMouseUp);

        });

        private void OnMouseMove(object sender, MouseEventArgs e)
        {
            if (_draggingPoint == null) return;
            var pos = e.GetPosition(Application.Current.MainWindow);
            _draggingPoint.X = pos.X;
            _draggingPoint.Y = pos.Y;

            UpdateGeometry();
        }

        private void OnMouseUp(object sender, MouseEventArgs e)
        {
            _draggingPoint = null;
            Mouse.RemoveMouseMoveHandler(Application.Current.MainWindow, OnMouseMove);
            Mouse.RemoveMouseUpHandler(Application.Current.MainWindow, OnMouseUp);
        }


        public TestPageVM()
        {
            coord = new ObservableCollection<Entries>();

            pointSupports = new ObservableCollection<Entries>();
            lineSupports = new ObservableCollection<Entries>();
            pointLoads = new ObservableCollection<Entries>();
            lineLoads = new ObservableCollection<Entries>();

            // Updates IDs when the collection is changed
            coord.CollectionChanged += (s, e) => UpdateIDs(coord, "P", (entry, id) => entry.ID = id);
            pointSupports.CollectionChanged += (s, e) => UpdateIDs(pointSupports, "PS", (entry, id) => entry.ID = id);
            lineSupports.CollectionChanged += (s, e) => UpdateIDs(lineSupports, "LS", (entry, id) => entry.ID = id);
            pointLoads.CollectionChanged += (s, e) => UpdateIDs(pointLoads, "PL", (entry, id) => entry.ID = id);
            lineLoads.CollectionChanged += (s, e) => UpdateIDs(lineLoads, "LL", (entry, id) => entry.ID = id);

            // Add commands
            AddRowCommandCoord = new RelayCommand(parameter => AddRow(parameter, coord, "P", (entry, id) => entry.ID = id));
            AddRowCommandCurvedLine = new RelayCommand(AddRowCurvedLine);
            AddRowCommandPointSupport = new RelayCommand(parameter => AddRow(parameter, pointSupports, "PS", (entry, id) => entry.ID = id));
            AddRowCommandLineSupport = new RelayCommand(parameter => AddRow(parameter, lineSupports, "LS", (entry, id) => entry.ID = id));
            AddRowCommandPointLoad = new RelayCommand(parameter => AddRow(parameter, pointLoads, "PL", (entry, id) => entry.ID = id));
            AddRowCommandLineLoad = new RelayCommand(parameter => AddRow(parameter, lineLoads, "LL", (entry, id) => entry.ID = id));

            // Delete commands
            DeleteRowCommandCoord = new RelayCommand(parameter => DeleteRow(parameter, coord, "P", (entry, id) => entry.ID = id));
            DeleteRowCommandCurvedLine = new RelayCommand(DeleteRowCurvedLine);
            DeleteRowCommandPointSupport = new RelayCommand(parameter => DeleteRow(parameter, pointSupports, "PS", (entry, id) => entry.ID = id));
            DeleteRowCommandLineSupport = new RelayCommand(parameter => DeleteRow(parameter, lineSupports, "LS", (entry, id) => entry.ID = id));
            DeleteRowCommandPointLoad = new RelayCommand(parameter => DeleteRow(parameter, pointLoads, "PL", (entry, id) => entry.ID = id));
            DeleteRowCommandLineLoad = new RelayCommand(parameter => DeleteRow(parameter, lineLoads, "LL", (entry, id) => entry.ID = id));


            BetonstyrkeListe = new ObservableCollection<string> { "C20/25", "C25/30", "C30/37", "C35/45", "C40/50", "C45/55", "C50/60" };
            Betonstyrke = "C30/37";

            Points.Add(new PointViewModel { X = 0, Y = 0 });
            Points.Add(new PointViewModel { X = 0, Y = 2000 });
            Points.Add(new PointViewModel { X = 1000, Y = 2000 });
            Points.Add(new PointViewModel { X = 1000, Y = 0 });
            Coord.Add(new Entries { X = 0, Y = 0 });
            Coord.Add(new Entries { X = 0, Y = 2000 });
            Coord.Add(new Entries { X = 1000, Y = 2000 });
            Coord.Add(new Entries { X = 1000, Y = 0 });

            // Updates geometry when collections are changed
            SubscribeToCollection(coord);

            SubscribeToCollection(Points);

            SubscribeToCollection(pointSupports);
            SubscribeToCollection(lineSupports);
            SubscribeToCollection(pointLoads);
            SubscribeToCollection(lineLoads);

            // Update visualization when initialized
            UpdateGeometry();
        }

        #region Collections for user inputs
        private ObservableCollection<Entries> coord;
        public ObservableCollection<Entries> Coord
        {
            get { return coord ?? (coord = new ObservableCollection<Entries>()); }

            set { coord = value; }
        }

        private ObservableCollection<CurvedLine> curvedLines;
        public ObservableCollection<CurvedLine> CurvedLines
        {
            get { return curvedLines ?? (curvedLines = new ObservableCollection<CurvedLine>()); }

            set { curvedLines = value; }
        }
        private ObservableCollection<Entries> pointSupports;
        public ObservableCollection<Entries> PointSupports
        {

            get { return pointSupports ?? (pointSupports = new ObservableCollection<Entries>()); }

            set
            {
                pointSupports = value;
            }
        }

        private ObservableCollection<Entries> lineSupports;
        public ObservableCollection<Entries> LineSupports
        {

            get { return lineSupports ?? (lineSupports = new ObservableCollection<Entries>()); }

            set { lineSupports = value; }
        }

        private ObservableCollection<Entries> pointLoads;

        public ObservableCollection<Entries> PointLoads
        {
            get { return pointLoads; }
            set { pointLoads = value; }
        }

        private ObservableCollection<Entries> lineLoads;
        public ObservableCollection<Entries> LineLoads
        {
            get { return lineLoads; }
            set { lineLoads = value; }
        }

        private ObservableCollection<Reinforcement> netReinforcement;
        public ObservableCollection<Reinforcement> NetReinforcement
        {
            get { return netReinforcement; }
            set { netReinforcement = value; }
        }

        private ObservableCollection<Reinforcement> additionalReinforcement;
        public ObservableCollection<Reinforcement> AdditionalReinforcement
        {
            get { return additionalReinforcement; }
            set { additionalReinforcement = value; }
        }

        public ObservableCollection<string> BetonstyrkeListe { get; set; }

        private string _betonstyrke;
        public string Betonstyrke
        {
            get => _betonstyrke;
            set
            {
                _betonstyrke = value;
                OnPropertyChanged(nameof(Betonstyrke));
            }
        }

        public ObservableCollection<string> KonsekvensklasseListe { get; set; }
        private string _konsekvensklasse { get; set; }
        public string Konsekvensklasse
        {
            get => _konsekvensklasse;
            set
            {
                _konsekvensklasse = value;
                OnPropertyChanged(nameof(Konsekvensklasse));
            }
        }

        private string _slabName;
        public string SlabName
        {
            get => _slabName; 
            set
            {
                _slabName = value;
                OnPropertyChanged(nameof(SlabName));
            }
        }
        private string _plateThickness;
        public string PlateThickness
        {
            get => _plateThickness;
            set
            {
                _plateThickness = value;
                OnPropertyChanged(nameof(PlateThickness));
            }
        }
        private string _plateLiveLoad;
        public string PlateLiveLoad
        {
            get => _plateLiveLoad;
            set
            {
                _plateLiveLoad = value;
                OnPropertyChanged(nameof(PlateLiveLoad));
            }
        }
        private string _psi_2q;
        public string Psi_2q
        {
            get => _psi_2q;
            set
            {
                _psi_2q = value;
                OnPropertyChanged(nameof(Psi_2q));
            }
        }
        private string _extraG;
        public string ExtraG
        {
            get => _extraG;
            set
            {
                _extraG = value;
                OnPropertyChanged(nameof(ExtraG));
            }
        }
        private string _sk;
        public string Sk
        {
            get => _sk;
            set
            {
                _sk = value;
                OnPropertyChanged(nameof(Sk));
            }
        }
        private string _sk_ophobning;
        public string Sk_ophobning
        {
            get => _sk_ophobning;
            set
            {
                _sk_ophobning = value;
                OnPropertyChanged(nameof(Sk_ophobning));
            }
        }
        private string _projectTitle;
        public string ProjectTitle
        {
            get => _projectTitle;
            set
            {
                _projectTitle = value;
                OnPropertyChanged(nameof(ProjectTitle));
            }
        }
        private string _projectNumber;
        public string ProjectNumber
        {
            get => _projectNumber;
            set
            {
                _projectNumber = value;
                OnPropertyChanged(nameof(ProjectNumber));
            }
        }

        #endregion

        #region Collections for drawing items
        private PathGeometry _slabGeometry;
        public PathGeometry SlabGeometry
        { get => _slabGeometry;
            private set {
                _slabGeometry = value;
                OnPropertyChanged(nameof(SlabGeometry));
            } }

        private GeometryGroup _pointGeometry;
        public GeometryGroup PointGeometry
        {
            get => _pointGeometry;
            private set
            {
                _pointGeometry = value;
                OnPropertyChanged(nameof(PointGeometry));
            }
        }
        private GeometryGroup _pointLoadGeometry;
        public GeometryGroup PointLoadGeometry
        {
            get => _pointLoadGeometry;
            private set
            {
                _pointLoadGeometry = value;
                OnPropertyChanged(nameof(PointLoadGeometry));
            }
        }

        private GeometryGroup _lineLoadGeometry;
        public GeometryGroup LineLoadGeometry
        {
            get => _lineLoadGeometry;
            private set
            {
                _lineLoadGeometry = value;
                OnPropertyChanged(nameof(LineLoadGeometry));
            }
        }

        private GeometryGroup _pointSupportGeometry;
        public GeometryGroup PointSupportGeometry
        {
            get => _pointSupportGeometry;
            private set
            {
                _pointSupportGeometry = value;
                OnPropertyChanged(nameof(PointSupportGeometry));
            }
        }

        private GeometryGroup _lineSupportGeometry;
        public GeometryGroup LineSupportGeometry
        {
            get => _lineSupportGeometry;
            private set
            {
                _lineSupportGeometry = value;
                OnPropertyChanged(nameof(LineSupportGeometry));
            }
        }

        private ObservableCollection<Labels> _canvasLabels;
        public ObservableCollection<Labels> CanvasLabels
        {
            get => _canvasLabels;
            private set
            {
                _canvasLabels = value;
                OnPropertyChanged(nameof(CanvasLabels));
            }
        }

        private GeometryGroup _canvasDimensionLines;
        public GeometryGroup CanvasDimensionLines
        {
            get => _canvasDimensionLines;
            private set
            {
                _canvasDimensionLines = value;
                OnPropertyChanged(nameof(CanvasDimensionLines));
            }
        }

        private ObservableCollection<Labels> _dimensionlineLabels;
        public ObservableCollection<Labels> DimensionlineLabels
        {
            get => _dimensionlineLabels;
            private set
            {
                _dimensionlineLabels = value;
                OnPropertyChanged(nameof(DimensionlineLabels));
            }
        }
        #endregion

        #region Commands
        public ICommand AddRowCommandCoord { get; set; }
        public ICommand AddRowCommandCurvedLine { get; set; }
        public ICommand DeleteRowCommandCoord { get; set; }
        public ICommand DeleteRowCommandCurvedLine { get; set; }
        public ICommand AddRowCommandPointSupport { get; set; }
        public ICommand DeleteRowCommandPointSupport { get; set; }
        public ICommand AddRowCommandLineSupport { get; set; }
        public ICommand DeleteRowCommandLineSupport { get; set; }
        public ICommand AddRowCommandPointLoad { get; set; }
        public ICommand DeleteRowCommandPointLoad { get; set; }
        public ICommand AddRowCommandLineLoad { get; set; }
        public ICommand DeleteRowCommandLineLoad { get; set; }
        public ICommand ImportCommand { get; set; }
        public ICommand SaveCommand { get; set; }
        public ICommand PushToFEM { get; set; }
        #endregion

        #region Add and delete functions
        // Generalized UpdateIDs method
        private void UpdateIDs<T>(ObservableCollection<T> collection, string prefix, Action<T, string> updateIDAction)
        {
            for (int i = 0; i < collection.Count; i++)
            {
                // Generate the new ID
                string newID = $"{prefix}{i + 1}";

                // Use the provided action to update the ID
                updateIDAction(collection[i], newID);
            }
        }

        // Generalized AddRow method
        private void AddRow<T>(object parameter, ObservableCollection<T> collection, string idPrefix, Action<T, string> updateIDAction)
            where T : class
        {
            // Cast the parameter to the generic type
            var entry = parameter as T;

            if (entry != null)
            {
                // Find the index of the current entry in the collection
                int index = collection.IndexOf(entry);

                // Create a new instance of the generic type
                var newEntry = Activator.CreateInstance<T>();

                // Insert the new entry at the desired index
                collection.Insert(index + 1, newEntry);

                // Update IDs to keep them consistent
                UpdateIDs(collection, idPrefix, updateIDAction);
            }
        }

        private void AddRowCurvedLine(object parameter)
        {
            var entry = parameter as CurvedLine;

            if (entry != null)
            {
                // Find the index of the current entry in the ObservableCollection
                int index = CurvedLines.IndexOf(entry);

                // Create a new entry (can customize default values here)
                var newEntry = new CurvedLine { };

                // Insert the new entry at the desired index
                CurvedLines.Insert(index + 1, newEntry);
            }
        }

        // Generalized DeleteRow method
        private void DeleteRow<T>(object parameter, ObservableCollection<T> collection, string idPrefix, Action<T, string> updateIDAction)
            where T : class
        {
            // Cast the parameter to the generic type
            var entry = parameter as T;

            if (entry != null)
            {
                // Remove the entry from the collection
                collection.Remove(entry);

                // Update IDs to keep them consistent
                UpdateIDs(collection, idPrefix, updateIDAction);
            }
        }


        private void DeleteRowCurvedLine(object parameter)
        {
            // Get the DataGridRow that contains this button
            var entry = parameter as CurvedLine;

            if (entry != null)
            {
                // Remove the entry from the corresponding ObservableCollection
                CurvedLines.Remove(entry);
            }
        }
        #endregion 

        #region Functions for visualisation
        private void SubscribeToCollection<T>(ObservableCollection<T> collection) where T : INotifyPropertyChanged
        {
            collection.CollectionChanged += (s, e) =>
            {
                if (e.NewItems != null)
                {
                    foreach (T item in e.NewItems)
                        item.PropertyChanged += OnEntryPropertyChanged;
                }
                if (e.OldItems != null)
                {
                    foreach (T item in e.OldItems)
                        item.PropertyChanged -= OnEntryPropertyChanged;
                }

                UpdateGeometry(); // Or any other logic that needs to run on collection change
            };

            // Ensure existing items are also subscribed
            foreach (var item in collection)
            {
                item.PropertyChanged += OnEntryPropertyChanged;
            }
        }

        private void OnEntryPropertyChanged(object sender, PropertyChangedEventArgs e) 
        { 
            UpdateGeometry(); 
        } 

        private void UpdateGeometry()
        {
            if (coord.Count < 3)
            {
                return;
            }

            // Define translation values (to add a margin around drawing)
            int translateX = 60;
            int translateY = 60;

            // Determine scaling factor for view
            var _canvasHeight = 600;
            var _canvasWidth = 600;

            double minX = coord.Min(c => c.X);
            double minY = coord.Min(c => c.Y);
            double maxX = coord.Max(c => c.X);
            double maxY = coord.Max(c => c.Y);

            double rangeX = maxX - minX;
            double rangeY = maxY - minY;

            // Prevent division by 0
            if (rangeX == 0) rangeX = 1;
            if (rangeY == 0) rangeY = 1;

            double scaleX = (_canvasWidth - 2 * translateX) / rangeX;
            double scaleY = (_canvasHeight - 2 * translateY) / rangeY;

            double scaleFactor = Math.Min(scaleX, scaleY);

            // Draw functions
            SlabGeometry = SlabDrawer.CreateSlabGeometry(Coord, CurvedLines, scaleFactor, translateX, translateY);
            PointLoadGeometry = LoadDrawer.DrawPointLoad(pointLoads, scaleFactor, translateX, translateY, 100);
            LineLoadGeometry = LoadDrawer.DrawLineLoad(lineLoads, scaleFactor, translateX, translateY);
            PointSupportGeometry = SupportDrawer.DrawPointSupport(pointSupports, scaleFactor, translateX, translateY, 90);
            LineSupportGeometry = SupportDrawer.DrawLineSupport(lineSupports, scaleFactor, translateX, translateY);

            PointGeometry = PointDrawer.DrawPoints(Points, scaleFactor, translateX, translateY);

            (CanvasDimensionLines, DimensionlineLabels) = DimensionLineDrawer.DrawCanvasDimensionlines(Coord, scaleFactor, translateX, translateY);
            

            UpdateLabels(scaleFactor, translateX, translateY);
        }

        private void UpdateLabels(double scaleFactor, int translateX, int translateY)
        {
            int offset = 25;
            
            // Corner points
            CanvasLabels = new ObservableCollection<Labels>(coord.Select(P => new Labels 
            { 
                X = P.X * scaleFactor + translateX + offset * scaleFactor, 
                Y = P.Y * scaleFactor + translateY - 3 * offset * scaleFactor, 
                ID = P.ID, 
                Colour = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFDBDBDB"))
            }
            ));

            // Point support
            foreach (var support in pointSupports)
            {
                CanvasLabels.Add(new Labels
                {
                    X = support.X * scaleFactor + translateX + offset * scaleFactor,
                    Y = support.Y * scaleFactor + translateY - 3 * offset * scaleFactor,
                    ID = support.ID,
                    Colour = new SolidColorBrush(Colors.Green)
                });
            }

            // Line supports
            foreach (var support in lineSupports)
            {
                CanvasLabels.Add(new Labels
                {
                    X = (support.X2 + support.X) / 2 * scaleFactor + translateX + offset * scaleFactor,
                    Y = (support.Y2 + support.Y) / 2 * scaleFactor + translateY - 3 * offset * scaleFactor,
                    ID = support.ID,
                    Colour = new SolidColorBrush(Colors.Green)
                });
            }

            // Pointloads
            foreach (var load in PointLoads)
            {
                CanvasLabels.Add(new Labels
                {
                    X = load.X * scaleFactor + translateX + offset * scaleFactor,
                    Y = load.Y * scaleFactor + translateY - 3 * offset * scaleFactor,
                    ID = load.ID,
                    Colour = new SolidColorBrush(Colors.Red)
                });
            }

            foreach (var load in lineLoads)
            {
                CanvasLabels.Add(new Labels
                {
                    X = (load.X2 + load.X) / 2 * scaleFactor + translateX + offset * scaleFactor,
                    Y = (load.Y2 + load.Y) / 2 * scaleFactor + translateY - 3 * offset * scaleFactor,
                    ID = load.ID,
                    Colour = new SolidColorBrush(Colors.Red)
                });
            }

            OnPropertyChanged(nameof(CanvasLabels));
        }
        #endregion

    }
}
