using FemDesign;
using FemDesign.Calculate;
using FemDesign.Drawing;
using FemDesign.GenericClasses;
using FemDesign.Geometry;
using FemDesign.Loads;
using FemDesign.Materials;
using FemDesign.Reinforcement;
using FemDesign.Supports;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using static Dalton_Trapper.ViewModel.FemdesignVM;
using Plane = FemDesign.Geometry.Plane;
using Thickness = FemDesign.Shells.Thickness;

namespace Dalton_Trapper.Model
{
    public class PlateToFEM
    {
        public static void SendToFEM(string plateThickness, string slabName, string materiale, string plateLiveLoad,
                             string psi2q, string extraG, string sk, string skOP, string CC,
                             bool FEMInBg, bool keepFEMOpen, bool RunAnalysis, bool ApplyDocTemplate,
                             ObservableCollection<Entries> coord, ObservableCollection<CurvedLine> curvedLines,
                             ObservableCollection<Entries> pointSup, ObservableCollection<Entries> lineSup,
                             ObservableCollection<Entries> pointLoads, ObservableCollection<Entries> lineLoads,
                             ObservableCollection<Reinforcement> netReinforcement, ObservableCollection<Reinforcement> additionalReinforcement,
                             string projectTitle, string projectNumber, string environmentClass)
        {
            string _plateThickness = plateThickness;
            string _slabName = slabName;
            string _materiale = materiale;
            string _plateLiveLoad = plateLiveLoad;
            string _psi2q = psi2q;
            string _extraG = extraG;
            string _sk = sk;
            string _skOP = skOP;
            string _CC = CC;
            bool _FEMInBg = FEMInBg;
            bool _keepFEMOpen = keepFEMOpen;
            bool _RunAnalysis = RunAnalysis;
            bool _ApplyDocTemplate = ApplyDocTemplate;
            var _coord = coord;
            var _curvedLines = curvedLines;
            var _pointSup = pointSup;
            var _lineSup = lineSup;
            var _pointLoads = pointLoads;
            var _lineLoads = lineLoads;
            var _netReinforcement = netReinforcement;
            var _additionalReinforcement = additionalReinforcement;
            var _project = projectTitle;
            var _sagsnummer = projectNumber;
            var _environmentClass = environmentClass;

            float thickness = float.Parse(_plateThickness) / 1000f;
            float liveload;
            float psi2_q;
            float extraDL;
            float snowload;
            float snowload2;
            float Kfi;
            double w_max;
            string config_wk_filepath;

            if (_CC == "CC3")
            {
                Kfi = 1.1f;
            }
            else
            {
                Kfi = 1.0f;
            }

            if (_plateLiveLoad == "" || _plateLiveLoad == null)
            {
                liveload = 0;
            }
            else if (_plateLiveLoad.Contains('.'))
            {
                liveload = float.Parse(_plateLiveLoad.Replace(".", ","));
            }
            else
            {
                liveload = float.Parse(_plateLiveLoad);
            }

            if (_psi2q == "" || _psi2q == null)
            {
                psi2_q = 1;
            }
            else if (_psi2q.Contains('.'))
            {
                psi2_q = float.Parse(_psi2q.Replace(".", ","));
            }
            else
            {
                psi2_q = float.Parse(_psi2q);
            }

            if (_extraG == "" || _extraG == null)
            {
                extraDL = 0;
            }
            else if (_extraG.Contains('.'))
            {
                extraDL = float.Parse(_extraG.Replace(".", ","));
            }
            else
            {
                extraDL = float.Parse(_extraG);
            }

            if (_sk == "" || _sk == null)
            {
                snowload = 0;
            }
            else if (_sk.Contains('.'))
            {
                snowload = float.Parse(_sk.Replace(".", ","));
            }
            else
            {
                snowload = float.Parse(_sk);
            }

            if (_skOP == "" || _skOP == null)
            {
                snowload2 = 0;
            }
            else if (_skOP.Contains('.'))
            {
                snowload2 = float.Parse(_skOP.Replace(".", ","));
            }
            else
            {
                snowload2 = float.Parse(_skOP);
            }

            if (_environmentClass == "Agressiv")
            {
                w_max = Math.Round(0.30f, 2);
                config_wk_filepath = System.IO.Path.Combine("Model", "cfg_Agressiv.xml");
            }
            else if (_environmentClass == "EkstraAgressiv")
            {
                w_max = Math.Round(0.20f, 2);
                config_wk_filepath = System.IO.Path.Combine("Model", "cfg_EkstraAgressiv.xml");
            }
            else
            {
                w_max = Math.Round(0.40f, 2);
                config_wk_filepath = System.IO.Path.Combine("Model", "cfg_Standard.xml");
            }


#pragma warning disable CS8603 // Possible null reference return.
            Task.Run(() =>
            {
                // Initiate lists of elements and loads for FEM-Design model
                var elements = new List<IStructureElement> { };
                var loads = new List<ILoadElement> { };

                // Define points that make up the plate
                var points = new List<Point3d> { };
                var pointsCurved = new List<Point3d> { };

                // Looping through the collection plate geometry point
                for (int i = 0; i < _coord.Count; i++)
                {
                    int coord_x = _coord[i].X;
                    int coord_y = _coord[i].Y; ;

                    points.Add(new FemDesign.Geometry.Point3d(coord_x / 1000.0, coord_y / 1000.0, 0));
                }

                // Define edges that make up the plate
                var edges = new List<Edge> { };

                for (int i = 0; i < points.Count; i++)  // Looping through point to make up the edges
                {
                    bool added = false;
                    if (i == points.Count - 1) // Last edge has to be connected to first point
                    {
                        for (int h = 0; h < _curvedLines.Count; h++)
                        {
                            double p1x = _curvedLines[h].P1.X / 1000.0;
                            double p1y = _curvedLines[h].P1.Y / 1000.0;
                            double p2x = _curvedLines[h].P2.X / 1000.0;
                            double p2y = _curvedLines[h].P2.Y / 1000.0;
                            double radius = _curvedLines[h].Radius / 1000.0;

                            if ((p1x == points[i].X) && (p1y == points[i].Y))
                            {
                                var midpoint_X = (p1x + p2x) / 2.0;
                                var midpoint_Y = (p1y + p2y) / 2.0;
                                var distToCenter = Math.Sqrt(Math.Pow(midpoint_X - p1x, 2) + Math.Pow(midpoint_Y - p1y, 2));

                                var angle = _curvedLines[h].Form == "Konveks" ? -Math.Asin(distToCenter / radius) : Math.Asin(distToCenter / radius);  //Radians
                                var center_X = _curvedLines[h].Form == "Konveks" ? midpoint_X - Math.Sqrt(Math.Pow(radius, 2) - (Math.Pow(p2x - p1x, 2) + Math.Pow(p2y - p1y, 2)) / 4.0) * (p1y - p2y) / Math.Sqrt(Math.Pow(p2x - p1x, 2) + Math.Pow(p2y - p1y, 2)) :
                                                                                   midpoint_X + Math.Sqrt(Math.Pow(radius, 2) - (Math.Pow(p2x - p1x, 2) + Math.Pow(p2y - p1y, 2)) / 4.0) * (p1y - p2y) / Math.Sqrt(Math.Pow(p2x - p1x, 2) + Math.Pow(p2y - p1y, 2));
                                var center_Y = _curvedLines[h].Form == "Konveks" ? midpoint_Y - Math.Sqrt(Math.Pow(radius, 2) - (Math.Pow(p2x - p1x, 2) + Math.Pow(p2y - p1y, 2)) / 4.0) * (p2x - p1x) / Math.Sqrt(Math.Pow(p2x - p1x, 2) + Math.Pow(p2y - p1y, 2)) :
                                                                                   midpoint_Y + Math.Sqrt(Math.Pow(radius, 2) - (Math.Pow(p2x - p1x, 2) + Math.Pow(p2y - p1y, 2)) / 4.0) * (p2x - p1x) / Math.Sqrt(Math.Pow(p2x - p1x, 2) + Math.Pow(p2y - p1y, 2));
                                var toppoint_X = (p1x - center_X) * Math.Cos(angle) - (p1y - center_Y) * Math.Sin(angle) + center_X;
                                var toppoint_Y = (p1x - center_X) * Math.Sin(angle) + (p1y - center_Y) * Math.Cos(angle) + center_Y;

                                var toppoint = new FemDesign.Geometry.Point3d(toppoint_X, toppoint_Y, 0);
                                edges.Add(new FemDesign.Geometry.Edge(points[i], toppoint, points[0], Plane.XY));
                                added = !added;
                            }
                        }

                        if (added != true)
                        { edges.Add(new FemDesign.Geometry.Edge(points[i], points[0], Plane.XY)); }
                    }
                    else
                    {
                        for (int h = 0; h < _curvedLines.Count; h++)
                        {
                            double p1x = _curvedLines[h].P1.X / 1000.0;
                            double p1y = _curvedLines[h].P1.Y / 1000.0;
                            double p2x = _curvedLines[h].P2.X / 1000.0;
                            double p2y = _curvedLines[h].P2.Y / 1000.0;
                            double radius = _curvedLines[h].Radius / 1000.0;

                            if ((p1x == points[i].X) && (p1y == points[i].Y))
                            {
                                var midpoint_X = (p1x + p2x) / 2.0;
                                var midpoint_Y = (p1y + p2y) / 2.0;
                                var distToCenter = Math.Sqrt(Math.Pow(midpoint_X - p1x, 2) + Math.Pow(midpoint_Y - p1y, 2));

                                var angle = _curvedLines[h].Form == "Konveks" ? -Math.Asin(distToCenter / radius) : Math.Asin(distToCenter / radius);  //Radians
                                var center_X = _curvedLines[h].Form == "Konveks" ? midpoint_X - Math.Sqrt(Math.Pow(radius, 2) - (Math.Pow(p2x - p1x, 2) + Math.Pow(p2y - p1y, 2)) / 4.0) * (p1y - p2y) / Math.Sqrt(Math.Pow(p2x - p1x, 2) + Math.Pow(p2y - p1y, 2)) :
                                                                                   midpoint_X + Math.Sqrt(Math.Pow(radius, 2) - (Math.Pow(p2x - p1x, 2) + Math.Pow(p2y - p1y, 2)) / 4.0) * (p1y - p2y) / Math.Sqrt(Math.Pow(p2x - p1x, 2) + Math.Pow(p2y - p1y, 2));
                                var center_Y = _curvedLines[h].Form == "Konveks" ? midpoint_Y - Math.Sqrt(Math.Pow(radius, 2) - (Math.Pow(p2x - p1x, 2) + Math.Pow(p2y - p1y, 2)) / 4.0) * (p2x - p1x) / Math.Sqrt(Math.Pow(p2x - p1x, 2) + Math.Pow(p2y - p1y, 2)) :
                                                                                   midpoint_Y + Math.Sqrt(Math.Pow(radius, 2) - (Math.Pow(p2x - p1x, 2) + Math.Pow(p2y - p1y, 2)) / 4.0) * (p2x - p1x) / Math.Sqrt(Math.Pow(p2x - p1x, 2) + Math.Pow(p2y - p1y, 2));
                                var toppoint_X = (p1x - center_X) * Math.Cos(angle) - (p1y - center_Y) * Math.Sin(angle) + center_X;
                                var toppoint_Y = (p1x - center_X) * Math.Sin(angle) + (p1y - center_Y) * Math.Cos(angle) + center_Y;

                                var toppoint = new FemDesign.Geometry.Point3d(toppoint_X, toppoint_Y, 0);
                                edges.Add(new FemDesign.Geometry.Edge(points[i], toppoint, points[i + 1], Plane.XY));
                                added = !added;
                            }
                        }

                        if (added != true)
                        { edges.Add(new FemDesign.Geometry.Edge(points[i], points[i + 1], Plane.XY)); }

                    }
                }

                // Define region of plate
                var contour = new FemDesign.Geometry.Contour(edges);
                var region = new FemDesign.Geometry.Region(new List<Contour> { contour }, Plane.XY);

                // Define thickness of plate (maximum of 3 different thicknesses)
                var thicknessList = new List<Thickness>
                            {
                                        new Thickness(points[0], thickness ),
                            };

                //Define properties of plate
                var materialDatabase = FemDesign.Materials.MaterialDatabase.GetDefault();
                var FEM_material = materialDatabase.MaterialByName(_materiale);
                FEM_material.Concrete.gammaC_0 = "1.4"; // Prefabricated safety factors
                FEM_material.Concrete.gammaS_0 = "1.2";
                FEM_material.Concrete.CreepSlc = 3.0;
                FEM_material.Concrete.CreepSlf = 3.0;
                FEM_material.Concrete.CreepSlq = 3.0;
                

                // Define plate
                var slab = FemDesign.Shells.Slab.Plate("Plate1", FEM_material, region, null, null, null, thicknessList);

                // Define supports

                // Looping through collection of point supports
                for (int i = 0; i < _pointSup.Count; i++)
                {
                    int coord_x = _pointSup[i].X;
                    int coord_y = _pointSup[i].Y;

                    Point3d point = new FemDesign.Geometry.Point3d(coord_x / 1000.0, coord_y / 1000.0, 0);
                    PointSupport support = new FemDesign.Supports.PointSupport(point, FemDesign.Releases.Motions.RigidPoint(), FemDesign.Releases.Rotations.Free(), "PS");
                    elements.Add(support);
                }

                // Looping through collection of line supports
                for (int i = 0; i < _lineSup.Count; i++)
                {
                    int coord_x1 = _lineSup[i].X;
                    int coord_y1 = _lineSup[i].Y;
                    int coord_x2 = _lineSup[i].X2;
                    int coord_y2 = _lineSup[i].Y2;

                    Point3d p1 = new FemDesign.Geometry.Point3d(coord_x1 / 1000.0, coord_y1 / 1000.0, 0);
                    Point3d p2 = new FemDesign.Geometry.Point3d(coord_x2 / 1000.0, coord_y2 / 1000.0, 0);
                    Edge line = new FemDesign.Geometry.Edge(p1, p2);
                    LineSupport support = new FemDesign.Supports.LineSupport(line, new FemDesign.Releases.Motions(1e7, 1e7, 1e7, 1e7, 1e7, 0), FemDesign.Releases.Rotations.Free(), false);
                    elements.Add(support);
                }

                //  Define load cases
                var loadCaseDL = new LoadCase("Egenlast - Repos", LoadCaseType.DeadLoad, LoadCaseDuration.Permanent);
                var loadCaseDL_extra = new LoadCase("Fri permanent last", LoadCaseType.Static, LoadCaseDuration.Permanent);
                var loadCaseLL = new LoadCase("Nyttelast - Repos", LoadCaseType.Static, LoadCaseDuration.Permanent);
                var loadCaseSL = new LoadCase("Snelast", LoadCaseType.Static, LoadCaseDuration.Permanent);
                var loadCaseSL_OP = new LoadCase("Snelast - Ophobning", LoadCaseType.Static, LoadCaseDuration.Permanent);
                var loadCaseDL_LL = new LoadCase("Egenlast - Linjelaster", LoadCaseType.Static, LoadCaseDuration.Permanent);
                var loadCaseLL_LL = new LoadCase("Nyttelast - Linjelaster", LoadCaseType.Static, LoadCaseDuration.Permanent);
                var loadCaseDL_PL = new LoadCase("Egenlast - Punktlaster", LoadCaseType.Static, LoadCaseDuration.Permanent);
                var loadCaseLL_PL = new LoadCase("Nyttelast - Punktlaster", LoadCaseType.Static, LoadCaseDuration.Permanent);

                // Define load combination
                var loadCombDL = new LoadCombination("ULS (6.10a) - Dominerende egenlast", LoadCombType.UltimateOrdinary, (loadCaseDL, Kfi * 1.2), (loadCaseDL_LL, Kfi * 1.2), (loadCaseDL_PL, Kfi * 1.2), (loadCaseDL_extra, Kfi * 1.2));
                var loadCombLL = new LoadCombination("ULS (6.10b) - Dominerende nyttelast", LoadCombType.UltimateOrdinary, (loadCaseDL, Kfi * 1.0), (loadCaseDL_extra, Kfi * 1.0), (loadCaseDL_LL, Kfi * 1.0), (loadCaseDL_PL, Kfi * 1.0), (loadCaseLL, Kfi * 1.5), (loadCaseLL_LL, Kfi * 1.5), (loadCaseLL_PL, Kfi * 1.5), (loadCaseSL, Kfi * 0.45));
                var loadCombSL = new LoadCombination("ULS (6.10b) - Sneophobning", LoadCombType.UltimateOrdinary, (loadCaseDL, Kfi * 1.0), (loadCaseSL_OP, Kfi * 1.5));
                var loadCombSLS = new LoadCombination("SLS - Kvasi-permanent", LoadCombType.ServiceabilityQuasiPermanent, (loadCaseDL, 1.0), (loadCaseDL_extra, 1.0), (loadCaseDL_LL, 1.0), (loadCaseDL_PL, 1.0), (loadCaseLL, psi2_q), (loadCaseLL_LL, psi2_q), (loadCaseLL_PL, psi2_q));

                // Define loads //

                // Define surfaceloads on plate
                var force_LL = new Vector3d(0, 0, -liveload);
                var force_DL_extra = new Vector3d(0, 0, -extraDL);
                var force_SL = new Vector3d(0, 0, -snowload);
                var force_SL_OP = new Vector3d(0, 0, -snowload2);

                var surfaceLoad = new FemDesign.Loads.SurfaceLoad(region, force_LL, loadCaseLL);
                var surfaceLoad_DL_extra = new FemDesign.Loads.SurfaceLoad(region, force_DL_extra, loadCaseDL_extra);
                var surfaceLoad_SL = new FemDesign.Loads.SurfaceLoad(region, force_SL, loadCaseSL);
                var surfaceLoad_SL_OP = new FemDesign.Loads.SurfaceLoad(region, force_SL_OP, loadCaseSL_OP);

                loads.Add(surfaceLoad);
                loads.Add(surfaceLoad_DL_extra);
                loads.Add(surfaceLoad_SL);
                loads.Add(surfaceLoad_SL_OP);

                // Looping through collection of point loads
                for (int i = 0; i < _pointLoads.Count; i++)
                {
                    int coord_x = _pointLoads[i].X;
                    int coord_y = _pointLoads[i].Y;
                    float gk = _pointLoads[i].G;
                    float qk = _pointLoads[i].Q;

                    Point3d point = new FemDesign.Geometry.Point3d(coord_x / 1000.0, coord_y / 1000.0, 0);
                    if (gk != 0)
                    {
                        Vector3d forceDL = new Vector3d(0, 0, -gk);
                        PointLoad pointLoadDL = new FemDesign.Loads.PointLoad(point, forceDL, loadCaseDL_PL, "", ForceLoadType.Force);
                        loads.Add(pointLoadDL);
                    }

                    if (qk != 0)
                    {
                        Vector3d forceLL = new Vector3d(0, 0, -qk);
                        PointLoad pointLoadLL = new FemDesign.Loads.PointLoad(point, forceLL, loadCaseLL_PL, "", ForceLoadType.Force);
                        loads.Add(pointLoadLL);
                    }
                }

                // Looping through collection of line loads
                for (int i = 0; i < _lineLoads.Count; i++)
                {
                    int x1 = _lineLoads[i].X;
                    int y1 = _lineLoads[i].Y;
                    int x2 = _lineLoads[i].X2;
                    int y2 = _lineLoads[i].Y2;
                    var gk = _lineLoads[i].G;
                    var qk = _lineLoads[i].Q;

                    Point3d p1 = new FemDesign.Geometry.Point3d(x1 / 1000.0, y1 / 1000.0, 0);
                    Point3d p2 = new FemDesign.Geometry.Point3d(x2 / 1000.0, y2 / 1000.0, 0);
                    Edge line = new FemDesign.Geometry.Edge(p1, p2);

                    if (gk != 0)
                    {
                        Vector3d forceDL = new Vector3d(0, 0, -gk);
                        LineLoad LineLoadDL = new FemDesign.Loads.LineLoad(line, forceDL, loadCaseDL_LL, ForceLoadType.Force);
                        loads.Add(LineLoadDL);
                    }

                    if (qk != 0)
                    {
                        Vector3d forceLL = new Vector3d(0, 0, -qk);
                        LineLoad LineLoadLL = new FemDesign.Loads.LineLoad(line, forceLL, loadCaseLL_LL, ForceLoadType.Force);
                        loads.Add(LineLoadLL);
                    }
                }

                // Combine all loadcases and -combinations
                var loadCases = new List<LoadCase> { loadCaseDL, loadCaseDL_extra, loadCaseLL, loadCaseSL, loadCaseSL_OP, loadCaseDL_LL, loadCaseLL_LL, loadCaseDL_PL, loadCaseLL_PL };
                var loadCombinations = new List<LoadCombination> { loadCombDL, loadCombLL, loadCombSL, loadCombSLS };

                // Define the analysis settings
                var settingsULS = new CombItem();
                var settingsSLS = new CombItem(PL: false, Cr: true);
                var combItems = new List<CombItem> { settingsULS, settingsULS, settingsULS, settingsSLS };
                var comb = new Comb
                {
                    NLEmaxiter = 30,
                    PLdefloadstep = 20,
                    PLminloadstep = 2,
                    PlKeepLoadStep = true,
                    PlTolerance = 1,
                    PLmaxeqiter = 50,
                    PlShellLayers = 10,
                    NLSMohr = true,
                    NLSinitloadstep = 10,
                    NLSminloadstep = 10,
                    NLSactiveelemratio = 5,
                    NLSplasticelemratio = 5,
                    CRloadstep = 20,
                    CRmaxiter = 30,
                    CRstifferror = 2,
                    CombItem = combItems,
                };
                var analysis = Analysis.StaticAnalysis(comb, calcCase: true, calccomb: true);

                // Define reinforcement properties - net
                var srfReinf = new List<SurfaceReinforcement>();
                MaterialDatabase materialeData = MaterialDatabase.GetDefault("DK");

                // Netarmering
                for (int i = 0; i < _netReinforcement.Count; i++)
                {
                    var reinforcement = materialeData.MaterialByName(_netReinforcement[i].Kvalitet);
                    var wire = new Wire(_netReinforcement[i].Diameter / 1000f, reinforcement, WireProfileType.Ribbed);

                    if (i == 0)
                    {
                        var bar = new Straight(ReinforcementDirection.X, _netReinforcement[i].Afstand / 1000f, FemDesign.GenericClasses.Face.Top, _netReinforcement[i].Dæklag / 1000f);
                        var straightReinf = SurfaceReinforcement.DefineStraightSurfaceReinforcement(region, bar, wire);
                        srfReinf.Add(straightReinf);
                    }
                    else if (i == 1)
                    {
                        var bar = new Straight(ReinforcementDirection.Y, _netReinforcement[i].Afstand / 1000f, FemDesign.GenericClasses.Face.Top, _netReinforcement[i].Dæklag / 1000f);
                        var straightReinf = SurfaceReinforcement.DefineStraightSurfaceReinforcement(region, bar, wire);
                        srfReinf.Add(straightReinf);
                    }
                    else if (i == 2)
                    {
                        var bar = new Straight(ReinforcementDirection.X, _netReinforcement[i].Afstand / 1000f, FemDesign.GenericClasses.Face.Bottom, _netReinforcement[i].Dæklag / 1000f);
                        var straightReinf = SurfaceReinforcement.DefineStraightSurfaceReinforcement(region, bar, wire);
                        srfReinf.Add(straightReinf);
                    }
                    else
                    {
                        var bar = new Straight(ReinforcementDirection.Y, _netReinforcement[i].Afstand / 1000f, FemDesign.GenericClasses.Face.Bottom, _netReinforcement[i].Dæklag / 1000f);
                        var straightReinf = SurfaceReinforcement.DefineStraightSurfaceReinforcement(region, bar, wire);
                        srfReinf.Add(straightReinf);
                    }
                }

                // Løse stænger
                for (int i = 0; i < _additionalReinforcement.Count; i++)
                {
                    if (_additionalReinforcement[i].Diameter == 0)
                    {
                        continue;
                    }

                    var reinforcement = materialeData.MaterialByName(_additionalReinforcement[i].Kvalitet);
                    var wire = new Wire(_additionalReinforcement[i].Diameter / 1000f, reinforcement, WireProfileType.Ribbed);

                    if (i == 0)
                    {
                        var bar = new Straight(ReinforcementDirection.X, _additionalReinforcement[i].Afstand / 1000f, FemDesign.GenericClasses.Face.Top, _additionalReinforcement[i].Dæklag / 1000f);
                        var straightReinf = SurfaceReinforcement.DefineStraightSurfaceReinforcement(region, bar, wire);
                        srfReinf.Add(straightReinf);
                    }
                    else if (i == 1)
                    {
                        var bar = new Straight(ReinforcementDirection.Y, _additionalReinforcement[i].Afstand / 1000f, FemDesign.GenericClasses.Face.Top, _additionalReinforcement[i].Dæklag / 1000f);
                        var straightReinf = SurfaceReinforcement.DefineStraightSurfaceReinforcement(region, bar, wire);
                        srfReinf.Add(straightReinf);
                    }
                    else if (i == 2)
                    {
                        var bar = new Straight(ReinforcementDirection.X, _additionalReinforcement[i].Afstand / 1000f, FemDesign.GenericClasses.Face.Bottom, _additionalReinforcement[i].Dæklag / 1000f);
                        var straightReinf = SurfaceReinforcement.DefineStraightSurfaceReinforcement(region, bar, wire);
                        srfReinf.Add(straightReinf);
                    }
                    else
                    {
                        var bar = new Straight(ReinforcementDirection.Y, _additionalReinforcement[i].Afstand / 1000f, FemDesign.GenericClasses.Face.Bottom, _additionalReinforcement[i].Dæklag / 1000f);
                        var straightReinf = SurfaceReinforcement.DefineStraightSurfaceReinforcement(region, bar, wire);
                        srfReinf.Add(straightReinf);
                    }
                }

                var reinfSlab = FemDesign.Reinforcement.SurfaceReinforcement.AddReinforcementToSlab(slab, srfReinf);
                elements.Add(reinfSlab);

               

                // Set up the model
                var model = new FemDesign.Model(Country.DK);
                model.AddElements(elements);
                model.AddLoads(loads);
                model.AddLoadCases(loadCases);
                model.AddLoadCombinations(loadCombinations);

                // Documentation
                string relativePathDocTemplate = System.IO.Path.Combine("Model\\FEM-Design tab", "Repos_Doc_Template.dsc");
                string relativePathDocTemplateDeploy = System.IO.Path.Combine("Model\\FEM-Design tab", "Repos_Doc_Template - Deploy.dsc");
                string filepathDocTemplate = System.IO.Path.GetFullPath(relativePathDocTemplate);

                

                // Adding dimension lines
                int j = 0;
                foreach (Point3d point in points)
                {
                    int nextI = (j + 1) % _coord.Count;

                    Vector3d edgeDirection = points[nextI] - point;
                    edgeDirection.Normalize();
                    Vector3d perpDirection = new Vector3d(-edgeDirection.Y, edgeDirection.X, 0);
                    Plane plane = new Plane(point, edgeDirection, perpDirection);

                    var dim = new DimensionLinear(new List<Point3d> { points[j], points[nextI] }, plane);
                    model.AddLinearDimension(dim, true);
                    j++;
                }

                // create a direct link to FEM-Design
                try
                {
                    using (var femDesign = new FemDesign.FemDesignConnection($@"C:\Program Files\StruSoft\FEM-Design 23\", _FEMInBg))
                    {
                        // Update meshsize and initialize model
                        model.Entities.Slabs[0].SlabPart.MeshSize = thickness;
                        femDesign.Open(model);

                        // Sets the design module to calculate crack widths
                        var config = new FemDesign.Calculate.ConcreteDesignConfig(new ConcreteDesignConfig.CalculationMethod(), reopeningCracks: true);
                        femDesign.SetConfig(config);

                        // Sets maximum allowable crack width based on environment class
                        femDesign.SetConfig(config_wk_filepath);

                        // Documentation information
                        var designer = Environment.UserName.ToUpper();
                        var date = DateTime.Now.ToString("dd/MM/yyyy");

                        // Sets projects information (to be used in the documentation header)
                        femDesign.SetProjDescription(_project, "", designer, "", "", new List<UserDefinedData>{ new UserDefinedData("Project No.:", _sagsnummer), new UserDefinedData("Published date", date) });

                        // Run analysis if selected
                        if (_RunAnalysis == true)
                        {
                            femDesign.RunAnalysis(analysis);

                            // Runs the RC design module for longitudinal and shear reinforcement (not auto-design but check of manual applied reinforcement)
                            Design design = new FemDesign.Calculate.Design(autoDesign: false, check: true, applyChanges: false);
                            
                            femDesign.RunDesign(CmdUserModule.RCDESIGN, design);
                        }

                        // Retrieves pointsupport reactions
                        var pointSupportReactions = femDesign.GetResults<FemDesign.Results.PointSupportReaction>();

                        // Retrieves crack widths
                        var wk = femDesign.GetResults<FemDesign.Results.RCShellCrackWidth>();
                        double maxCrackWidth = 0;

                        foreach (var crack in wk)
                        {
                            if (crack.Width1 > maxCrackWidth) 
                            {
                                maxCrackWidth = crack.Width1;
                            }
                            if (crack.Width2 > maxCrackWidth)
                            {
                                maxCrackWidth = crack.Width2;
                            }
                        }

                        // Applies and updates documentation template
                        if (_ApplyDocTemplate == true)
                        {
                            try
                            {
                                string fileContent = File.ReadAllText(filepathDocTemplate);
                                string crackWdithConclusion = maxCrackWidth <= w_max ? String.Format("Maksimale revnevidde ({0}mm) er mindre end wk = {1}mm. OK!", Math.Round(maxCrackWidth, 2), w_max) : String.Format("Maksimale revnevidde ({0}mm) overstiger wk = {1}mm. IKKE OK!", Math.Round(maxCrackWidth, 2), w_max);

                                if (snowload == 0)
                                {
                                    string updatedContent = fileContent.Replace("Repos XXX", slabName).Replace("CCX", CC).Replace("KFIX", Kfi.ToString()).Replace(" - Sneophobning (6.10b)", "").Replace("RevnEXXX", crackWdithConclusion);
                                    File.WriteAllText(relativePathDocTemplateDeploy, updatedContent);
                                }
                                else
                                {
                                    string updatedContent = fileContent.Replace("Repos XXX", slabName).Replace("CCX", CC).Replace("KFIX", Kfi.ToString()).Replace("RevnEXXX", crackWdithConclusion);
                                    File.WriteAllText(relativePathDocTemplateDeploy, updatedContent);
                                }

                            }
                            catch (Exception ex) { Debug.WriteLine($"Fejl! {ex.Message}"); }

                            femDesign.ApplyDocumentationTemplate(relativePathDocTemplateDeploy);
                        }

                        // Disconnects connection to FEM-Design if selected
                        if (_keepFEMOpen == true)
                        {
                            femDesign.Disconnect();
                        }

                        //return pointSupportReactions;
                    }
                }
                catch
                {
                    MessageBox.Show($"Kan ikke finde installation af FEM-Design", "Fejl!", MessageBoxButton.OK, MessageBoxImage.Error);
                    //return null;
                }
            });
        #pragma warning restore CS8603 // Possible null reference return.
        }
    }
}
