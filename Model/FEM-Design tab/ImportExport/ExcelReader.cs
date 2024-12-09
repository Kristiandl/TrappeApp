using System.Collections.ObjectModel;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Dalton_Trapper.Model.ImportExport
{
    public class ExcelReader
    {
        public static TransferData ReadSpecificCells(string filePath)
        {
            var cellData = new Dictionary<string, string>();
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            Excel.Worksheet worksheetFF = null;
            var slabData = new TransferData { };

            try
            {
                excelApp = new Excel.Application();
                workbook = excelApp.Workbooks.Open(filePath);
                worksheet = (Excel.Worksheet)workbook.Sheets[1]; // Access the first worksheet
                worksheetFF = (Excel.Worksheet)workbook.Sheets[3]; // Access the first worksheet
                bool? bundLængst;
                int ff = 0;
                Entries P3;
                Entries P4;
                Entries P5;
                Entries P8;
                Entries P9;
                Entries P10;

                var cellAddresses = new List<string> { "C6", "O77", "I9", "I11", "J22", "J14", "H41", "H51", "J16", "J17", "J18", "H44", "J35", "J36", "L35", "L36", "H48", "H49",
                                                       "H50", "H51", "H52", "H56", "H57", "H58", "J48", "J49", "J50", "J51", "J52", "J56", "J57", "J58", "N22", "P22", "Q11" };
                var cellAddressesSheetFF = new List<string> { "J12", "N7" };

                foreach (var address in cellAddresses)
                {
                    var cell = worksheet.Range[address];
                    cellData[address] = cell.Value?.ToString() ?? string.Empty; // Get the cell value
                }

                foreach (var address in cellAddressesSheetFF)
                {
                    var cell = worksheetFF.Range[address];
                    if (cell.Value?.ToString() == "-2146826246")
                    {
                        cellData[address] = "fejl";
                    }
                    else
                    {
                        cellData[address] = cell.Value?.ToString() ?? string.Empty; // Get the cell value
                    }
                }

                int length = Int32.Parse(cellData["J16"]);
                int width = Int32.Parse(cellData["J17"]);
                int udsparing = Int32.Parse(cellData["J18"]);

                int spanLoadRight = (int)(float.Parse(cellData["J48"]) * 1000);
                int widthLoadRight = (int)(float.Parse(cellData["J49"]) * 1000);
                int distToEndRight = (int)(float.Parse(cellData["J50"]) * 1000);
                int distToPLRight = (int)(float.Parse(cellData["J56"]) * 1000);
                float deadLoadLLRight = float.Parse(cellData["J51"]) * spanLoadRight / 1000f * 0.5f;
                float liveLoadLLRight = float.Parse(cellData["J52"]) * spanLoadRight / 1000f * 0.5f;

                int spanLoadLeft = (int)(float.Parse(cellData["H48"]) * 1000);
                int widthLoadLeft = (int)(float.Parse(cellData["H49"]) * 1000);
                int distToEndLeft = (int)(float.Parse(cellData["H50"]) * 1000);
                int distToPLLeft = (int)(float.Parse(cellData["H56"]) * 1000);
                float deadLoadLLLeft = float.Parse(cellData["H51"]) * spanLoadLeft / 1000f * 0.5f;
                float liveLoadLLLeft = float.Parse(cellData["H52"]) * spanLoadLeft / 1000f * 0.5f;

                // Coordinates - repos
                Entries P1 = new Entries { X = 0, Y = 0 };
                Entries P2 = new Entries { X = 0, Y = length };

                if (cellData["J12"] == "fejl")
                {
                    bundLængst = null;
                    ff = 0;
                }
                else if (cellData["N7"] == "ja")
                {
                    bundLængst = true;
                    ff = Int32.Parse(cellData["J12"]);
                }
                else
                {
                    bundLængst = false;
                    ff = Int32.Parse(cellData["J12"]);
                }

                if (bundLængst == false)
                {
                    P3 = new Entries { X = width, Y = length };
                    P4 = new Entries { X = width, Y = length - distToEndRight };
                    P5 = new Entries { X = width, Y = length - distToEndRight - widthLoadRight };
                    P8 = new Entries { X = width - ff, Y = distToEndRight + widthLoadLeft };
                    P9 = new Entries { X = width - ff, Y = distToEndLeft };
                    P10 = new Entries { X = width - ff, Y = 0 };

                }
                else
                {
                    P3 = new Entries { X = width - ff, Y = length };
                    P4 = new Entries { X = width - ff, Y = length - distToEndRight };
                    P5 = new Entries { X = width - ff, Y = length - distToEndRight - widthLoadRight };
                    P8 = new Entries { X = width, Y = distToEndRight + widthLoadLeft };
                    P9 = new Entries { X = width, Y = distToEndLeft };
                    P10 = new Entries { X = width, Y = 0 };
                }

                Entries P6 = new Entries { X = width - Math.Max(udsparing, ff), Y = length - distToEndRight - widthLoadRight };
                Entries P7 = new Entries { X = width - Math.Max(udsparing, ff), Y = distToEndRight + widthLoadLeft };

                // Point supports
                Entries PS1 = new Entries { X = Int32.Parse(cellData["J35"]), Y = 0 };
                Entries PS2 = new Entries { X = width - Int32.Parse(cellData["J36"]), Y = 0 };
                Entries PS3 = new Entries { X = Int32.Parse(cellData["L35"]), Y = length };
                Entries PS4 = new Entries { X = width - Int32.Parse(cellData["L36"]), Y = length };

                // Pointloads
                slabData.PointLoads = new List<Entries> { };

                if (float.Parse(cellData["H57"]) != 0f && float.Parse(cellData["H58"]) != 0f)
                {
                    Entries pointLoadLeft = new Entries { X = P8.X, Y = P8.Y - distToPLLeft, G = float.Parse(cellData["H57"]), Q = float.Parse(cellData["H58"]) };
                    slabData.PointLoads.Add(pointLoadLeft);
                }
                if (float.Parse(cellData["J57"]) != 0f && float.Parse(cellData["J58"]) != 0f)
                {
                    Entries pointLoadRight = new Entries { X = P4.X, Y = P5.Y + distToPLRight, G = float.Parse(cellData["J57"]), Q = float.Parse(cellData["J58"]) };
                    slabData.PointLoads.Add(pointLoadRight);
                }


                // Lineloads
                slabData.LineLoads = new List<Entries> { };

                if (deadLoadLLRight != 0 || liveLoadLLRight != 0)
                {
                    Entries lineLoadRight = new Entries { X = P4.X, Y = P4.Y, X2 = P5.X, Y2 = P5.Y, G = deadLoadLLRight, Q = liveLoadLLRight };
                    slabData.LineLoads.Add(lineLoadRight);
                }
                if (deadLoadLLLeft != 0 || liveLoadLLLeft != 0)
                {
                    Entries lineLoadLeft = new Entries { X = P8.X, Y = P8.Y, X2 = P9.X, Y2 = P9.Y, G = deadLoadLLLeft, Q = liveLoadLLLeft };
                    slabData.LineLoads.Add(lineLoadLeft);
                }

                // Concrete strenght
                if (cellData["I11"] == "A")
                {
                    slabData.Fck = $"C{cellData["J22"]}/45";
                }
                else if (cellData["I11"] == "E")
                {
                    slabData.Fck = $"C{cellData["J22"]}/50";
                }
                else
                {
                    slabData.Fck = $"C{cellData["J22"]}/37";
                }

                // Coordinates to slabData class
                //bool noLLLoadRight = (deadLoadLLRight == 0 && liveLoadLLRight == 0);
                //bool noLLLoadLeft = (deadLoadLLLeft == 0 && liveLoadLLLeft == 0);

                slabData.Coordinates = new List<Entries> { P1, P2, P3, P4, P5, P6, P7, P8, P9, P10 };
                //if (distToEndRight > 0 && (deadLoadLLRight != 0 || liveLoadLLRight != 0))
                //{
                //    slabData.Coordinates.Add(P4);
                //}
                //if (!noLLLoadRight)
                //{
                //    slabData.Coordinates.Add(P5);
                //}
                //if (udsparing > 0 || ff > 0)
                //{
                //    slabData.Coordinates.Add(P6);
                //    slabData.Coordinates.Add(P7);
                //}
                //if (!noLLLoadLeft)
                //{
                //    slabData.Coordinates.Add(P8);
                //}
                //if (distToEndLeft > 0 && (deadLoadLLLeft != 0 || liveLoadLLLeft != 0))
                //{
                //    slabData.Coordinates.Add(P9);
                //}


                // Filter the list to only keep unique coordinates
                var uniqueCoordinates = new HashSet<(int, int)>();
                slabData.Coordinates = slabData.Coordinates.Where(e => uniqueCoordinates.Add((e.X, e.Y))).ToList();

                // Data stored in slabData class
                slabData.Name = cellData["C6"];
                slabData.PointSupports = new List<Entries> { PS1, PS2, PS3, PS4 };
                slabData.LineSupports = new List<Entries> { };
                slabData.Thickness = cellData["J14"];
                slabData.Q = cellData["H44"];
                slabData.Psi2q = cellData["O77"];
                slabData.ExtraG = cellData["H41"];
                slabData.Snow = null;
                slabData.Snow2 = null;
                slabData.CC = $"CC{cellData["I9"]}";
                slabData.RunInBgYes = false;
                slabData.RunInBgNo = true;
                slabData.DisconnectYes = false;
                slabData.DisconnectNo = true;
                slabData.RunAnalysisYes = true;
                slabData.RunAnalysisNo = false;
                slabData.ApplyDocTemplateYes = false;
                slabData.ApplyDocTemplateNo = true;
                slabData.NetReinforcement = new List<Reinforcement>
                {
                    new Reinforcement {Navn = "x_top", Diameter = Int32.Parse(cellData["N22"]), Afstand = Int32.Parse(cellData["P22"]), Dæklag = Int32.Parse(cellData["Q11"])+Int32.Parse(cellData["N22"]), Kvalitet = "K"},
                    new Reinforcement {Navn = "y_top", Diameter = Int32.Parse(cellData["N22"]), Afstand = Int32.Parse(cellData["P22"]), Dæklag = Int32.Parse(cellData["Q11"]), Kvalitet = "K"},
                    new Reinforcement {Navn = "x_bund", Diameter = Int32.Parse(cellData["N22"]), Afstand = Int32.Parse(cellData["P22"]), Dæklag = Int32.Parse(cellData["Q11"])+Int32.Parse(cellData["N22"]), Kvalitet = "K"},
                    new Reinforcement {Navn = "y_bund", Diameter = Int32.Parse(cellData["N22"]), Afstand = Int32.Parse(cellData["P22"]), Dæklag = Int32.Parse(cellData["Q11"]), Kvalitet = "K"}
                };
            }
            catch (Exception ex)
            {
                throw new Exception("Error reading Excel file: " + ex.Message);
            }
            finally
            {
                // Clean up
                if (worksheet != null)
                {
                    Marshal.ReleaseComObject(worksheet);
                    Marshal.ReleaseComObject(worksheetFF);
                }
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return slabData; // Return a dictionary of cell addresses and their values
        }
    }
}
