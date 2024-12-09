using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Media;
using Point = System.Windows.Point;

namespace Dalton_Trapper.Model.Draw
{
    public class DimensionLineDrawer
    {

        // Helper methods to determine which side the dimension line is on
        private static bool IsRightSide(Point p1, Point p2) => p1.X == p2.X && p1.Y > p2.Y;
        private static bool IsLeftSide(Point p1, Point p2) => p1.X == p2.X && p1.Y < p2.Y;

        
        public static bool? IsMoreHorizontal(Point p1, Point p2)
        {
            // Calculate the absolute differences in x and y coordinates
            double deltaX = Math.Abs(p2.X - p1.X);
            double deltaY = Math.Abs(p2.Y - p1.Y);

            // Compare the differences
            if (deltaX >= deltaY)
            {
                return true; // More horizontal or diagonal
            }
            else
            {
                return false; // More vertical
            }
        }


        // Method to draw dimensions
        public static (GeometryGroup, ObservableCollection<Labels>) DrawCanvasDimensionlines(ObservableCollection<Entries> coordinates, double scaleFactor, int translateX, int translateY)
        {
            GeometryGroup geometry = new GeometryGroup();
            ObservableCollection<Labels> labels = new ObservableCollection<Labels>();

            // Determine the bounding box of the shape
            double minX = coordinates.Min(c => c.X * scaleFactor + translateX);
            double minY = coordinates.Min(c => c.Y * scaleFactor + translateY);
            double maxX = coordinates.Max(c => c.X * scaleFactor + translateX);
            double maxY = coordinates.Max(c => c.Y * scaleFactor + translateY);

            for (int i = 0; i < coordinates.Count; i++)
            {
                int nextIndex = (i + 1) % coordinates.Count;
                Point point1 = new Point(coordinates[i].X * scaleFactor + translateX,
                                         coordinates[i].Y * scaleFactor + translateY);
                Point point2 = new Point(coordinates[nextIndex].X * scaleFactor + translateX,
                                         coordinates[nextIndex].Y * scaleFactor + translateY);

                Point point1Annotation = new Point(coordinates[i].X * scaleFactor + translateX,
                                         coordinates[i].Y * scaleFactor + translateY);
                Point point2Annotation = new Point(coordinates[nextIndex].X * scaleFactor + translateX,
                                         coordinates[nextIndex].Y * scaleFactor + translateY);

                // Prevents the program from crashing in the situation before an input is entered.
                // (Before anything is entered the program reads 0, 0 as inputs and connecting this to the startpoint makes it crash)
                if (point1 == point2)
                {
                    break;                
                }

                // Calculate offset for dimension lines
                Vector edgeDirection = point2 - point1;
                edgeDirection.Normalize();
                Vector offset = new Vector(-edgeDirection.Y, edgeDirection.X) * 50; // Perpendicular to edge

                // Places dimensioning lines together on the left and right side
                if (IsRightSide(point1, point2))
                {
                    double dist_to_shift = maxX - point1.X;
                    point1.Offset(offset.X + dist_to_shift, offset.Y);
                    point2.Offset(offset.X + dist_to_shift, offset.Y);
                }
                else if (IsLeftSide(point1, point2))
                {
                    double dist_to_shift = minX - point1.X;
                    point1.Offset(offset.X + dist_to_shift, offset.Y);
                    point2.Offset(offset.X + dist_to_shift, offset.Y);
                }
                else
                {
                    point1.Offset(offset.X, offset.Y);
                    point2.Offset(offset.X, offset.Y);
                }

                // Draw dimension line
                // Rotates edge direction 90 degrees
                Vector direction = new Vector(-edgeDirection.Y, edgeDirection.X);

                // Rotates 45 degree for cross line in beginning and end signature
                double angle = 45;
                RotateTransform rotateTransform_start = new RotateTransform(angle, point1.X, point1.Y);
                RotateTransform rotateTransform_end = new RotateTransform(angle, point2.X, point2.Y);

                // Main line
                LineGeometry line1 = new LineGeometry(new Point(point1.X, point1.Y), 
                                                      new Point(point2.X, point2.Y));
                
                // Start signature
                LineGeometry line2 = new LineGeometry(new Point((point1 - direction * 10).X, (point1 - direction * 10).Y),
                                                      new Point((point1 + direction * 10).X, (point1 + direction * 10).Y));

                LineGeometry line3 = new LineGeometry(new Point((point1 - direction * 12).X, (point1 - direction * 12).Y),
                                                      new Point((point1 + direction * 12).X, (point1 + direction * 12).Y));


                // End signature
                LineGeometry line4 = new LineGeometry(new Point((point2 - direction * 10).X, (point2 - direction * 10).Y),
                                                      new Point((point2 + direction * 10).X, (point2 + direction * 10).Y));
                
                LineGeometry line5 = new LineGeometry(new Point((point2 - direction * 12).X, (point2 - direction * 12).Y),
                                                      new Point((point2 + direction * 12).X, (point2 + direction * 12).Y));


                line3.Transform = rotateTransform_start;
                line5.Transform = rotateTransform_end;

                geometry.Children.Add(line1);
                geometry.Children.Add(line2);
                geometry.Children.Add(line3);
                geometry.Children.Add(line4);
                geometry.Children.Add(line5);

                // Draw dimension label
                double length = Math.Sqrt(Math.Pow(coordinates[nextIndex].X - coordinates[i].X, 2) + Math.Pow(coordinates[nextIndex].Y - coordinates[i].Y, 2));
                Point midPoint = new Point((point1.X + point2.X) / 2, (point1.Y + point2.Y) / 2);

                // Determine angle between points
                double deltaX = Math.Abs(point2.X - point1.X);
                double deltaY = Math.Abs(point2.Y - point1.Y);
                //double angle2 = Math.Atan2(deltaY, deltaX) * 180 / Math.PI;
                double angle2 = Vector.AngleBetween(edgeDirection, new Vector(1, 0));

                // Position the text on the outside of the dimension line
                if (IsRightSide(point1, point2))
                {
                    labels.Add(new Labels
                    {
                        X = midPoint.X + 200 * scaleFactor,
                        Y = midPoint.Y + 60 * scaleFactor,
                        ID = length.ToString("F0"),
                        Colour = new SolidColorBrush(Colors.Yellow),
                        Angle = -angle2,
                    });
                }
                else if (IsLeftSide(point1, point2))
                {
                    labels.Add(new Labels
                    {
                        X = midPoint.X - 250 * scaleFactor,
                        Y = midPoint.Y - 100 * scaleFactor,
                        ID = length.ToString("F0"),
                        Colour = new SolidColorBrush(Colors.Yellow),
                        Angle = -angle2,
                    });
                }
                else if (point1Annotation.Y == maxY && point2Annotation.Y == maxY)  // Top line
                {
                    labels.Add(new Labels
                    {
                        X = midPoint.X - 190 * scaleFactor,
                        Y = midPoint.Y + 180 * scaleFactor,
                        ID = length.ToString("F0"),
                        Colour = new SolidColorBrush(Colors.Yellow),
                        Angle = angle2,
                    });
                }
                else if (point1Annotation.Y == minY && point2Annotation.Y == minY)  // Bottom line
                {
                    labels.Add(new Labels
                    {
                        X = midPoint.X - 190 * scaleFactor,
                        Y = midPoint.Y + 180 * scaleFactor,
                        ID = length.ToString("F0"),
                        Colour = new SolidColorBrush(Colors.Yellow),
                        Angle = 180 + angle2,
                    });
                }
                else if ((bool)IsMoreHorizontal(point1, point2) == false && point1.Y < point2.Y) 
                {
                    labels.Add(new Labels
                    {
                        X = midPoint.X - 300 * scaleFactor,
                        Y = midPoint.Y,
                        ID = length.ToString("F0"),
                        Colour = new SolidColorBrush(Colors.Yellow),
                        Angle = -angle2,
                    });
                }
                else if ((bool)IsMoreHorizontal(point1, point2) == false && point1.Y > point2.Y)
                {
                    labels.Add(new Labels
                    {
                        X = midPoint.X + 240 * scaleFactor,
                        Y = midPoint.Y,
                        ID = length.ToString("F0"),
                        Colour = new SolidColorBrush(Colors.Yellow),
                        Angle = -angle2,
                    });
                }
                else if ((bool)IsMoreHorizontal(point1, point2) == true && point1.X > point2.X)
                {
                    labels.Add(new Labels
                    {
                        X = midPoint.X - 100 * scaleFactor,
                        Y = midPoint.Y - 100 * scaleFactor,
                        ID = length.ToString("F0"),
                        Colour = new SolidColorBrush(Colors.Yellow),
                        Angle = 180 - angle2,
                    });
                }
                else if ((bool)IsMoreHorizontal(point1, point2) == true && point1.X < point2.X)
                {
                    labels.Add(new Labels
                    {
                        X = midPoint.X - 50 * scaleFactor,
                        Y = midPoint.Y + 250 * scaleFactor,
                        ID = length.ToString("F0"),
                        Colour = new SolidColorBrush(Colors.Yellow),
                        Angle = -angle2,
                    });
                }
            }

            return (geometry, labels);
        }
    }
}
