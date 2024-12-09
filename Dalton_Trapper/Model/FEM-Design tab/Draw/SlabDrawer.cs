using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Media;
using LineSegment = System.Windows.Media.LineSegment;
using Point = System.Windows.Point;
using Size = System.Windows.Size;

namespace Dalton_Trapper.Model
{
    public class SlabDrawer
    {
        public static PathGeometry CreateSlabGeometry(ObservableCollection<Entries> coordinates, ObservableCollection<CurvedLine> curvedlines, double scaleFactor, int translateX, int translateY)
        {
            // Create a PathGeometry for the scaled drawing
            PathGeometry pathGeometry = new PathGeometry();
            PathFigure pathFigure = new PathFigure
            {
                StartPoint = new Point(coordinates[0].X, coordinates[0].Y),
                IsClosed = true // Set to true if you want to close the shape
            };

            // Iterate through the user coordinates to add segments
            for (int i = 0; i < coordinates.Count; i++)
            {
                bool added = false;
                Entries currentPoint = coordinates[i];

                foreach (CurvedLine curve in curvedlines)
                {
                    if (curve.P1 != null && curve.Radius != 0 && currentPoint.X == curve.P1.X && currentPoint.Y == curve.P1.Y)
                    {

                        // Close previous line segment
                        LineSegment lineSegment = new LineSegment
                        {
                            Point = new Point(currentPoint.X * scaleFactor, currentPoint.Y * scaleFactor)
                        };
                        pathFigure.Segments.Add(lineSegment);

                        // Draw arc segment
                        Point P1 = new Point(curve.P1.X * scaleFactor, curve.P1.Y * scaleFactor);
                        Point P2 = new Point(curve.P2.X * scaleFactor, curve.P2.Y * scaleFactor);
                        ArcSegment segment = DrawArc(P1, P2, curve.Radius * scaleFactor, curve.Form);
                        pathFigure.Segments.Add(segment);
                        added = true;
                    }
                }

                if (added != true)
                {
                    // Add either a LineSegment or ArcSegment depending on your needs
                    LineSegment segment = new LineSegment
                    {
                        Point = new Point(currentPoint.X * scaleFactor, currentPoint.Y * scaleFactor)
                    };

                    pathFigure.Segments.Add(segment);
                }
            }

            // Add the PathFigure to the PathGeometry
            pathGeometry.Figures.Add(pathFigure);

            // Apply translation to the entire geometry
            TranslateTransform translate = new TranslateTransform(translateX, translateY);
            pathGeometry.Transform = translate;

            return pathGeometry;
        }

        private static ArcSegment DrawArc(Point startPoint, Point endPoint, double radius, string form)
        {
            // Calculate the midpoint between the start and end points
            Point midPoint = new Point((startPoint.X + endPoint.X) / 2, (startPoint.Y + endPoint.Y) / 2);

            // Calculate the distance between the start and end points
            double distance = Math.Sqrt(Math.Pow(endPoint.X - startPoint.X, 2) + Math.Pow(endPoint.Y - startPoint.Y, 2));

            // Ensure the radius is large enough to form a valid arc with the given start and end points
            if (radius < distance / 2)
            {
                MessageBox.Show("Radius er for lille til at kunne lave et buet linjestykke", "Fejl!", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            // Calculate the arc's perpendicular distance from the midpoint to the circle's center
            double offset = Math.Sqrt(Math.Pow(radius, 2) - Math.Pow(distance / 2, 2));

            // Determine the unit vector perpendicular to the line segment between start and end
            double dx = (endPoint.Y - startPoint.Y) / distance;
            double dy = (endPoint.X - startPoint.X) / distance;

            // Calculate the potential centers of the circle on either side of the line segment
            Point chosenCenter = form == "Konkav" ? new Point(midPoint.X + dx * offset, midPoint.Y - dy * offset) : new Point(midPoint.X - dx * offset, midPoint.Y + dy * offset);

            // Calculate the Size for the ArcSegment
            Size arcSize = new Size(radius, radius);

            // Determine if this is a large arc (we choose false for a smaller arc)
            bool isLargeArc = false;

            // Set sweep direction (Clockwise or Counterclockwise)
            SweepDirection sweepDirection = form == "Konkav" ? SweepDirection.Clockwise : SweepDirection.Counterclockwise;

            // Create the ArcSegment
            ArcSegment arcSegment = new ArcSegment
            {
                Point = endPoint,         // End point of the arc
                Size = arcSize,           // Size (radius) of the arc
                IsLargeArc = isLargeArc,  // Set to false for a smaller arc
                SweepDirection = sweepDirection
            };

            return arcSegment;
        }
    }
}
