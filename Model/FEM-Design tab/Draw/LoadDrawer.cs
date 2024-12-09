using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Media;

namespace Dalton_Trapper.Model.Draw
{
    public class LoadDrawer
    {
        // Method to draw point loads
        public static GeometryGroup DrawPointLoad(ObservableCollection<Entries> pointload, double scaleFactor, int transX, int transY, double size)
        {
            var geometries = new GeometryGroup();

            foreach (var load in pointload)
            {
                // First diagonal line (top-left to bottom-right)
                Point P1 = new Point((load.X - size / 2) * scaleFactor + transX, (load.Y - size / 2) * scaleFactor + transY);
                Point P2 = new Point((load.X + size / 2) * scaleFactor + transX, (load.Y + size / 2) * scaleFactor + transY);
                geometries.Children.Add(new LineGeometry(P1, P2));

                // Second diagonal line (bottom-left to top-right)
                Point P3 = new Point((load.X - size / 2) * scaleFactor + transX, (load.Y + size / 2) * scaleFactor + transY);
                Point P4 = new Point((load.X + size / 2) * scaleFactor + transX, (load.Y - size / 2) * scaleFactor + transY);
                geometries.Children.Add(new LineGeometry(P3, P4));
            }

            return geometries;
        }

        // Method to draw line loads
        public static GeometryGroup DrawLineLoad(ObservableCollection<Entries> lineloads, double scaleFactor, int translateX, int translateY)
        {
            var geometries = new GeometryGroup();

            foreach (var line in lineloads)
            {
                Point P1 = new Point(line.X * scaleFactor + translateX, line.Y * scaleFactor + translateY);
                Point P2 = new Point(line.X2 * scaleFactor + translateX, line.Y2 * scaleFactor + translateY);
                geometries.Children.Add(new LineGeometry(P1, P2));
            }

            return geometries;
        }
    }
}
