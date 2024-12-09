using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Media;

namespace Dalton_Trapper.Model.Draw
{
    public class SupportDrawer
    {
        // Method to draw point loads
        public static GeometryGroup DrawPointSupport(ObservableCollection<Entries> pointsupport, double scaleFactor, int transX, int transY, double size)
        {
            var geometries = new GeometryGroup();
            
            // Size of the square 
            double height = size * scaleFactor;  
            double width = height;

            foreach (var support in pointsupport)
            {
                Point origin = new Point(support.X * scaleFactor - width / 2 + transX, support.Y * scaleFactor - height / 2 + transY);
                geometries.Children.Add(new RectangleGeometry(new Rect(origin, new Size(height, width))));
            }

            return geometries;
        }

        // Method to draw line loads
        public static GeometryGroup DrawLineSupport(ObservableCollection<Entries> linesupport, double scaleFactor, int translateX, int translateY)
        {
            var geometries = new GeometryGroup();

            foreach (var line in linesupport)
            {
                Point P1 = new Point(line.X * scaleFactor + translateX, line.Y * scaleFactor + translateY);
                Point P2 = new Point(line.X2 * scaleFactor + translateX, line.Y2 * scaleFactor + translateY);
                geometries.Children.Add(new LineGeometry(P1, P2));
            }

            return geometries;
        }
    }
}
