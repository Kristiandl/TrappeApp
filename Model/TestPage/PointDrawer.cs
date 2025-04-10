using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Media;

namespace Dalton_Trapper.Model.TestPage
{
    public class PointDrawer
    {
        public static GeometryGroup DrawPoints(ObservableCollection<PointViewModel> coordinates, double scaleFactor, double translateX, double translateY)
        {
            var geometries = new GeometryGroup();

            foreach (var point in coordinates)
            {
                // Draw points as ellipse 
                EllipseGeometry ellipse = new EllipseGeometry(new Point(point.X * scaleFactor + translateX, point.Y * scaleFactor + translateY), 30 * scaleFactor, 30 * scaleFactor);

                geometries.Children.Add(ellipse);
            }

            return geometries;
        }
    }
}
