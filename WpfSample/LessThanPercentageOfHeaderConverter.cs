using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;

namespace WpfSample;

public class LessThanPercentageOfHeaderConverter : IValueConverter {
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture) {
        if (value is string cellValue && 
            parameter is string headerStr && 
            double.TryParse(headerStr, out double headerVal)) {

            double cellVal = 0; 
            if (!string.IsNullOrWhiteSpace(cellValue)) { 
                double.TryParse(cellValue, NumberStyles.Float, CultureInfo.InvariantCulture, out cellVal); 
            }

            double threshold = headerVal * 0.75;

            var defaultBrush = (SolidColorBrush)Application.Current.Resources["SystemControlBackgroundAltHighBrush"]; 
            var baseColor = defaultBrush.Color; 
            if (cellVal < threshold) { // Blend with red (simple average of RGB channels) 
                var redOverlay = Colors.Red; 
                byte r = (byte)((baseColor.R + redOverlay.R) / 2); 
                byte g = (byte)((baseColor.G + redOverlay.G) / 2); 
                byte b = (byte)((baseColor.B + redOverlay.B) / 2); 
                var reddishColor = Color.FromRgb(r, g, b); 
                return new SolidColorBrush(reddishColor); 
            } // Above threshold → keep default theme background return defaultBrush; 
            return DependencyProperty.UnsetValue;
        } 
        // Fallback if parsing fails 
        return DependencyProperty.UnsetValue;

    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) {
        throw new NotImplementedException();
    }
}
