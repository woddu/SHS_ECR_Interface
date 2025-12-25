using System.Globalization;
using System.Windows.Data;

namespace WpfSample;
class IndexToActivityLabelConverter : IValueConverter {
  public object Convert(object value, Type targetType, object parameter, CultureInfo culture) {
    if (value is int index) {
      return $"#{index + 1}";
    }
    return "";
  }

  public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) {
    throw new NotImplementedException();
  }
}

