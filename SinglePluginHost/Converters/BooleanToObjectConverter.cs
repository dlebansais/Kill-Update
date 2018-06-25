using System;
using System.Globalization;
using System.Windows.Data;

namespace Converters
{
    [ValueConversion(typeof(bool), typeof(object))]
    public class BooleanToObjectConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool BoolValue = (bool)value;
            CompositeCollection CollectionOfItems = parameter as CompositeCollection;

            // Return the first or second object of a collection depending on a bool value.
            return CollectionOfItems[BoolValue ? 1 : 0];
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return null;
        }
    }
}
