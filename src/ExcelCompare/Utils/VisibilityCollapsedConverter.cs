using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Data;

namespace ExcelCompare.Utils
{
    public class VisibilityCollapsedConverter : BooleanConverter
    {
        public override object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (targetType != typeof(Visibility)) 
            {
                throw new InvalidOperationException("The target must be a Visibility.");
            }
            var b = ConvertToBoolean(value, parameter);
            return b ? Visibility.Visible : Visibility.Collapsed;
        }
    }
}
