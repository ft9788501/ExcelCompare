using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Windows;
using System.Windows.Data;

namespace ExcelCompare.Utils
{
    public class BooleanConverter : IValueConverter
    {
        public const string INVERT = "Invert";
        public virtual object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (targetType != typeof(bool))
            {
                throw new InvalidOperationException("The target must be a bool.");
            }
            return ConvertToBoolean(value, parameter);
        }

        public virtual object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return value;
        }

        protected bool ConvertToBoolean(object value, object parameter)
        {

            bool? bNullable = value as bool?;
            if (bNullable == null)
            {
                throw new InvalidOperationException("The value can not be null.");
            }
            bool b = bNullable.Value;
            if (parameter?.ToString() == INVERT)
            {
                b = !b;
            }
            return b;
        }
    }
}
