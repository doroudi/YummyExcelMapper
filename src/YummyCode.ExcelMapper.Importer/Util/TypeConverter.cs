using System;
using System.Globalization;
using YummyCode.ExcelMapper.Shared.Extensions;

namespace ExcelMapper.Util
{
    public static class TypeConverter
    {
        public static void SetValue(object source, string propertyName, object valueToSet)
        {
            // find out the type
            Type type = source.GetType();

            var property = type.GetProperty(propertyName);

            // Convert.ChangeType does not handle conversion to nullable types
            // if the property type is nullable, we need to get the underlying type of the property
            Type propertyType = property.PropertyType;
            var targetType = IsNullableType(propertyType) ? Nullable.GetUnderlyingType(propertyType) : propertyType;
            try
            {
                // special case for enums
                if (valueToSet != null && targetType.IsEnum)
                {
                    // we could be going from an int -> enum so specifically let
                    // the Enum object take care of this conversion
                    if (valueToSet != null && valueToSet != null)
                    {
                        valueToSet = Enum.ToObject(targetType, valueToSet);
                    }
                }
                else
                {
                    // returns an System.Object with the specified System.Type and whose value is
                    // equivalent to the specified object.

                    if (valueToSet != null && string.IsNullOrEmpty(valueToSet.ToString()))
                    {
                        valueToSet = null;
                    }
                    else if (valueToSet != null && targetType == typeof(DateTime))
                    {
                        // TODO: improve implementation
                        valueToSet = valueToSet.ToString().Replace(".", "/");
                        valueToSet = valueToSet.ToString().NormalizeDateFormat();
                        string[] formats = { "dd/MM/yyyy", "dd/M/yyyy", "d/M/yyyy", "d/MM/yyyy",
                                "dd/MM/yy", "dd/M/yy", "d/M/yy", "d/MM/yy"};

                        if (valueToSet != null && DateTime.TryParse(valueToSet.ToString(), out DateTime converted))
                        {
                            valueToSet = converted;
                        }
                        else if (valueToSet != null && DateTime.TryParseExact(valueToSet.ToString(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out converted))
                        {
                            valueToSet = converted;
                        }
                        else
                        {
                            valueToSet = null;
                        }
                    }
                    else
                    {
                        var basicType = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;
                        valueToSet = (valueToSet == null) ? null : Convert.ChangeType(valueToSet, basicType);
                    }
                }

                // set the value of the property
                property.SetValue(source, valueToSet);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"cannot set {valueToSet} value because: {ex.Message}");
            }

        }

        private static bool IsNullableType(Type type)
        {
            return type.IsGenericType &&
                type.GetGenericTypeDefinition().Equals(typeof(Nullable<>));
        }
    }
}
