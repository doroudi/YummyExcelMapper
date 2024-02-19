using System;
using System.Globalization;

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
                if (targetType.IsEnum)
                {
                    // we could be going from an int -> enum so specifically let
                    // the Enum object take care of this conversion
                    if (valueToSet != null)
                    {
                        valueToSet = Enum.ToObject(targetType, valueToSet);
                    }
                }
                else
                {
                    // returns an System.Object with the specified System.Type and whose value is
                    // equivalent to the specified object.

                    if (string.IsNullOrEmpty(valueToSet.ToString()))
                    {
                        valueToSet = null;
                    }
                    else if (targetType == typeof(DateTime))
                    {
                        // TODO: improve implementation
                        valueToSet = valueToSet.ToString().Replace(".", "/");
                        string[] formats = { "dd/MM/yyyy", "d/MM/yyyy" };
                        if (DateTime.TryParse(valueToSet.ToString(), out DateTime converted))
                        {
                            valueToSet = converted;
                        }
                        else if (DateTime.TryParseExact(valueToSet.ToString(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out converted))
                        {
                            valueToSet = converted;
                        }
                        else
                        {
                            valueToSet = DateTime.MinValue;
                        }
                    }
                    else
                    {
                        valueToSet = Convert.ChangeType(valueToSet, targetType);
                    }
                }

                // set the value of the property
                property.SetValue(source, valueToSet, null);
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
