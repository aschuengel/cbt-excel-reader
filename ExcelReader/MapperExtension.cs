using ClosedXML.Excel;

namespace ExcelReader;

public static class MapperExtension
{
    public static List<T> Map<T>(this IXLWorksheet sheet)
    {
        var mappings = ProcessMapping(sheet, typeof(T));
        return ProcessDataRows<T>(sheet, mappings);
    }

    private static List<T> ProcessDataRows<T>(IXLWorksheet sheet, List<Mapping> mappings)
    {
        var items = new List<T>();
        var type = typeof(T);
        var constructor = type.GetConstructor(Array.Empty<Type>()) ?? throw new Exception($"No void constructor for type {type}");
        foreach (var row in sheet.RowsUsed().Where(row => row.RowNumber() > 1))
        {
            var item = (T)constructor.Invoke(Array.Empty<object>());
            for (var column = 0; column < mappings.Count; column++)
            {
                var mapping = mappings[column];
                var value = row.Cell(column + 1).Value;
                if (mapping.Property == null)
                {
                    continue;
                }

                Type propertyType = mapping.Property.PropertyType;
                switch (value.Type)
                {
                    case XLDataType.Text:
                        if (propertyType == typeof(string))
                        {
                            mapping.Property.SetValue(item, value.GetText());
                        }
                        else if (propertyType == typeof(int))
                        {
                            var intValue = Convert.ToInt32(value.GetText());
                            mapping.Property.SetValue(item, intValue);
                        }
                        else if (propertyType == typeof(double))
                        {
                            var doubleValue = Convert.ToDouble(value.GetText());
                            mapping.Property.SetValue(item, doubleValue);
                        }
                        else
                        {
                            throw new MapperException($"Don't know how to handle {mapping.Property} on POCOs, property name is {mapping.Name}");
                        }
                        break;
                    case XLDataType.Number:
                        if (propertyType == typeof(string))
                        {
                            var doubleValue = value.GetNumber();
                            mapping.Property.SetValue(item, doubleValue.ToString());
                        }
                        else if (propertyType == typeof(int))
                        {
                            var doubleValue = value.GetNumber();
                            mapping.Property.SetValue(item, (int)doubleValue);
                        }
                        else if (propertyType == typeof(double))
                        {
                            var doubleValue = value.GetNumber();
                            mapping.Property.SetValue(item, doubleValue);
                        }
                        else
                        {
                            throw new MapperException($"Don't know how to handle {mapping.Property} on POCOs, property name is {mapping.Name}");
                        }
                        break;
                    // TODO: Additional types
                    default:
                        throw new MapperException($"Don't know how to handle cell type {value.Type}");
                }
            }
            items.Add(item);
        }
        return items;
    }

    private static List<Mapping> ProcessMapping(IXLWorksheet sheet, Type type)
    {
        var headerRow = sheet.Row(1);
        var mappings = new List<Mapping>();
        foreach (var headerCell in headerRow.CellsUsed())
        {
            var mapping = new Mapping();
            if (headerCell.TryGetValue<string>(out var headerValue))
            {
                mapping.Name = headerValue;
            }
            else
            {
                throw new MapperException("Values in the header row must be strings");
            }
            var property = type.GetProperty(headerValue);
            if (property == null)
            {
                // TODO: Replace with logger
                Console.Error.WriteLine($"Warning: No matching property in type {type} for header {headerValue}");
                mapping.Property = null;
            }
            else
            {
                mapping.Property = property;
            }
            mappings.Add(mapping);
        }
        return mappings;
    }
}
