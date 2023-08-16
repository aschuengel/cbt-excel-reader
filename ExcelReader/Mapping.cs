using System.Reflection;

namespace ExcelReader;

internal class Mapping
{
    public string? Name { get; set; }
    public PropertyInfo? Property { get; set; }
}
