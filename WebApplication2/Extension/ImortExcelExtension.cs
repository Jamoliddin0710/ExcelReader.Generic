using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2013.PowerPoint.Roaming;
using WebApplication2.Model.DTOs;

namespace WebApplication2.Extension;

public static class ImortExcelExtension
{
    public static async Task<List<T>> ImportExcel<T>(string excelFilePath, string sheetName)
    {
        Dictionary<string, string> Translators = new Dictionary<string, string>()
        {
            { "Id", "Id" },
            { "FirstName", "FirstName" },
            { "LastName", "LastName" },
            { "Phone", "Phone" },
            { "Country", "Country" },
            { "Region", "Region" }
        };
        List<T> list = new List<T>();
        Type typeofObject = typeof(T);
        using (IXLWorkbook workbook = new XLWorkbook())
        {
            var workSheet = workbook.Worksheets.Where(w => w.Name == sheetName).First();
            var properties = typeofObject.GetProperties();
            var columns = workSheet.FirstRow().Cells().Select((v, i) => new { Value = v.Value, Index = i + 1 });
            foreach (IXLRow row in workSheet.RowsUsed().Skip(1))
            {
                T obj = (T)Activator.CreateInstance(typeofObject);
                foreach (var prop in properties)
                {
                    var header = Translators[prop.Name.ToString()];
                    if (!columns.Any(val => val.ToString() == header))
                    {
                        continue;
                    }
                    
                    int colIndex = columns.SingleOrDefault(c => c.Value.ToString() == header).Index;
                    var val = row.Cell(colIndex).Value;
                    var type = prop.PropertyType;
                    prop.SetValue(obj, Convert.ChangeType(val, type));
                }
                list.Add(obj);
            }
            
            return list;
        }
    }

}