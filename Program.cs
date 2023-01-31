// string? chapter = "1";
// string? point = null;
// string? subpoint = "1.2.3";

// string?[] values = new string?[] {chapter, point, subpoint};

// string message = String.Format("В главе {0}, подпункте {2} внесены изменения", values);

// Console.WriteLine(message);


// namespace filterActivity
// {
//     public class filterRow
//     {
//         public string columnName;
//         //public string 
//     }
//     class Programm
//     {

//     }
// }

using System.Linq.Expressions;
using System.Data;
using RpaHelper;

internal class Program
{
    private static void Main(string[] args)
    {
        //Read Excel
        string pathToExcel = @"C:\Users\oxy2c\Documents\Test.xlsx";
        string sheetName = "Sheet1";
        DataTable myDt = Excel.ReadWorksheet(pathToExcel, sheetName, true);
        
        pathToExcel = @"C:\Users\oxy2c\Documents\Test2.xlsm";
        DataTable myDt2 = Excel.ReadWorksheet(pathToExcel, 0, true);

        //Expression trees part
        string? startsWith = null;
        string? endsWith = "y";

        Expression<Func<string, bool>> expr = (startsWith, endsWith) switch
        {
            ("" or null, "" or null) => x => true,
            (_, "" or null) => x => x.StartsWith(startsWith),
            ("" or null, _) => x => x.EndsWith(endsWith),
            (_, _) => x => x.StartsWith(startsWith) || x.EndsWith(endsWith)
        };
        
        string[] companyNames = new string[] {
            "Consolidated Messenger", "Alpine Ski House", "Southridge Video",
            "City Power & Light", "Coho Winery", "Wide World Importers",
            "Graphic Design Institute", "Adventure Works", "Humongous Insurance",
            "Woodgrove Bank", "Margie's Travel", "Northwind Traders",
            "Blue Yonder Airlines", "Trey Research", "The Phone Company",
            "Wingtip Toys", "Lucerne Publishing", "Fourth Coffee"
        };

        IQueryable<string> companyNamesSource = companyNames.AsQueryable();

        var qry = companyNamesSource.Where(expr);
        Console.WriteLine("");
    }
}