using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using OfficeOpenXml;
using System.Text.RegularExpressions;
namespace DimDb
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ConnectionDB.Open(@"Server=BAYRAMOV\SQLEXPRESS;Database=Muqavile;User Id=sa;Password=123");
            ExcelRead excelRead = new ExcelRead();
            
            excelRead.ReadandWrite();

            

            Console.ReadKey();
        }
    }
}
