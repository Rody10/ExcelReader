using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Storage;


namespace ExcelReader
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");
            string excelFilePath = @"C:\Users\Rodney\source\repos\ExcelReader\ExcelReader\employees.xlsx";

            // check if database exists. If it does not exist, create it. If it exists, delete it and create it again.
            using (var context = new EFCoreEmployeeContext())
            {
                bool databaseExists = context.Database.CanConnect();
                if (databaseExists)
                {
                    Console.WriteLine("Database exists");
                    Console.WriteLine("Deleting database");
                    context.Database.EnsureDeleted();
                    Console.WriteLine("Database deleted!");
                }
                else
                {
                    Console.WriteLine("Database does not exist.");
                   
                }
                Console.WriteLine("-----------------------------");
                Console.WriteLine("Creating the database!");
                context.Database.Migrate();
                Console.WriteLine("Created the database!");

                DatabaseOperations.ReadFromExcelAndWriteToDatabase(context, excelFilePath);
                DatabaseOperations.ReadFromDatabase(context);
            }
            

        }
    }
}
