using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ExcelReader
{
    internal static class DatabaseOperations
    {
        public static void ReadFromExcelAndWriteToDatabase(EFCoreEmployeeContext context, string pathOfExcelFile)
        {
            FileInfo employeesFile = new FileInfo(pathOfExcelFile);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(employeesFile))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                int colCount = worksheet.Dimension.End.Column; // get Column count
                int rowCount = worksheet.Dimension.End.Row; // get Row count
                List<Employee> employeesList = new List<Employee>(); // list to hold Employee objects as data is being read from the excel file
                for (int row = 2; row <= rowCount; row++) // starting at 2 so as to skip headers
                {
                    int employeeID = 0;
                    string firstName = "";
                    string lastName = "";
                    string department = "";
                    float salary = 0;
                    DateOnly hireDate = DateOnly.ParseExact("21 Oct 2015", "dd MMM yyyy", CultureInfo.InvariantCulture); ;
                    bool isActive = false;

                    for (int col = 1; col <= colCount; col++)
                    {
                        switch (col)
                        {
                            case 1:
                                employeeID = Convert.ToInt32(worksheet.Cells[row, col].Value.ToString().Trim());
                                break;
                            case 2:
                                firstName = (string) worksheet.Cells[row, col].Value;
                                break;
                            case 3:
                                lastName =  (string) worksheet.Cells[row, col].Value;
                                break;
                            case 4:
                                department = (string) worksheet.Cells[row, col].Value;
                                break;
                            case 5:
                                salary = (float)Convert.ToDouble(worksheet.Cells[row, col].Value);
                                break;
                            case 6:
                                hireDate = DateOnly.FromDateTime(DateTime.ParseExact(worksheet.Cells[row, col].Value.ToString().Trim(), "yyyy/MM/dd HH:mm:ss", CultureInfo.InvariantCulture));
                                break;
                            case 7:
                                isActive = (bool) worksheet.Cells[row, col].Value;
                                break;
                            default:
                                // do nothing - will never happen
                                break;
                        }
                    }
                    employeesList.Add(new Employee(employeeID, firstName, lastName, department, salary, hireDate, isActive));
                }
                Console.WriteLine("Adding employees data read from Excel file.");
                foreach (Employee employee in employeesList)
                {
                    context.Add(employee);
                    context.SaveChanges();
                }
                Console.WriteLine("Finished adding employees data read from Excel file.");
            }
        }
        public static void ReadFromDatabase(EFCoreEmployeeContext context)
        {
            Console.WriteLine("Reading data from the database");
            Console.WriteLine("");
            foreach(Employee employee in context.Employees)
            {
                Console.WriteLine(employee.EmployeeID);
                Console.WriteLine(employee.FirstName);
                Console.WriteLine(employee.LastName);
                Console.WriteLine(employee.FirstName);
                Console.WriteLine(employee.Department);
                Console.WriteLine(employee.Salary);
                Console.WriteLine(employee.HireDate);
                Console.WriteLine(employee.IsActive);
                Console.WriteLine("");
            }
            Console.WriteLine("");
            Console.WriteLine("Completed reading data drom the database!");
        }
    }
}
