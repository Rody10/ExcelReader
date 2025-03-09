using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader
{
    internal class Employee
    {
       
        public Employee(int employeeID, string firstName, string lastName, string department, float salary, DateOnly hireDate, bool IsActive)
        {
            this.EmployeeID = employeeID;
            this.FirstName = firstName;
            this.LastName = lastName;
            this.Department = department;
            this.Salary = salary;
            this.HireDate = hireDate;
            this.IsActive = IsActive;
        }  // constructor

        [Key]
        public int EmployeeID { get; set; }
        public string FirstName { get; set; }

        public string LastName { get; set; }

        public string Department { get; set; }

        public float Salary { get; set; }

        public DateOnly HireDate { get; set; }

        public bool IsActive { get; set; }

        

    }
}
