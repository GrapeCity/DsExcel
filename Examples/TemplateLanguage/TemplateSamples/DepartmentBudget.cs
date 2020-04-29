using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Templates.TemplateSamples
{
    public class DepartmentBudget : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file from resource
            var templateFile = this.GetResourceStream("xlsx\\Template_DepartmentBudget.xlsx");
            workbook.Open(templateFile);

            #region Define custom classes
            //class Departments
            //{
            //    public List<Department> dpt;
            //}

            //class Department
            //{
            //    public string name;
            //    public string mgr;
            //    public double bud;
            //    public List<Employee> emp;
            //}

            //class Employee
            //{
            //    public string name;
            //    public double salary;
            //}
            #endregion

            #region Init Data
            var departments = new Departments
            {
                dpt = new List<Department>()
            };

            //Department 1
            var department1 = new Department
            {
                name = "Marketing",
                mgr = "Carl Sommerset",
                bud = 354586,
            };

            department1.emp = new List<Employee>
            {
                new Employee
                {
                    name = "JoeKline",
                    salary = 49402
                },
                new Employee
                {
                    name = "Lisa Crane",
                    salary = 81337
                },
                new Employee
                {
                    name = "John Ryes",
                    salary = 43503
                },
                new Employee
                {
                    name = "Elli Davidson",
                    salary = 67334
                },
                new Employee
                {
                    name = "Jack Reaze",
                    salary = 68314
                },
                new Employee
                {
                    name = "Ben Lam",
                    salary = 44696
                }
            };

            departments.dpt.Add(department1);

            //Department 2
            var department2 = new Department
            {
                name = "Sales",
                mgr = "Kelly Johnson",
                bud = 237721,
            };

            department2.emp = new List<Employee>
            {
                new Employee
                {
                    name = "Liam Elmerson",
                    salary = 61892
                },
                new Employee
                {
                    name = "Angela Sanderson",
                    salary = 38020
                },
                new Employee
                {
                    name = "Blake Schwarz",
                    salary = 55701
                },
                new Employee
                {
                    name = "Linda Barataz",
                    salary = 82108
                }
            };

            departments.dpt.Add(department2);

            //Department 3
            var department3 = new Department
            {
                name = "Engineering",
                mgr = "Gina Davis",
                bud = 624789,
            };
            department3.emp = new List<Employee>
            {
                new Employee
                {
                    name = "Christopher Dean",
                    salary = 58329
                },
                new Employee
                {
                    name = "Jack Linner",
                    salary = 63684
                },
                new Employee
                {
                    name = "Cathy Raines",
                    salary = 73147
                },
                new Employee
                {
                    name = "Scott Ashton",
                    salary = 77213
                },
                new Employee
                {
                    name = "Larry Wisell",
                    salary = 72796
                },
                new Employee
                {
                    name = "Bart Ingram",
                    salary = 50009
                },
                new Employee
                {
                    name = "Wesley Page",
                    salary = 82378
                },
                new Employee
                {
                    name = "Alan Keyes",
                    salary = 67105
                },
                new Employee
                {
                    name = "Wilson Musk",
                    salary = 80128
                }
            };

            departments.dpt.Add(department3);

            //Department 4
            var department4 = new Department
            {
                name = "Customer Service",
                mgr = "Kenneth Smith",
                bud = 127596
            };

            department4.emp = new List<Employee>
            {
                new Employee
                {
                    name = "Sherry Meeks",
                    salary = 38919
                },
                new Employee
                {
                    name = "Sharon Reeves",
                    salary = 40963
                },
                new Employee
                {
                    name = "Max Devillo",
                    salary = 47714
                }
            };

            departments.dpt.Add(department4);
            #endregion

            //Add data source
            workbook.AddDataSource("ds", departments);
            //Invoke to process the template
            workbook.ProcessTemplate();
        }

        public override string TemplateName
        {
            get
            {
                return "Template_DepartmentBudget.xlsx";
            }
        }

        public override bool ShowTemplate
        {
            get
            {
                return true;
            }
        }

        public override bool HasTemplate
        {
            get
            {
                return true;
            }
        }

        public override bool CanDownloadZip
        {
            get
            {
                return false;
            }
        }

        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\Template_DepartmentBudget.xlsx" };
            }
        }

        public override string[] Refs
        {
            get
            {
                return new string[] { "Departments", "Department", "Employee" };
            }
        }
    }

    class Departments
    {
        public List<Department> dpt;
    }

    class Department
    {
        public string name;
        public string mgr;
        public double bud;
        public List<Employee> emp;
    }

    class Employee
    {
        public string name;
        public double salary;
    }
}
