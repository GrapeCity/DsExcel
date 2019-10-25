using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Templates.DataSource
{
    public class Variable : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file Template_StudentInfo.xlsx from resource
            var templateFile = this.GetResourceStream("xlsx\\Template_StudentInfo.xlsx");
            workbook.Open(templateFile);

            #region Define custom classes
            //public class StudentInfo
            //{
            //    public string name;
            //    public string address;
            //    public List<Family> family;
            //}
            #endregion

            #region Init Data
            var studentInfos = new List<StudentInfo>
            {
                new StudentInfo
                {
                    name = "Jane",
                    address = "101, Halford Avenue, Fremont, CA"
                },
                new StudentInfo
                {
                    name = "Mark",
                    address = "2005 Klamath Ave APT, Santa Clara, CA"
                },
                new StudentInfo
                {
                    name = "Carol",
                    address = "1063 E EI Camino Real, Sunnyvale, CA 94087, USA"
                },
                new StudentInfo
                {
                    name = "Liano",
                    address = "1977 St Lawrence Dr, Santa Clara, CA 95051, USA"
                },
                new StudentInfo
                {
                    name = "Hellen",
                    address = "3661 Peacock Ct, Santa Clara, CA 95051, USA"
                }
            };

            var className = "Class 3"; 
            #endregion

            //Add data source
            workbook.AddDataSource("className", className);
            workbook.AddDataSource("s", studentInfos);

            //Invoke to process the template
            workbook.ProcessTemplate();
        }

        public override bool IsNew
        {
            get
            {
                return true;
            }
        }

        public override string TemplateName
        {
            get
            {
                return "Template_StudentInfo.xlsx";
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
                return new string[] { "xlsx\\Template_StudentInfo.xlsx"};
            }
        }

        public override string[] Refs
        {
            get
            {
                return new string[] { "StudentInfo", "Family", "Guardian" };
            }
        }
    }
}
