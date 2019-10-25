using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Templates.DataSource
{
    public class JsonFile : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file Template_FamilyInfo.xlsx from resource
            var templateFile = this.GetResourceStream("xlsx\\Template_FamilyInfo.xlsx");
            workbook.Open(templateFile);

            #region Define custom classes
            //public class StudentInfos
            //{
            //    public List<StudentInfo> student;
            //}

            //public class StudentInfo
            //{
            //    public string name;
            //    public string address;
            //    public List<Family> family;
            //}

            //public class Family
            //{
            //    public Guardian father;
            //    public Guardian mother;
            //}

            //public class Guardian
            //{
            //    public string name;
            //    public string occupation;
            //}
            #endregion

            //Get data from json file
            string jsonText = string.Empty;
            using (Stream stream = this.GetResourceStream("Template_FamilyInfo.json"))
            using (StreamReader reader = new StreamReader(stream))
            {
                jsonText = reader.ReadToEnd();
            }

            var datasource = JsonConvert.DeserializeObject<StudentInfos>(jsonText);

            //Add data source
            workbook.AddDataSource("ds", datasource);
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
                return "Template_FamilyInfo.xlsx";
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
                return new string[] { "xlsx\\Template_FamilyInfo.xlsx", "Template_FamilyInfo.json" };
            }
        }

        public override string[] Refs
        {
            get
            {
                return new string[] { "StudentInfos", "StudentInfo", "Family", "Guardian" };
            }
        }
    }

    public class StudentInfos
    {
        public List<StudentInfo> student;
    }


    public class StudentInfo
    {
        public string name;
        public string address;
        public List<Family> family;
    }

    public class Family
    {
        public Guardian father;
        public Guardian mother;
    }

    public class Guardian
    {
        public string name;
        public string occupation;
    }
}
