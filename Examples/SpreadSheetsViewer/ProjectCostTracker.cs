using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.SpreadSheetsViewer
{
    public class ProjectCostTracker : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file
            var fileStream = this.GetResourceStream("xlsx\\Project cost tracker.xlsx");
            workbook.Open(fileStream);
        }

        public override string TemplateName
        {
            get
            {
                return "Project cost tracker.xlsx";
            }
        }

        public override bool HasTemplate
        {
            get
            {
                return true;
            }
        }

        public override bool IsViewReadOnly
        {
            get
            {
                return false;
            }
        }

        public override bool ShowCode
        {
            get
            {
                return false;
            }
        }
        public override bool CanDownloadZip => false;
        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\Project cost tracker.xlsx" };
            }
        }
    }
}
