using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.SpreadSheetsViewer
{
    public class StudentCalendar : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file
            var fileStream = this.GetResourceStream("xlsx\\Student calendar (Mon).xlsx");
            workbook.Open(fileStream);
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

        public override string TemplateName => "Student calendar(Mon).xlsx";

        public override bool CanDownloadZip => false;

        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\Student calendar (Mon).xlsx" };
            }
        }
    }
}
