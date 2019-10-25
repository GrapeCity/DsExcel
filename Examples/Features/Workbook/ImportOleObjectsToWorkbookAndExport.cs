namespace GrapeCity.Documents.Excel.Examples.Features.Workbook
{
    public class ImportOleObjectsToWorkbookAndExport : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            workbook.Open(GetResourceStream(@"xlsx\OleTemplates.xlsx"));
        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }

        public override bool CanDownload
        {
            get
            {
                return true;
            }
        }

        public override string[] UsedResources { get; } = 
            new string[] { @"xlsx\OleTemplates.xlsx" };
    }

}
