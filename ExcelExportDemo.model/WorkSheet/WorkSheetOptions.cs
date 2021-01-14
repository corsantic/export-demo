namespace ExcelExportDemo.model.WorkSheet
{
    public class WorkSheetOptions
    {
        public string WorkSheetTitle { get; set; }
        public int[] FreezeColumns { get; set; }
        public int[] FreezeRows { get; set; }
    }
}