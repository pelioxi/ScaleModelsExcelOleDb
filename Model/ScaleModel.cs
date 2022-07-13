using LinqToExcel.Attributes;

namespace ScaleModelsExcel.Model
{
    public class ScaleModel
    {
        [ExcelColumn("Collection")]
        public string Collection { get; }

        [ExcelColumn("Brand")]
        public string Brand { get; }

        [ExcelColumn("Model")]
        public string Model { get; }

        [ExcelColumn("CarYear")]
        public string CarYear { get; }

        [ExcelColumn("Maker")]
        public string Maker { get; }

        [ExcelColumn("Scale18")]
        public string Scale18 { get; }

        [ExcelColumn("Status")]
        public string Status { get; }

        [ExcelColumn("Scale")]
        public string Scale { get; }

        [ExcelColumn("PartNumber")]
        public string PartNumber { get; }

        [ExcelColumn("CarNumber")]
        public string CarNumber { get; }

        [ExcelColumn("ColourSponsor")]
        public string ColourSponsor { get; }

        [ExcelColumn("Driver")]
        public string Driver { get; }

        [ExcelColumn("Details")]
        public string Details { get; }

        [ExcelColumn("ModelDate")]
        public string ModelDate { get; }

        [ExcelColumn("Serial")]
        public string Serial { get; }

        [ExcelColumn("Ledition")]
        public string Ledition { get; }

        [ExcelColumn("Comments")]
        public string Comments { get; }
    }
}