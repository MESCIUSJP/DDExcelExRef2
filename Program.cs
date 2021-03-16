namespace GrapeCity.Documents.Excel.Examples.Features.Formulas.CrossWorkbookFormula2
{
    class Program
    {
        static void Main(string[] args)
        {
            Workbook.SetLicenseKey("");

            // 集計用のワークブックを読み込み
            var workbook = new Workbook();
            workbook.Open("sendai.xlsx");

            // 外部参照を設定
            workbook.Worksheets[0].Range["B4:B9"].Formula = "='[aoba.xlsx]Sheet1'!D4+'[izumi.xlsx]Sheet1'!D4+'[miyagino.xlsx]Sheet1'!D4+'[taihaku.xlsx]Sheet1'!D4";


            // 外部ワークブックを読み込み（青葉区）
            var aoba = new Workbook();
            aoba.Open("aoba.xlsx");

            // 外部ワークブックを読み込み（泉区）
            var izumi = new Workbook();
            izumi.Open("izumi.xlsx");

            // 外部ワークブックを読み込み（宮城野区）
            var miyagino = new Workbook();
            miyagino.Open("miyagino.xlsx");

            // 外部ワークブックを読み込み（太白区）
            var taihaku = new Workbook();
            taihaku.Open("taihaku.xlsx");

            // 外部参照を更新
            workbook.UpdateExcelLink("aoba.xlsx", aoba);
            workbook.UpdateExcelLink("izumi.xlsx", izumi);
            workbook.UpdateExcelLink("miyagino.xlsx", miyagino);
            workbook.UpdateExcelLink("taihaku.xlsx", taihaku);
            workbook.Calculate();

            // EXCELファイル（.xlsx）に保存
            workbook.Save("crossworkbookformula2.xlsx");

        }
    }
}