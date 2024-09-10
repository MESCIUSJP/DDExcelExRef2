using GrapeCity.Documents.Excel;

//Workbook.SetLicenseKey("");

// 集計用のワークブックを読み込み
Workbook workbook = new();
workbook.Open("sendai.xlsx");

// 外部参照を設定
workbook.Worksheets[0].Range["B4:B9"].Formula = "='[aoba.xlsx]Sheet1'!D4+'[izumi.xlsx]Sheet1'!D4+'[miyagino.xlsx]Sheet1'!D4+'[taihaku.xlsx]Sheet1'!D4";

// 外部ワークブックを読み込み（青葉区）
Workbook aoba = new();
aoba.Open("aoba.xlsx");

// 外部ワークブックを読み込み（泉区）
Workbook izumi = new();
izumi.Open("izumi.xlsx");

// 外部ワークブックを読み込み（宮城野区）
Workbook miyagino = new();
miyagino.Open("miyagino.xlsx");

// 外部ワークブックを読み込み（太白区）
Workbook taihaku = new();
taihaku.Open("taihaku.xlsx");

// 外部参照を更新
workbook.UpdateExcelLink("aoba.xlsx", aoba);
workbook.UpdateExcelLink("izumi.xlsx", izumi);
workbook.UpdateExcelLink("miyagino.xlsx", miyagino);
workbook.UpdateExcelLink("taihaku.xlsx", taihaku);
workbook.Calculate();

// EXCELファイルに保存
workbook.Save("result.xlsx");

