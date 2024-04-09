#From https://stackoverflow.com/a/16537996   https://hackmd.io/@DailyOops/batch-convert-word-docx-to-pdf

$documents_path = 'D:\work2'
# From 這邊工作目錄要改掉
$excel_app = New-Object -ComObject excel.application
# 啟動EXCELapp
# 下面是管線操作，搜索xls檔案餵給後面的ecxcel開啟
Get-ChildItem -Path $documents_path -Filter *.xls? -Recurse | ForEach-Object {
    $document = $excel_app.workbooks.open($_.FullName)
    $pdf_filename = "$($_.DirectoryName)\$($_.BaseName).pdf"
    $document.ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $pdf_filename)
    $excel_app.Workbooks.close()
}
# 關閉excelapp離開
$excel_app.Quit()
