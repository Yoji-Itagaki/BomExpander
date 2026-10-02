Imports System.IO
Imports BomTool.Core
Imports ClosedXML.Excel
Imports Xunit

''' <summary>
''' ExcelExporter のテストです。実際に .xlsx を作り、読み直して中身を確認します。
''' </summary>
Public Class ExcelExporterTests
    Implements IDisposable

    Private ReadOnly _outputPath As String = Path.Combine(Path.GetTempPath(), $"BomToolTest_{Guid.NewGuid():N}.xlsx")

    Public Sub Dispose() Implements IDisposable.Dispose
        If File.Exists(_outputPath) Then File.Delete(_outputPath)
    End Sub

    Private Function ExportAndOpen(csvText As String) As XLWorkbook
        Dim r = ReadAndExpand(csvText)
        Call New ExcelExporter().Export(_outputPath, r.Expansion, r.Read.Errors.Concat(r.Expansion.Errors))
        Return New XLWorkbook(_outputPath)
    End Function

    <Fact>
    Public Sub 三つのシートが作られ見出し行が固定される()
        Using wb = ExportAndOpen(Csv("P,A,部品A,2,個,購入"))
            Assert.Equal({"BOM展開", "部品集計", "エラー"}, wb.Worksheets.Select(Function(s) s.Name))
            For Each sheet In wb.Worksheets
                Assert.Equal(1, sheet.SheetView.SplitRow)
                Assert.True(sheet.Cell(1, 1).Style.Font.Bold)
                ' 内側の罫線は細線、表の外枠は中太線
                Dim lastRow = sheet.LastRowUsed().RowNumber()
                Assert.Equal(XLBorderStyleValues.Thin, sheet.Cell(2, 1).Style.Border.TopBorder)
                Assert.Equal(XLBorderStyleValues.Medium, sheet.Cell(2, 1).Style.Border.LeftBorder)
                Assert.Equal(XLBorderStyleValues.Medium, sheet.Cell(lastRow, 1).Style.Border.BottomBorder)
            Next
        End Using
    End Sub

    <Fact>
    Public Sub BOM展開シートにレベル_字下げ_累計員数を出力する()
        Using wb = ExportAndOpen(Csv(
                "P,A,組品A,2,個,内製",
                "A,B,部品B,3,本,購入"))
            Dim sheet = wb.Worksheet("BOM展開")
            Assert.Equal("レベル", sheet.Cell(1, 1).GetString())
            Assert.Equal("品番", sheet.Cell(1, 2).GetString())

            ' 2行目：製品（レベル0）
            Assert.Equal(0, sheet.Cell(2, 1).GetValue(Of Integer)())
            Assert.Equal("P", sheet.Cell(2, 2).GetString())
            ' 4行目：部品B（レベル2、累計員数 2×3=6）
            Assert.Equal(2, sheet.Cell(4, 1).GetValue(Of Integer)())
            Assert.Equal("B", sheet.Cell(4, 2).GetString())
            Assert.Equal(4, sheet.Cell(4, 2).Style.Alignment.Indent)
            Assert.Equal("部品B", sheet.Cell(4, 3).GetString())
            Assert.Equal(3D, sheet.Cell(4, 4).GetValue(Of Decimal)())
            Assert.Equal(6D, sheet.Cell(4, 5).GetValue(Of Decimal)())
            Assert.Equal("本", sheet.Cell(4, 6).GetString())
            Assert.Equal("購入", sheet.Cell(4, 7).GetString())
        End Using
    End Sub

    <Fact>
    Public Sub エラーシートにエラーを行番号順に出力する()
        Using wb = ExportAndOpen(Csv(
                "P,A,部品A,0,個,購入",
                "P,B,部品B,1,個,購入",
                "P,B,部品B,1,個,購入"))
            Dim sheet = wb.Worksheet("エラー")
            ' CSV 2行目：員数が0、CSV 4行目：親子の重複
            Assert.Equal(2, sheet.Cell(2, 1).GetValue(Of Integer)())
            Assert.Equal("員数が不正", sheet.Cell(2, 2).GetString())
            Assert.Equal(4, sheet.Cell(3, 1).GetValue(Of Integer)())
            Assert.Equal("親子の重複", sheet.Cell(3, 2).GetString())
        End Using
    End Sub

    <Fact>
    Public Sub エラーがない場合はその旨を出力する()
        Using wb = ExportAndOpen(Csv("P,A,部品A,1,個,購入"))
            Assert.Equal("エラーはありません。", wb.Worksheet("エラー").Cell(2, 3).GetString())
        End Using
    End Sub

    <Theory>
    <InlineData("ABC", 3)>
    <InlineData("部品", 4)>
    <InlineData("ﾈｼﾞ", 3)>
    Public Sub 全角文字は2文字分_半角文字は1文字分の幅として数える(text As String, expected As Integer)
        Assert.Equal(expected, ExcelExporter.DisplayWidth(text))
    End Sub

End Class
