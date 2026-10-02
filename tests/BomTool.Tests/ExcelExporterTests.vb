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
    Public Sub 四つのシートが作られ見出し行が固定される()
        Using wb = ExportAndOpen(Csv("P,A,部品A,2,個,購入"))
            Assert.Equal({"BOM展開", "部品集計", "逆展開", "エラー"}, wb.Worksheets.Select(Function(s) s.Name))
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
    Public Sub 逆展開シートに部品から製品までの経路を出力する()
        Using wb = ExportAndOpen(Csv(
                "P,A,組品A,2,個,内製",
                "A,B,部品B,3,本,購入"))
            Dim sheet = wb.Worksheet("逆展開")
            Assert.Equal("部品品番", sheet.Cell(1, 1).GetString())
            Assert.Equal("製品", sheet.Cell(1, 10).GetString())

            ' 2〜3行目：部品A（A → P）、4〜6行目：部品B（B → A → P）
            Dim codes = Enumerable.Range(2, 5).Select(Function(r) (sheet.Cell(r, 1).GetString(), sheet.Cell(r, 4).GetString())).ToArray()
            Assert.Equal({("A", "A"), ("A", "P"), ("B", "B"), ("B", "A"), ("B", "P")}, codes)

            ' 部品自身の行：員数は空、累計員数1、色付き。部品が切り替わる行の上は太線
            Assert.True(sheet.Cell(4, 6).IsEmpty())
            Assert.Equal(1D, sheet.Cell(4, 7).GetValue(Of Decimal)())
            Assert.True(sheet.Cell(4, 1).Style.Font.Bold)
            Assert.Equal(XLBorderStyleValues.Medium, sheet.Cell(4, 1).Style.Border.TopBorder)

            ' 組品Aの行：字下げ2文字、員数3、累計員数3
            Assert.Equal(2, sheet.Cell(5, 4).Style.Alignment.Indent)
            Assert.Equal("組品A", sheet.Cell(5, 5).GetString())
            Assert.Equal(3D, sheet.Cell(5, 6).GetValue(Of Decimal)())
            Assert.Equal(3D, sheet.Cell(5, 7).GetValue(Of Decimal)())
            Assert.Equal("内製", sheet.Cell(5, 9).GetString())
            Assert.True(sheet.Cell(5, 10).IsEmpty())

            ' 製品の行：累計員数 3×2=6（製品1台あたりの部品Bの数）、単位は部品Bの単位、製品列に品番
            Assert.Equal(6D, sheet.Cell(6, 7).GetValue(Of Decimal)())
            Assert.Equal("本", sheet.Cell(6, 8).GetString())
            Assert.Equal("P", sheet.Cell(6, 10).GetString())
            Assert.True(sheet.Cell(6, 4).Style.Font.Bold)
        End Using
    End Sub

    <Fact>
    Public Sub 部品集計シートは製品が切り替わる行の上に太線を引く()
        Using wb = ExportAndOpen(Csv(
                "P1,A,部品A,1,個,購入",
                "P1,B,部品B,1,個,購入",
                "P2,A,部品A,1,個,購入"))
            Dim sheet = wb.Worksheet("部品集計")
            Assert.Equal("P2", sheet.Cell(4, 1).GetString())
            Assert.Equal(XLBorderStyleValues.Medium, sheet.Cell(4, 1).Style.Border.TopBorder)
            Assert.Equal(XLBorderStyleValues.Thin, sheet.Cell(3, 1).Style.Border.TopBorder)
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
