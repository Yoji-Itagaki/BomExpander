Imports ClosedXML.Excel

''' <summary>
''' 展開結果をExcelファイル（.xlsx）に出力するクラスです。
''' ClosedXML を使うため、Excelがインストールされていないパソコンでも動きます。
'''
''' 出力するシート：
''' ・「BOM展開」：製品ごとの多段階部品表
''' ・「部品集計」：製品ごと・品番ごとの合計員数
''' ・「逆展開」　：部品ごとに、使われている組品から製品までを逆にたどった表
''' ・「エラー」　：読み込み・展開で見つかった問題の一覧
''' </summary>
Public Class ExcelExporter

    Public Const BomSheetName As String = "BOM展開"
    Public Const SummarySheetName As String = "部品集計"
    Public Const WhereUsedSheetName As String = "逆展開"
    Public Const ErrorSheetName As String = "エラー"

    ''' <summary>ブック全体の文字フォント</summary>
    Public Property FontName As String = "游ゴシック"

    ''' <summary>文字サイズ</summary>
    Public Property FontSize As Double = 10

    ' 見た目の設定（色を変えたいときはここを修正します）
    Private Shared ReadOnly HeaderBackColor As XLColor = XLColor.FromHtml("#1F4E78")
    Private Shared ReadOnly HeaderFontColor As XLColor = XLColor.White
    Private Shared ReadOnly ProductBackColor As XLColor = XLColor.FromHtml("#DDEBF7")
    Private Shared ReadOnly PartBackColor As XLColor = XLColor.FromHtml("#FFF2CC")
    Private Shared ReadOnly ErrorBackColor As XLColor = XLColor.FromHtml("#FCE4D6")

    ''' <summary>
    ''' Excelファイルを出力します。同じ名前のファイルがあれば上書きします。
    ''' </summary>
    ''' <param name="path">出力先のパス（.xlsx）</param>
    ''' <param name="expansion">展開結果</param>
    ''' <param name="errors">エラーシートに出すエラー（読み込み時と展開時の両方）</param>
    Public Sub Export(path As String, expansion As BomExpansionResult, errors As IEnumerable(Of BomError))
        Using workbook As New XLWorkbook()
            workbook.Style.Font.FontName = FontName
            workbook.Style.Font.FontSize = FontSize

            WriteBomSheet(workbook.Worksheets.Add(BomSheetName), expansion.Lines)
            WriteSummarySheet(workbook.Worksheets.Add(SummarySheetName), expansion.Summary)
            WriteWhereUsedSheet(workbook.Worksheets.Add(WhereUsedSheetName), expansion.WhereUsed)
            WriteErrorSheet(workbook.Worksheets.Add(ErrorSheetName), errors.ToList())

            workbook.Worksheet(1).SetTabActive()
            workbook.SaveAs(path)
        End Using
    End Sub

    ''' <summary>
    ''' 「BOM展開」シートを作ります。品番はレベルに応じて字下げします。製品の行（レベル0）は色付き・太字です。
    ''' </summary>
    Private Sub WriteBomSheet(sheet As IXLWorksheet, lines As IList(Of ExpandedLine))
        Dim headers = {"レベル", "品番", "品名", "員数", "累計員数" & vbLf & "（製品1台あたり）", "単位", "手配区分"}
        WriteHeader(sheet, headers)

        Dim r = 2
        For Each line In lines
            sheet.Cell(r, 1).Value = line.Level
            sheet.Cell(r, 2).Value = line.ItemCode
            sheet.Cell(r, 2).Style.Alignment.Indent = line.Level * 2   ' レベル1つにつき2文字分の字下げ
            sheet.Cell(r, 3).Value = line.ItemName
            sheet.Cell(r, 4).Value = line.Quantity
            sheet.Cell(r, 5).Value = line.CumulativeQuantity
            sheet.Cell(r, 6).Value = line.Unit
            sheet.Cell(r, 7).Value = line.Procurement

            If line.Level = 0 Then
                Dim productRow = sheet.Range(r, 1, r, headers.Length)
                productRow.Style.Fill.BackgroundColor = ProductBackColor
                productRow.Style.Font.Bold = True
            End If
            r += 1
        Next

        sheet.Column(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        sheet.Column(6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        sheet.Column(7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        FinishSheet(sheet, headers.Length, r - 1)
    End Sub

    ''' <summary>
    ''' 「部品集計」シートを作ります。製品が切り替わる行の上に太線を引きます。
    ''' </summary>
    Private Sub WriteSummarySheet(sheet As IXLWorksheet, summary As IList(Of SummaryLine))
        Dim headers = {"製品品番", "手配区分", "品番", "品名", "合計員数" & vbLf & "（製品1台あたり）", "単位", "使用箇所数"}
        WriteHeader(sheet, headers)

        Dim r = 2
        Dim previousProduct As String = Nothing
        Dim separatorRows As New List(Of Integer)
        For Each item In summary
            sheet.Cell(r, 1).Value = item.ProductCode
            sheet.Cell(r, 2).Value = item.Procurement
            sheet.Cell(r, 3).Value = item.ItemCode
            sheet.Cell(r, 4).Value = item.ItemName
            sheet.Cell(r, 5).Value = item.TotalQuantity
            sheet.Cell(r, 6).Value = item.Unit
            sheet.Cell(r, 7).Value = item.UsageCount

            If previousProduct IsNot Nothing AndAlso previousProduct <> item.ProductCode Then separatorRows.Add(r)
            previousProduct = item.ProductCode
            r += 1
        Next

        sheet.Column(2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        sheet.Column(6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        FinishSheet(sheet, headers.Length, r - 1)
        DrawSeparators(sheet, separatorRows, headers.Length)
    End Sub

    ''' <summary>
    ''' 「逆展開」シートを作ります。品番はレベルに応じて字下げします。
    ''' 部品自身の行（レベル0）と製品の行は色付き・太字にし、部品が切り替わる行の上に太線を引きます。
    ''' 「部品品番」列で絞り込むと、1つの部品の使用先だけを見られます。
    ''' </summary>
    Private Sub WriteWhereUsedSheet(sheet As IXLWorksheet, lines As IList(Of WhereUsedLine))
        Dim headers = {"部品品番", "部品品名", "レベル", "品番", "品名", "員数",
                       "累計員数" & vbLf & "（この品番1個あたり）", "単位", "手配区分", "製品"}
        WriteHeader(sheet, headers)

        Dim r = 2
        Dim separatorRows As New List(Of Integer)
        For Each line In lines
            sheet.Cell(r, 1).Value = line.PartCode
            sheet.Cell(r, 2).Value = line.PartName
            sheet.Cell(r, 3).Value = line.Level
            sheet.Cell(r, 4).Value = line.ItemCode
            sheet.Cell(r, 4).Style.Alignment.Indent = line.Level * 2   ' レベル1つにつき2文字分の字下げ
            sheet.Cell(r, 5).Value = line.ItemName
            If line.Quantity.HasValue Then sheet.Cell(r, 6).Value = line.Quantity.Value
            sheet.Cell(r, 7).Value = line.CumulativeQuantity
            sheet.Cell(r, 8).Value = line.Unit
            sheet.Cell(r, 9).Value = line.Procurement
            If line.IsProduct Then sheet.Cell(r, 10).Value = line.ItemCode

            Dim rowRange = sheet.Range(r, 1, r, headers.Length)
            If line.Level = 0 Then
                rowRange.Style.Fill.BackgroundColor = PartBackColor
                rowRange.Style.Font.Bold = True
                If r > 2 Then separatorRows.Add(r)
            ElseIf line.IsProduct Then
                rowRange.Style.Fill.BackgroundColor = ProductBackColor
                rowRange.Style.Font.Bold = True
            End If
            r += 1
        Next

        sheet.Column(3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        sheet.Column(8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        sheet.Column(9).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        FinishSheet(sheet, headers.Length, r - 1)
        DrawSeparators(sheet, separatorRows, headers.Length)
    End Sub

    ''' <summary>
    ''' 「エラー」シートを作ります。エラーがない場合は「エラーはありません」と1行出します。
    ''' </summary>
    Private Sub WriteErrorSheet(sheet As IXLWorksheet, errors As IList(Of BomError))
        Dim headers = {"行番号", "エラー種別", "内容", "元データ"}
        WriteHeader(sheet, headers)

        Dim r = 2
        If errors.Count = 0 Then
            sheet.Cell(r, 3).Value = "エラーはありません。"
            r += 1
        Else
            ' 行番号順（行番号のないエラーは最後）に並べます
            For Each item In errors.OrderBy(Function(e) If(e.LineNumber = 0, Integer.MaxValue, e.LineNumber))
                If item.LineNumber > 0 Then sheet.Cell(r, 1).Value = item.LineNumber
                sheet.Cell(r, 2).Value = item.KindText
                sheet.Cell(r, 3).Value = item.Message
                sheet.Cell(r, 4).Value = item.RawText
                sheet.Cell(r, 2).Style.Fill.BackgroundColor = ErrorBackColor
                r += 1
            Next
        End If

        sheet.Column(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        FinishSheet(sheet, headers.Length, r - 1, maxColumnWidth:=80)
        ' 内容の列は長くなるため、折り返して表示します
        sheet.Range(2, 3, Math.Max(2, r - 1), 3).Style.Alignment.WrapText = True
    End Sub

    ''' <summary>
    ''' 1行目に見出しを書きます。
    ''' </summary>
    Private Sub WriteHeader(sheet As IXLWorksheet, headers As String())
        For c = 0 To headers.Length - 1
            sheet.Cell(1, c + 1).Value = headers(c)
        Next
        Dim headerRange = sheet.Range(1, 1, 1, headers.Length)
        headerRange.Style.Font.Bold = True
        headerRange.Style.Font.FontColor = HeaderFontColor
        headerRange.Style.Fill.BackgroundColor = HeaderBackColor
        headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center
        headerRange.Style.Alignment.WrapText = True
        sheet.Row(1).Height = 30
    End Sub

    ''' <summary>
    ''' シートの仕上げ：罫線、見出し行の固定、オートフィルター、列幅の自動調整、印刷設定。
    ''' </summary>
    ''' <param name="columnCount">列数</param>
    ''' <param name="lastRow">データの最終行（見出しのみの場合は1）</param>
    ''' <param name="maxColumnWidth">列幅の上限（文字数）</param>
    Private Sub FinishSheet(sheet As IXLWorksheet, columnCount As Integer, lastRow As Integer,
                            Optional maxColumnWidth As Double = 50)
        Dim table = sheet.Range(1, 1, Math.Max(1, lastRow), columnCount)

        ' 罫線（内側は細線、外枠は少し太く）
        table.Style.Border.InsideBorder = XLBorderStyleValues.Thin
        table.Style.Border.OutsideBorder = XLBorderStyleValues.Medium
        table.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center

        ' 見出し行を固定し、絞り込みできるようにする
        sheet.SheetView.FreezeRows(1)
        If lastRow > 1 Then table.SetAutoFilter()

        AdjustColumnWidths(sheet, columnCount, lastRow, maxColumnWidth)

        ' 印刷設定：A4、横幅を1ページに収める、見出し行を毎ページ印刷、ページ番号
        With sheet.PageSetup
            .PaperSize = XLPaperSize.A4Paper
            .PageOrientation = If(columnCount >= 7, XLPageOrientation.Landscape, XLPageOrientation.Portrait)
            .FitToPages(1, 0)
            .SetRowsToRepeatAtTop(1, 1)
            .CenterHorizontally = True
            .Margins.Top = 0.6
            .Margins.Bottom = 0.6
            .Margins.Left = 0.4
            .Margins.Right = 0.4
            .Header.Left.AddText(sheet.Name)
            .Footer.Center.AddText(XLHFPredefinedText.PageNumber)
            .Footer.Center.AddText(" / ")
            .Footer.Center.AddText(XLHFPredefinedText.NumberOfPages)
        End With
        sheet.PageSetup.PrintAreas.Add(1, 1, Math.Max(1, lastRow), columnCount)
    End Sub

    ''' <summary>
    ''' 指定した行の上に区切りの太線を引きます。
    ''' FinishSheet の罫線（内側は細線）で上書きされないよう、FinishSheet の後に呼び出します。
    ''' </summary>
    ''' <param name="rows">太線を引く行（その行の上に引きます）</param>
    ''' <param name="columnCount">列数</param>
    Private Shared Sub DrawSeparators(sheet As IXLWorksheet, rows As IEnumerable(Of Integer), columnCount As Integer)
        For Each r In rows
            sheet.Range(r, 1, r, columnCount).Style.Border.TopBorder = XLBorderStyleValues.Medium
            sheet.Range(r - 1, 1, r - 1, columnCount).Style.Border.BottomBorder = XLBorderStyleValues.Medium
        Next
    End Sub

    ''' <summary>
    ''' 列幅を内容に合わせて調整します。
    ''' ClosedXML の標準の自動調整は日本語（全角文字）の幅を小さく見積もることがあるため、
    ''' 全角文字を2文字分、半角文字を1文字分として数え、字下げ分も加えて幅を決めます。
    ''' </summary>
    Private Shared Sub AdjustColumnWidths(sheet As IXLWorksheet, columnCount As Integer, lastRow As Integer, maxWidth As Double)
        For c = 1 To columnCount
            Dim widest = 0.0
            For r = 1 To Math.Max(1, lastRow)
                Dim cell = sheet.Cell(r, c)
                ' 見出しなど改行を含む文字は、一番長い行で測ります
                Dim textWidth = cell.GetFormattedString().
                    Split(ControlChars.Lf).
                    Select(Function(part) DisplayWidth(part)).
                    DefaultIfEmpty(0).
                    Max()
                Dim width = textWidth + cell.Style.Alignment.Indent * 1.5
                If width > widest Then widest = width
            Next
            ' 余白とオートフィルターのボタンの分を足します
            sheet.Column(c).Width = Math.Min(maxWidth, Math.Max(6, widest * 1.1 + 4))
        Next
    End Sub

    ''' <summary>
    ''' 文字列の表示幅を返します（全角文字＝2、半角文字＝1）。
    ''' </summary>
    Friend Shared Function DisplayWidth(text As String) As Integer
        Dim width = 0
        For Each ch In text
            Dim code = Convert.ToInt32(ch)
            Dim isHalfWidthKana = code >= &HFF61 AndAlso code <= &HFF9F
            If (code >= &H2E80 OrElse (code >= &H3000 AndAlso code <= &H30FF)) AndAlso Not isHalfWidthKana Then
                width += 2
            Else
                width += 1
            End If
        Next
        Return width
    End Function

End Class
