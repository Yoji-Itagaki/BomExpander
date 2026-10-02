Imports System.ComponentModel
Imports System.IO
Imports BomTool.Core

''' <summary>
''' メイン画面です。
''' CSVの選択、出力先の選択、結果の表示だけを行い、
''' 読み込み・展開・Excel出力は BomTool.Core のクラス（CsvBomReader / BomTreeExpander / ExcelExporter）に任せます。
''' </summary>
Public Class MainForm

    ''' <summary>
    ''' 1回の処理結果をまとめたものです（別スレッドから画面へ渡すために使います）。
    ''' </summary>
    Private Class ProcessResult
        Public Property ProductCount As Integer
        Public Property PartLineCount As Integer
        Public Property ErrorCount As Integer
        Public Property EncodingName As String = ""
    End Class

    Public Sub New()
        InitializeComponent()
        ClearResult()
    End Sub

    ''' <summary>
    ''' 「CSVを選択」ボタン：ファイルを選ぶダイアログを開きます。
    ''' </summary>
    Private Sub btnSelectCsv_Click(sender As Object, e As EventArgs) Handles btnSelectCsv.Click
        Using dialog As New OpenFileDialog With {
            .Title = "部品表CSVを選択",
            .Filter = "CSVファイル (*.csv)|*.csv|すべてのファイル (*.*)|*.*",
            .CheckFileExists = True
        }
            If dialog.ShowDialog(Me) <> DialogResult.OK Then Return
            txtCsvPath.Text = dialog.FileName
            btnExport.Enabled = True
            ClearResult()
        End Using
    End Sub

    ''' <summary>
    ''' 「展開してExcelに出力」ボタン：保存先を選び、展開とExcel出力を行います。
    ''' 処理中に例外が起きても、メッセージを出して画面はそのまま使えるようにします。
    ''' </summary>
    Private Async Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Dim csvPath = txtCsvPath.Text
        If Not File.Exists(csvPath) Then
            MessageBox.Show(Me, "CSVファイルが見つかりません。もう一度選択してください。", "確認",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim outputPath As String
        Using dialog As New SaveFileDialog With {
            .Title = "Excelファイルの保存先",
            .Filter = "Excelブック (*.xlsx)|*.xlsx",
            .DefaultExt = "xlsx",
            .AddExtension = True,
            .OverwritePrompt = True,
            .InitialDirectory = Path.GetDirectoryName(csvPath),
            .FileName = Path.GetFileNameWithoutExtension(csvPath) & "_BOM展開.xlsx"
        }
            If dialog.ShowDialog(Me) <> DialogResult.OK Then Return
            outputPath = dialog.FileName
        End Using

        SetBusy(True)
        Try
            ' 時間のかかる処理は別スレッドで行い、画面が固まらないようにする
            Dim result = Await Task.Run(Function() RunConversion(csvPath, outputPath))
            ShowResult(result, outputPath)

            Dim message = $"Excelファイルを出力しました。{vbCrLf}{outputPath}"
            If result.ErrorCount > 0 Then
                message &= $"{vbCrLf}{vbCrLf}エラーが{result.ErrorCount}件あります。「エラー」シートを確認してください。"
            End If
            message &= $"{vbCrLf}{vbCrLf}出力したファイルを開きますか？"
            If MessageBox.Show(Me, message, "完了", MessageBoxButtons.YesNo,
                               If(result.ErrorCount > 0, MessageBoxIcon.Warning, MessageBoxIcon.Information)) = DialogResult.Yes Then
                OpenFile(outputPath)
            End If

        Catch ex As IOException
            ShowError("ファイルの読み込みまたは保存に失敗しました。" & vbCrLf &
                      "出力先のファイルをExcelで開いている場合は、閉じてからもう一度実行してください。", ex)
        Catch ex As UnauthorizedAccessException
            ShowError("ファイルにアクセスする権限がありません。別の保存先を選んでください。", ex)
        Catch ex As Exception
            ShowError("処理中にエラーが発生しました。", ex)
        Finally
            SetBusy(False)
        End Try
    End Sub

    ''' <summary>
    ''' CSV読み込み → BOM展開 → Excel出力 を順番に行います（画面の部品には触らないこと）。
    ''' </summary>
    Private Shared Function RunConversion(csvPath As String, outputPath As String) As ProcessResult
        Dim readResult = New CsvBomReader().ReadFile(csvPath)
        Dim expansion = New BomTreeExpander().Expand(readResult.Rows)
        Dim allErrors = readResult.Errors.Concat(expansion.Errors).ToList()

        Call New ExcelExporter().Export(outputPath, expansion, allErrors)

        Return New ProcessResult With {
            .ProductCount = expansion.Products.Count,
            .PartLineCount = expansion.PartLineCount,
            .ErrorCount = allErrors.Count,
            .EncodingName = readResult.EncodingName
        }
    End Function

    ''' <summary>処理結果を画面に表示します。</summary>
    Private Sub ShowResult(result As ProcessResult, outputPath As String)
        lblProductCountValue.Text = $"{result.ProductCount:#,0} 件"
        lblPartLineCountValue.Text = $"{result.PartLineCount:#,0} 行"
        lblErrorCountValue.Text = $"{result.ErrorCount:#,0} 件"
        lblErrorCountValue.ForeColor = If(result.ErrorCount > 0, Color.Firebrick, SystemColors.ControlText)
        lblEncodingValue.Text = result.EncodingName
        lblOutputValue.Text = outputPath
        toolTip.SetToolTip(lblOutputValue, outputPath)
    End Sub

    ''' <summary>処理結果の表示を消します。</summary>
    Private Sub ClearResult()
        For Each valueLabel In {lblProductCountValue, lblPartLineCountValue, lblErrorCountValue, lblEncodingValue, lblOutputValue}
            valueLabel.Text = "－"
            valueLabel.ForeColor = SystemColors.ControlText
        Next
        toolTip.SetToolTip(lblOutputValue, Nothing)
    End Sub

    ''' <summary>処理中はボタンを押せないようにし、マウスカーソルを砂時計にします。</summary>
    Private Sub SetBusy(busy As Boolean)
        btnSelectCsv.Enabled = Not busy
        btnExport.Enabled = Not busy AndAlso txtCsvPath.Text.Length > 0
        btnExport.Text = If(busy, "処理中...", "展開してExcelに出力")
        UseWaitCursor = busy
    End Sub

    ''' <summary>エラーメッセージを表示します。</summary>
    Private Sub ShowError(message As String, ex As Exception)
        MessageBox.Show(Me, $"{message}{vbCrLf}{vbCrLf}詳細：{ex.Message}", "エラー",
                        MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    ''' <summary>出力したファイルを、パソコンに関連付けられたアプリ（Excelなど）で開きます。</summary>
    Private Sub OpenFile(filePath As String)
        Try
            Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
        Catch ex As Win32Exception
            ShowError(".xlsx ファイルを開けるアプリが見つかりません。ファイルは保存されています。", ex)
        End Try
    End Sub

End Class
