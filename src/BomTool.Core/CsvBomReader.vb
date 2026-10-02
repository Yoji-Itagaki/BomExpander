Imports System.Globalization
Imports System.IO
Imports System.Text

''' <summary>
''' 部品表のCSVファイルを読み込み、1行ずつチェックして BomRow の一覧にするクラスです。
'''
''' ・文字コード：UTF-8（BOM付き・なし）と Shift_JIS を自動で判定します。
''' ・1行目はヘッダーとして読み飛ばします。
''' ・列の順番：親品番, 子品番, 品名, 員数, 単位, 手配区分
''' ・問題のある行はエラーとして記録し、その行は読み飛ばします。
''' </summary>
Public Class CsvBomReader

    ''' <summary>必要な列の数</summary>
    Public Const RequiredColumnCount As Integer = 6

    ''' <summary>手配区分として認める値</summary>
    Public Shared ReadOnly ProcurementTypes As IReadOnlyList(Of String) = {"購入", "内製"}

    ''' <summary>各列の名前（エラーメッセージに使います。並びはCSVの列順と同じ）</summary>
    Private Shared ReadOnly ColumnNames As String() = {"親品番", "子品番", "品名", "員数", "単位", "手配区分"}

    ''' <summary>
    ''' Shift_JIS を使えるようにするための初期化です（.NET では標準で登録されていないため）。
    ''' </summary>
    Shared Sub New()
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
    End Sub

    ''' <summary>
    ''' CSVファイルを読み込みます。
    ''' </summary>
    ''' <param name="path">CSVファイルのパス</param>
    Public Function ReadFile(path As String) As CsvReadResult
        Dim bytes = File.ReadAllBytes(path)
        Return ReadBytes(bytes)
    End Function

    ''' <summary>
    ''' CSVファイルの中身（バイト列）を、文字コードを判定してから読み込みます。
    ''' </summary>
    Public Function ReadBytes(bytes As Byte()) As CsvReadResult
        Dim encodingName As String = ""
        Dim text = DecodeText(bytes, encodingName)
        Dim result = ReadText(text)
        result.EncodingName = encodingName
        Return result
    End Function

    ''' <summary>
    ''' バイト列の文字コードを判定して文字列に変換します。
    ''' 1. 先頭にBOM（EF BB BF）があれば UTF-8
    ''' 2. UTF-8 として矛盾なく読めれば UTF-8（BOMなし）
    ''' 3. どちらでもなければ Shift_JIS（Windowsの日本語で使われるCP932）
    ''' </summary>
    ''' <param name="bytes">ファイルの中身</param>
    ''' <param name="encodingName">判定した文字コードの名前を返します</param>
    Public Shared Function DecodeText(bytes As Byte(), ByRef encodingName As String) As String
        If bytes.Length >= 3 AndAlso bytes(0) = &HEF AndAlso bytes(1) = &HBB AndAlso bytes(2) = &HBF Then
            encodingName = "UTF-8（BOM付き）"
            Return New UTF8Encoding(False).GetString(bytes, 3, bytes.Length - 3)
        End If

        Try
            ' 第2引数 True：UTF-8 として不正なバイトがあれば例外にする
            Dim strictUtf8 As New UTF8Encoding(False, True)
            Dim text = strictUtf8.GetString(bytes)
            encodingName = "UTF-8（BOMなし）"
            Return text
        Catch ex As DecoderFallbackException
            encodingName = "Shift_JIS"
            Return Encoding.GetEncoding(932).GetString(bytes)
        End Try
    End Function

    ''' <summary>
    ''' CSVの文字列を読み込み、チェックを通った行とエラーの一覧を返します。
    ''' </summary>
    ''' <param name="text">CSV全体の文字列（1行目はヘッダー）</param>
    Public Function ReadText(text As String) As CsvReadResult
        Dim result As New CsvReadResult()
        Dim records = ParseCsv(text)

        If records.Count = 0 Then
            result.Errors.Add(New BomError(0, BomErrorKind.FormatError, "CSVにデータがありません。"))
            Return result
        End If

        ' 1件目はヘッダー。列数だけ確認します。
        Dim header = records(0)
        If header.Fields.Count < RequiredColumnCount Then
            result.Errors.Add(New BomError(header.LineNumber, BomErrorKind.FormatError,
                $"ヘッダーの列数が{header.Fields.Count}列しかありません（{RequiredColumnCount}列必要です）。2行目以降は列の順番どおりに読み込みます。",
                header.RawText))
        End If

        ' 重複チェック用：「親品番 + 子品番」→ 最初に出てきた行番号
        Dim firstLineOfPair As New Dictionary(Of String, Integer)(StringComparer.Ordinal)

        For i = 1 To records.Count - 1
            Dim record = records(i)
            result.DataLineCount += 1

            Dim row = ValidateRecord(record, result.Errors)
            If row Is Nothing Then Continue For

            ' 同じ親子の組み合わせは、最初の行だけを使い、2件目以降はエラーにして読み飛ばします。
            Dim key = row.ParentCode & vbTab & row.ChildCode
            Dim firstLine As Integer
            If firstLineOfPair.TryGetValue(key, firstLine) Then
                result.Errors.Add(New BomError(record.LineNumber, BomErrorKind.DuplicateRelation,
                    $"親品番「{row.ParentCode}」と子品番「{row.ChildCode}」の組み合わせが{firstLine}行目と重複しています。この行は読み飛ばし、{firstLine}行目を使います。",
                    record.RawText))
                Continue For
            End If
            firstLineOfPair.Add(key, record.LineNumber)

            result.Rows.Add(row)
        Next

        Return result
    End Function

    ''' <summary>
    ''' 1行分の項目をチェックし、問題がなければ BomRow を返します。
    ''' 問題があればエラーを追加して Nothing を返します（1行に複数の問題があればすべて記録します）。
    ''' </summary>
    Private Function ValidateRecord(record As CsvRecord, errors As List(Of BomError)) As BomRow
        Dim f = record.Fields

        ' 列数のチェック（多すぎる列は無視します）
        If f.Count < RequiredColumnCount Then
            errors.Add(New BomError(record.LineNumber, BomErrorKind.FormatError,
                $"列が不足しています（{RequiredColumnCount}列必要ですが{f.Count}列です）。この行は読み飛ばします。",
                record.RawText))
            Return Nothing
        End If

        Dim hasError = False

        ' 必須列が空でないか（6列すべて必須）
        Dim emptyColumns = Enumerable.Range(0, RequiredColumnCount).
            Where(Function(c) f(c).Length = 0).
            Select(Function(c) ColumnNames(c)).
            ToList()
        If emptyColumns.Count > 0 Then
            errors.Add(New BomError(record.LineNumber, BomErrorKind.MissingRequired,
                $"必須列が空です：{String.Join("、", emptyColumns)}。この行は読み飛ばします。",
                record.RawText))
            hasError = True
        End If

        ' 員数：数値で、0より大きいこと（空の場合は上の必須チェックで記録済み）
        Dim quantity As Decimal
        Dim quantityText = f(3)
        If quantityText.Length > 0 Then
            If Not TryParseQuantity(quantityText, quantity) Then
                errors.Add(New BomError(record.LineNumber, BomErrorKind.InvalidQuantity,
                    $"員数「{quantityText}」が数値ではありません。この行は読み飛ばします。",
                    record.RawText))
                hasError = True
            ElseIf quantity <= 0D Then
                errors.Add(New BomError(record.LineNumber, BomErrorKind.InvalidQuantity,
                    $"員数「{quantityText}」が0以下です。この行は読み飛ばします。",
                    record.RawText))
                hasError = True
            End If
        End If

        ' 手配区分：「購入」か「内製」
        Dim procurement = f(5)
        If procurement.Length > 0 AndAlso Not ProcurementTypes.Contains(procurement) Then
            errors.Add(New BomError(record.LineNumber, BomErrorKind.InvalidProcurement,
                $"手配区分「{procurement}」は使えません（{String.Join("／", ProcurementTypes)}のどちらかにしてください）。この行は読み飛ばします。",
                record.RawText))
            hasError = True
        End If

        If hasError Then Return Nothing

        Return New BomRow With {
            .LineNumber = record.LineNumber,
            .ParentCode = f(0),
            .ChildCode = f(1),
            .ItemName = f(2),
            .Quantity = quantity,
            .Unit = f(4),
            .Procurement = procurement
        }
    End Function

    ''' <summary>
    ''' 員数の文字列を数値に変換します。全角数字（「２」など）も半角にしてから変換します。
    ''' </summary>
    Friend Shared Function TryParseQuantity(text As String, ByRef value As Decimal) As Boolean
        Dim normalized = text.Normalize(NormalizationForm.FormKC).Trim()
        Return Decimal.TryParse(normalized, NumberStyles.Number, CultureInfo.InvariantCulture, value)
    End Function

    ''' <summary>
    ''' CSVの1レコード（解析途中のデータ）です。
    ''' </summary>
    Friend Class CsvRecord
        ''' <summary>レコードが始まる行番号</summary>
        Public Property LineNumber As Integer
        ''' <summary>各列の値（前後の空白は取り除き済み）</summary>
        Public Property Fields As New List(Of String)
        ''' <summary>元の行の文字列（エラーシートに表示します）</summary>
        Public Property RawText As String = ""
    End Class

    ''' <summary>
    ''' CSVの文字列をレコードの一覧に分解します。
    ''' ・ダブルクォートで囲まれた値（カンマや改行を含む値、"" による " の表現）に対応します。
    ''' ・改行は CRLF / LF / CR のどれでも構いません。
    ''' ・空行は読み飛ばします（行番号は数えます）。
    ''' </summary>
    Friend Shared Function ParseCsv(text As String) As List(Of CsvRecord)
        Dim records As New List(Of CsvRecord)
        Dim fields As New List(Of String)
        Dim value As New StringBuilder()
        Dim inQuotes = False
        Dim line = 1                ' 現在の行番号
        Dim recordStartLine = 1     ' 現在のレコードが始まった行番号
        Dim recordStartIndex = 0    ' 現在のレコードが始まった文字位置

        ' 現在のレコードを確定して一覧に追加し、次のレコードの準備をします。
        Dim endRecord =
            Sub(endIndex As Integer)
                fields.Add(value.ToString().Trim())
                value.Clear()
                ' 全列が空（空行）のレコードは追加しません
                If fields.Exists(Function(x) x.Length > 0) Then
                    records.Add(New CsvRecord With {
                        .LineNumber = recordStartLine,
                        .Fields = New List(Of String)(fields),
                        .RawText = text.Substring(recordStartIndex, endIndex - recordStartIndex).TrimEnd(ControlChars.Cr, ControlChars.Lf)
                    })
                End If
                fields.Clear()
            End Sub

        Dim i = 0
        While i < text.Length
            Dim c = text(i)
            If inQuotes Then
                If c = """"c Then
                    If i + 1 < text.Length AndAlso text(i + 1) = """"c Then
                        value.Append(""""c)   ' "" は " 1文字として扱う
                        i += 1
                    Else
                        inQuotes = False
                    End If
                Else
                    If c = ControlChars.Lf OrElse (c = ControlChars.Cr AndAlso (i + 1 >= text.Length OrElse text(i + 1) <> ControlChars.Lf)) Then
                        line += 1
                    End If
                    value.Append(c)
                End If
            Else
                Select Case c
                    Case """"c
                        inQuotes = True
                    Case ","c
                        fields.Add(value.ToString().Trim())
                        value.Clear()
                    Case ControlChars.Cr, ControlChars.Lf
                        ' CRLF は1つの改行として扱う
                        If c = ControlChars.Cr AndAlso i + 1 < text.Length AndAlso text(i + 1) = ControlChars.Lf Then
                            i += 1
                        End If
                        endRecord(i)
                        line += 1
                        recordStartLine = line
                        recordStartIndex = i + 1
                    Case Else
                        value.Append(c)
                End Select
            End If
            i += 1
        End While

        ' 最後の行が改行で終わっていない場合
        If value.Length > 0 OrElse fields.Count > 0 Then
            endRecord(text.Length)
        End If

        Return records
    End Function

End Class
