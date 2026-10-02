Imports System.Text
Imports BomTool.Core
Imports Xunit

''' <summary>
''' CsvBomReader（文字コード判定・CSV解析・エラー行の扱い）のテストです。
''' </summary>
Public Class CsvBomReaderTests

    Shared Sub New()
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
    End Sub

#Region "エラー行の扱い"

    <Theory>
    <InlineData("abc", "数値ではありません")>
    <InlineData("0", "0以下")>
    <InlineData("-1", "0以下")>
    <InlineData("1個", "数値ではありません")>
    Public Sub 員数が不正な行はエラーにして読み飛ばす(quantity As String, expectedMessage As String)
        Dim result = New CsvBomReader().ReadText(Csv(
            "P,A,部品A,1,個,購入",
            $"P,B,部品B,{quantity},個,購入"))

        Assert.Equal({"A"}, result.Rows.Select(Function(r) r.ChildCode))
        Dim errorItem = Assert.Single(result.Errors)
        Assert.Equal(BomErrorKind.InvalidQuantity, errorItem.Kind)
        Assert.Equal(3, errorItem.LineNumber)
        Assert.Contains(expectedMessage, errorItem.Message)
    End Sub

    <Fact>
    Public Sub 必須列が空の行はエラーにして空の列名を示す()
        Dim result = New CsvBomReader().ReadText(Csv(
            "P,,部品A,1,個,購入",
            ",B,,1,,購入"))

        Assert.Empty(result.Rows)
        Assert.Equal(2, result.Errors.Count)
        Assert.All(result.Errors, Sub(e) Assert.Equal(BomErrorKind.MissingRequired, e.Kind))
        Assert.Contains("子品番", result.Errors(0).Message)
        Assert.Contains("親品番、品名、単位", result.Errors(1).Message)
    End Sub

    <Fact>
    Public Sub 列が足りない行はエラーにする()
        Dim result = New CsvBomReader().ReadText(Csv("P,A,部品A,1"))

        Assert.Empty(result.Rows)
        Assert.Equal(BomErrorKind.FormatError, Assert.Single(result.Errors).Kind)
    End Sub

    <Fact>
    Public Sub 手配区分が購入_内製以外の行はエラーにする()
        Dim result = New CsvBomReader().ReadText(Csv("P,A,部品A,1,個,外注"))

        Assert.Empty(result.Rows)
        Assert.Equal(BomErrorKind.InvalidProcurement, Assert.Single(result.Errors).Kind)
    End Sub

    <Fact>
    Public Sub 同じ親子の組み合わせの重複は2件目以降をエラーにして最初の行を使う()
        Dim result = New CsvBomReader().ReadText(Csv(
            "P,A,部品A,2,個,購入",
            "P,B,部品B,1,個,購入",
            "P,A,部品A,5,個,購入"))

        Assert.Equal(2, result.Rows.Count)
        Assert.Equal(2D, result.Rows.Single(Function(r) r.ChildCode = "A").Quantity)
        Dim errorItem = Assert.Single(result.Errors)
        Assert.Equal(BomErrorKind.DuplicateRelation, errorItem.Kind)
        Assert.Equal(4, errorItem.LineNumber)
        Assert.Contains("2行目", errorItem.Message)
    End Sub

    <Fact>
    Public Sub 子品番が同じでも親品番が違えば重複ではない()
        Dim result = New CsvBomReader().ReadText(Csv(
            "P,S,ねじ,1,本,購入",
            "A,S,ねじ,1,本,購入"))

        Assert.Empty(result.Errors)
        Assert.Equal(2, result.Rows.Count)
    End Sub

    <Fact>
    Public Sub エラー行があっても正しい行はすべて読み込む()
        Dim result = New CsvBomReader().ReadText(Csv(
            "P,A,部品A,1,個,購入",
            "P,B,部品B,x,個,購入",
            "",
            "P,C,部品C,3,個,購入"))

        Assert.Equal({"A", "C"}, result.Rows.Select(Function(r) r.ChildCode))
        Assert.Equal(5, result.Rows(1).LineNumber)   ' 空行も行番号には数える
        Assert.Equal(3, result.DataLineCount)        ' 空行はデータ行に数えない
    End Sub

#End Region

#Region "CSVの解析"

    <Fact>
    Public Sub ダブルクォートで囲まれたカンマや引用符を正しく読む()
        Dim result = New CsvBomReader().ReadText(Csv(
            "P,A,""ケーブル（2m,黒）"",1,本,購入",
            "P,B,""通称""""ミニ"""""",1,個,購入"))

        Assert.Empty(result.Errors)
        Assert.Equal("ケーブル（2m,黒）", result.Rows(0).ItemName)
        Assert.Equal("通称""ミニ""", result.Rows(1).ItemName)
    End Sub

    <Fact>
    Public Sub 前後の空白と全角数字の員数を受け付ける()
        Dim result = New CsvBomReader().ReadText(Csv(" P , A ,部品A, １２ ,個,購入"))

        Dim row = Assert.Single(result.Rows)
        Assert.Equal("P", row.ParentCode)
        Assert.Equal("A", row.ChildCode)
        Assert.Equal(12D, row.Quantity)
    End Sub

    <Fact>
    Public Sub 改行がLFだけのファイルも読める()
        Dim text = Header & vbLf & "P,A,部品A,1,個,購入" & vbLf & "P,B,部品B,2,個,購入"
        Dim result = New CsvBomReader().ReadText(text)

        Assert.Equal(2, result.Rows.Count)
        Assert.Equal(3, result.Rows(1).LineNumber)
    End Sub

    <Fact>
    Public Sub 空のファイルはエラーにする()
        Dim result = New CsvBomReader().ReadText("")
        Assert.Equal(BomErrorKind.FormatError, Assert.Single(result.Errors).Kind)
    End Sub

#End Region

#Region "文字コードの判定"

    Private Shared ReadOnly SampleText As String = Csv("製品Ａ,部品①,ねじ M3x8,4,本,購入")

    <Fact>
    Public Sub UTF8_BOM付きを読める()
        Dim bytes = New UTF8Encoding(True).GetPreamble().Concat(Encoding.UTF8.GetBytes(SampleText)).ToArray()
        AssertDecoded(New CsvBomReader().ReadBytes(bytes), "UTF-8（BOM付き）")
    End Sub

    <Fact>
    Public Sub UTF8_BOMなしを読める()
        Dim bytes = New UTF8Encoding(False).GetBytes(SampleText)
        AssertDecoded(New CsvBomReader().ReadBytes(bytes), "UTF-8（BOMなし）")
    End Sub

    <Fact>
    Public Sub ShiftJISを読める()
        Dim bytes = Encoding.GetEncoding(932).GetBytes(SampleText)
        AssertDecoded(New CsvBomReader().ReadBytes(bytes), "Shift_JIS")
    End Sub

    Private Shared Sub AssertDecoded(result As CsvReadResult, expectedEncoding As String)
        Assert.Equal(expectedEncoding, result.EncodingName)
        Assert.Empty(result.Errors)
        Dim row = Assert.Single(result.Rows)
        Assert.Equal("製品Ａ", row.ParentCode)
        Assert.Equal("部品①", row.ChildCode)
        Assert.Equal("本", row.Unit)
        Assert.Equal("購入", row.Procurement)
    End Sub

#End Region

End Class
