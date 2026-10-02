Imports System.IO
Imports BomTool.Core
Imports Xunit

''' <summary>
''' samples フォルダーのダミーCSVを実際に読み込み、展開結果が想定どおりかを確認するテストです。
''' </summary>
Public Class SampleFileTests

    Private Shared Function SamplePath(fileName As String) As String
        Return Path.Combine(AppContext.BaseDirectory, "samples", fileName)
    End Function

    Private Shared Function Run(fileName As String) As (Read As CsvReadResult, Expansion As BomExpansionResult)
        Dim readResult = New CsvBomReader().ReadFile(SamplePath(fileName))
        Return (readResult, New BomTreeExpander().Expand(readResult.Rows))
    End Function

    <Fact>
    Public Sub 正常なサンプルはエラーなしで2製品_4階層に展開される()
        Dim r = Run("sample_normal.csv")

        Assert.Equal("UTF-8（BOM付き）", r.Read.EncodingName)
        Assert.Equal(30, r.Read.DataLineCount)
        Assert.Empty(r.Read.Errors)
        Assert.Empty(r.Expansion.Errors)
        Assert.Equal({"P-100", "P-200"}, r.Expansion.Products)
        Assert.Equal(38, r.Expansion.PartLineCount)
        Assert.Equal(4, r.Expansion.Lines.Max(Function(l) l.Level))

        ' P-200 はファンモジュール A-110 を2個使うため、その下のベアリングは 2×1×1×2 = 4個
        Dim bearing = r.Expansion.Lines.Single(Function(l) l.ProductCode = "P-200" AndAlso l.ItemCode = "C-222")
        Assert.Equal(4, bearing.Level)
        Assert.Equal(4D, bearing.CumulativeQuantity)

        ' P-100 のねじ S-001 はモーター（4本）と制御基板（2本）の合計 6本
        Dim screw = r.Expansion.Summary.Single(Function(s) s.ProductCode = "P-100" AndAlso s.ItemCode = "S-001")
        Assert.Equal(6D, screw.TotalQuantity)
        Assert.Equal(2, screw.UsageCount)

        ' 小数の員数（配線ケーブル 1.5m）
        Assert.Equal(1.5D, FindLine(r.Expansion, "P-200", "C-413").CumulativeQuantity)
    End Sub

    <Fact>
    Public Sub エラーを含むサンプルは各エラーを記録し正しい行だけ展開する()
        Dim r = Run("sample_errors.csv")
        Dim allErrors = r.Read.Errors.Concat(r.Expansion.Errors).ToList()

        Assert.Equal("Shift_JIS", r.Read.EncodingName)

        ' 行番号とエラーの種類の組み合わせ
        Dim actual = allErrors.Select(Function(e) (e.LineNumber, e.Kind)).OrderBy(Function(x) x.LineNumber).ToList()
        Dim expected = {
            (3, BomErrorKind.InvalidQuantity),       ' abc
            (4, BomErrorKind.InvalidQuantity),       ' 0
            (5, BomErrorKind.InvalidQuantity),       ' -2
            (6, BomErrorKind.MissingRequired),       ' 子品番が空
            (7, BomErrorKind.MissingRequired),       ' 品名が空
            (8, BomErrorKind.InvalidProcurement),    ' 外注
            (10, BomErrorKind.DuplicateRelation),    ' X-110 → X-111 の重複
            (13, BomErrorKind.CircularReference),    ' X-170 → X-180 → X-170
            (15, BomErrorKind.FormatError),          ' 列不足
            (18, BomErrorKind.CircularReference)     ' Z-10 → Z-20 → Z-10（製品からたどれないループ）
        }
        Assert.Equal(expected, actual)

        ' 製品は X-100 だけ（Z-10 / Z-20 はお互いの子になっているため製品ではない）
        Assert.Equal({"X-100"}, r.Expansion.Products)
        Dim codes = r.Expansion.Lines.Select(Function(l) l.ItemCode).ToArray()
        Assert.Equal({"X-100", "X-110", "X-111", "X-170", "X-180", "X-181", "X-195"}, codes)

        ' 重複行は最初の行（員数2）を使う。クォート内のカンマと全角数字の員数も読める
        Assert.Equal(2D, FindLine(r.Expansion, "X-100", "X-111").CumulativeQuantity)
        Assert.Equal(8D, FindLine(r.Expansion, "X-100", "X-181").CumulativeQuantity)
        Assert.Equal("部品J（φ5,黒）", FindLine(r.Expansion, "X-100", "X-195").ItemName)
        Assert.Equal(2D, FindLine(r.Expansion, "X-100", "X-195").Quantity)
    End Sub

End Class
