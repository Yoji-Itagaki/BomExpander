Imports BomTool.Core
Imports Xunit

''' <summary>
''' BomTreeExpander（多段階展開・累計員数・循環参照・集計）のテストです。
''' </summary>
Public Class BomTreeExpanderTests

#Region "累計員数の計算"

    <Fact>
    Public Sub 累計員数は上位の員数をすべて掛け合わせた値になる()
        ' P → A(2) → B(3) → C(4)：C の累計員数は 2×3×4 = 24
        Dim r = ReadAndExpand(Csv(
            "P,A,組品A,2,個,内製",
            "A,B,組品B,3,個,内製",
            "B,C,部品C,4,個,購入"))

        Assert.Empty(r.Expansion.Errors)
        Assert.Equal(2D, FindLine(r.Expansion, "P", "A").CumulativeQuantity)
        Assert.Equal(6D, FindLine(r.Expansion, "P", "B").CumulativeQuantity)
        Assert.Equal(24D, FindLine(r.Expansion, "P", "C").CumulativeQuantity)
        ' 員数（直上の親あたり）はCSVの値のまま
        Assert.Equal(4D, FindLine(r.Expansion, "P", "C").Quantity)
    End Sub

    <Fact>
    Public Sub 小数の員数も正しく掛け合わせる()
        ' 0.5 × 1.5 × 3 = 2.25（Decimal で計算するので誤差が出ない）
        Dim r = ReadAndExpand(Csv(
            "P,A,組品A,0.5,個,内製",
            "A,B,ケーブル,1.5,m,内製",
            "B,C,端子,3,個,購入"))

        Assert.Equal(0.75D, FindLine(r.Expansion, "P", "B").CumulativeQuantity)
        Assert.Equal(2.25D, FindLine(r.Expansion, "P", "C").CumulativeQuantity)
    End Sub

    <Fact>
    Public Sub レベルは製品が0でその下に1ずつ増える()
        Dim r = ReadAndExpand(Csv(
            "P,A,組品A,1,個,内製",
            "A,B,組品B,1,個,内製",
            "B,C,組品C,1,個,内製",
            "C,D,部品D,1,個,購入"))

        Dim levels = r.Expansion.Lines.Select(Function(l) (l.ItemCode, l.Level)).ToList()
        Assert.Equal({("P", 0), ("A", 1), ("B", 2), ("C", 3), ("D", 4)}, levels)
    End Sub

    <Fact>
    Public Sub 子はCSVの順番どおりに深さ優先で並ぶ()
        Dim r = ReadAndExpand(Csv(
            "P,A,組品A,1,個,内製",
            "P,B,部品B,1,個,購入",
            "A,A1,部品A1,1,個,購入",
            "A,A2,部品A2,1,個,購入"))

        Dim codes = r.Expansion.Lines.Select(Function(l) l.ItemCode).ToArray()
        Assert.Equal({"P", "A", "A1", "A2", "B"}, codes)
    End Sub

    <Fact>
    Public Sub どの子品番にもなっていない親品番を製品として扱う()
        Dim r = ReadAndExpand(Csv(
            "P1,A,共通組品,1,個,内製",
            "A,X,部品X,2,個,購入",
            "P2,A,共通組品,3,個,内製"))

        Assert.Equal({"P1", "P2"}, r.Expansion.Products)
        ' 共通の組品 A は両方の製品の下に展開され、製品ごとに累計員数が変わる
        Assert.Equal(2D, FindLine(r.Expansion, "P1", "X").CumulativeQuantity)
        Assert.Equal(6D, FindLine(r.Expansion, "P2", "X").CumulativeQuantity)
        Assert.Equal(4, r.Expansion.PartLineCount)
    End Sub

#End Region

#Region "部品集計"

    <Fact>
    Public Sub 同じ品番の累計員数は製品ごとに合算される()
        ' ねじ S は A の下に 2×4=8本、B の下に 1×3=3本、P の直下に 1本 → 合計 12本
        Dim r = ReadAndExpand(Csv(
            "P,A,組品A,2,個,内製",
            "P,B,組品B,1,個,内製",
            "P,S,ねじ,1,本,購入",
            "A,S,ねじ,4,本,購入",
            "B,S,ねじ,3,本,購入"))

        Dim screw = r.Expansion.Summary.Single(Function(s) s.ProductCode = "P" AndAlso s.ItemCode = "S")
        Assert.Equal(12D, screw.TotalQuantity)
        Assert.Equal(3, screw.UsageCount)
    End Sub

    <Fact>
    Public Sub 集計は製品ごとに手配区分_品番の順に並ぶ()
        Dim r = ReadAndExpand(Csv(
            "P,Z,組品Z,1,個,内製",
            "P,B,部品B,1,個,購入",
            "P,A,部品A,1,個,購入",
            "Q,C,部品C,1,個,購入"))

        Dim order = r.Expansion.Summary.Select(Function(s) $"{s.ProductCode}:{s.Procurement}:{s.ItemCode}").ToArray()
        Assert.Equal({"P:購入:A", "P:購入:B", "P:内製:Z", "Q:購入:C"}, order)
    End Sub

#End Region

#Region "循環参照の検出"

    <Fact>
    Public Sub 循環参照を検出して展開を止める()
        ' P → A → B → A のループ
        Dim r = ReadAndExpand(Csv(
            "P,A,組品A,1,個,内製",
            "A,B,組品B,1,個,内製",
            "B,A,組品A,1,個,内製",
            "B,C,部品C,2,個,購入"))

        Dim cycleError = Assert.Single(r.Expansion.Errors)
        Assert.Equal(BomErrorKind.CircularReference, cycleError.Kind)
        Assert.Equal(4, cycleError.LineNumber)               ' ループを閉じている行（B→A）
        Assert.Contains("A → B → A", cycleError.Message)

        ' ループの手前までは展開され、ループ以外の子（C）も展開される
        Dim codes = r.Expansion.Lines.Select(Function(l) l.ItemCode).ToArray()
        Assert.Equal({"P", "A", "B", "C"}, codes)
    End Sub

    <Fact>
    Public Sub 自分自身を子に持つ行も循環参照として検出する()
        Dim r = ReadAndExpand(Csv(
            "P,A,組品A,1,個,内製",
            "A,A,組品A,1,個,内製"))

        Dim cycleError = Assert.Single(r.Expansion.Errors)
        Assert.Equal(BomErrorKind.CircularReference, cycleError.Kind)
        Assert.Contains("A → A", cycleError.Message)
    End Sub

    <Fact>
    Public Sub 複数の製品から同じループに到達してもエラーは1件だけ記録する()
        Dim r = ReadAndExpand(Csv(
            "P1,A,組品A,1,個,内製",
            "P2,A,組品A,1,個,内製",
            "A,B,組品B,1,個,内製",
            "B,A,組品A,1,個,内製"))

        Assert.Single(r.Expansion.Errors, Function(e) e.Kind = BomErrorKind.CircularReference)
        Assert.Equal(2, r.Expansion.Products.Count)
    End Sub

    <Fact>
    Public Sub 製品からたどれない場所にあるループも検出する()
        ' X ⇄ Y は上に製品がないため展開されないが、エラーとして記録する
        Dim r = ReadAndExpand(Csv(
            "P,A,部品A,1,個,購入",
            "X,Y,組品Y,1,個,内製",
            "Y,X,組品X,1,個,内製"))

        Assert.Equal({"P"}, r.Expansion.Products)
        Dim cycleError = Assert.Single(r.Expansion.Errors)
        Assert.Equal(BomErrorKind.CircularReference, cycleError.Kind)
        Assert.Contains("X → Y → X", cycleError.Message)
    End Sub

    <Fact>
    Public Sub すべてがループしていて製品がない場合はエラーにする()
        Dim r = ReadAndExpand(Csv(
            "A,B,組品B,1,個,内製",
            "B,A,組品A,1,個,内製"))

        Assert.Empty(r.Expansion.Products)
        Assert.Empty(r.Expansion.Lines)
        Assert.Contains(r.Expansion.Errors, Function(e) e.Kind = BomErrorKind.NoTopLevelItem)
        Assert.Contains(r.Expansion.Errors, Function(e) e.Kind = BomErrorKind.CircularReference)
    End Sub

    <Fact>
    Public Sub 同じ部品を複数の組品で使うだけなら循環参照にならない()
        ' A と B の両方が C を使う（ひし形の構成）のはループではない
        Dim r = ReadAndExpand(Csv(
            "P,A,組品A,1,個,内製",
            "P,B,組品B,1,個,内製",
            "A,C,部品C,1,個,購入",
            "B,C,部品C,1,個,購入"))

        Assert.Empty(r.Expansion.Errors)
        Assert.Equal(2, r.Expansion.Lines.Where(Function(l) l.ItemCode = "C").Count())
    End Sub

#End Region

    <Fact>
    Public Sub 展開行が上限を超えたら例外にする()
        Dim rows = New CsvBomReader().ReadText(Csv(
            "P,A,部品A,1,個,購入",
            "P,B,部品B,1,個,購入",
            "P,C,部品C,1,個,購入")).Rows
        Dim expander As New BomTreeExpander With {.MaxLines = 3}

        Assert.Throws(Of InvalidOperationException)(Function() expander.Expand(rows))
    End Sub

End Class
