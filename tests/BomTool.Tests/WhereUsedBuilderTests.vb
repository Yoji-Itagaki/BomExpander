Imports BomTool.Core
Imports Xunit

''' <summary>
''' WhereUsedBuilder（部品から製品までの逆展開）のテストです。
''' </summary>
Public Class WhereUsedBuilderTests

    ''' <summary>指定した部品の逆展開の行を取り出します。</summary>
    Private Shared Function RowsOf(expansion As BomExpansionResult, partCode As String) As List(Of WhereUsedLine)
        Return expansion.WhereUsed.Where(Function(l) l.PartCode = partCode).ToList()
    End Function

    <Fact>
    Public Sub 部品から親をたどり製品で止まる()
        ' P → A(2) → B(3) → C(4)
        Dim r = ReadAndExpand(Csv(
            "P,A,組品A,2,個,内製",
            "A,B,組品B,3,個,内製",
            "B,C,部品C,4,本,購入"))

        Dim rows = RowsOf(r.Expansion, "C")
        Assert.Equal({("C", 0), ("B", 1), ("A", 2), ("P", 3)}, rows.Select(Function(l) (l.ItemCode, l.Level)))

        ' レベル0は部品自身：員数なし、累計員数1
        Assert.Null(rows(0).Quantity)
        Assert.Equal(1D, rows(0).CumulativeQuantity)
        ' 員数は「この品番1個に使う、1つ下の品番の数」、累計員数は「この品番1個に使う部品Cの数」
        Assert.Equal(4D, rows(1).Quantity)
        Assert.Equal(4D, rows(1).CumulativeQuantity)
        Assert.Equal(3D, rows(2).Quantity)
        Assert.Equal(12D, rows(2).CumulativeQuantity)
        Assert.Equal(2D, rows(3).Quantity)
        Assert.Equal(24D, rows(3).CumulativeQuantity)

        ' 製品の行には印が付き、品名・手配区分は空。単位はすべて部品Cの単位
        Assert.True(rows(3).IsProduct)
        Assert.Equal("", rows(3).ItemName)
        Assert.Equal("", rows(3).Procurement)
        Assert.All(rows, Sub(l) Assert.Equal("本", l.Unit))
        Assert.All(rows, Sub(l) Assert.Equal("部品C", l.PartName))
        Assert.Equal("組品B", rows(1).ItemName)
        Assert.Equal("内製", rows(1).Procurement)
    End Sub

    <Fact>
    Public Sub 製品以外のすべての品番を品番の順に逆展開する()
        Dim r = ReadAndExpand(Csv(
            "P,Z,組品Z,1,個,内製",
            "Z,B,部品B,1,個,購入",
            "P,A,部品A,1,個,購入"))

        Dim parts = r.Expansion.WhereUsed.Where(Function(l) l.Level = 0).Select(Function(l) l.PartCode).ToArray()
        Assert.Equal({"A", "B", "Z"}, parts)
        Assert.DoesNotContain(r.Expansion.WhereUsed, Function(l) l.PartCode = "P")
    End Sub

    <Fact>
    Public Sub 複数の組品と製品で使われる部品はすべての経路を出す()
        ' S は A と B の両方で使われ、A は P1 と P2 の両方で使われる
        Dim r = ReadAndExpand(Csv(
            "P1,A,組品A,2,個,内製",
            "P1,B,組品B,1,個,内製",
            "P2,A,組品A,3,個,内製",
            "A,S,ねじ,4,本,購入",
            "B,S,ねじ,5,本,購入"))

        Dim rows = RowsOf(r.Expansion, "S")
        Dim actual = rows.Select(Function(l) (l.Level, l.ItemCode, l.CumulativeQuantity)).ToArray()
        Assert.Equal({
            (0, "S", 1D),
            (1, "A", 4D),
            (2, "P1", 8D),
            (2, "P2", 12D),
            (1, "B", 5D),
            (2, "P1", 5D)
        }, actual)

        ' 製品ごとに合計すると、部品集計の合計員数と一致する
        For Each summary In r.Expansion.Summary.Where(Function(s) s.ItemCode = "S")
            Dim total = rows.Where(Function(l) l.IsProduct AndAlso l.ItemCode = summary.ProductCode).
                Sum(Function(l) l.CumulativeQuantity)
            Assert.Equal(summary.TotalQuantity, total)
        Next
    End Sub

    <Fact>
    Public Sub 製品の直下の部品は製品の行だけを出す()
        Dim r = ReadAndExpand(Csv("P,A,部品A,3,個,購入"))

        Dim rows = RowsOf(r.Expansion, "A")
        Assert.Equal({("A", 0), ("P", 1)}, rows.Select(Function(l) (l.ItemCode, l.Level)))
        Assert.True(rows(1).IsProduct)
        Assert.Equal(3D, rows(1).CumulativeQuantity)
    End Sub

    <Fact>
    Public Sub 循環参照があっても止まらず製品までの経路だけを出す()
        ' P1 → A → B → A、P2 → B → A → B のループ。どちらの経路も展開時に途中で止まる
        Dim r = ReadAndExpand(Csv(
            "P1,A,組品A,1,個,内製",
            "P2,B,組品B,1,個,内製",
            "A,B,組品B,2,個,内製",
            "B,A,組品A,3,個,内製"))

        Assert.Contains(r.Expansion.Errors, Function(e) e.Kind = BomErrorKind.CircularReference)

        Dim rows = RowsOf(r.Expansion, "A")
        ' すべての経路が製品で終わる
        Dim lastRows = rows.Where(Function(l, i) i = rows.Count - 1 OrElse rows(i + 1).Level <= l.Level).ToList()
        Assert.All(lastRows.Where(Function(l) l.Level > 0), Sub(l) Assert.True(l.IsProduct))
        Assert.Contains(rows, Function(l) l.IsProduct AndAlso l.ItemCode = "P1")
        Assert.Contains(rows, Function(l) l.IsProduct AndAlso l.ItemCode = "P2")
    End Sub

    <Fact>
    Public Sub 部品がなければ逆展開も空になる()
        Dim r = ReadAndExpand(Csv(
            "A,B,組品B,1,個,内製",
            "B,A,組品A,1,個,内製"))

        Assert.Empty(r.Expansion.WhereUsed)
    End Sub

    <Fact>
    Public Sub 逆展開の行が上限を超えたら例外にする()
        Dim lines = New BomTreeExpander().Expand(New CsvBomReader().ReadText(Csv(
            "P,A,組品A,1,個,内製",
            "A,B,部品B,1,個,購入")).Rows).Lines

        ' A：2行、B：3行 → 合計5行
        Assert.Equal(5, WhereUsedBuilder.Build(lines, 5).Count)
        Assert.Throws(Of InvalidOperationException)(Function() WhereUsedBuilder.Build(lines, 4))
    End Sub

End Class
