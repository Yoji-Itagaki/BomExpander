''' <summary>
''' 親子関係の一覧（BomRow）から、製品ごとの多段階部品表を作るクラスです。
'''
''' ・最上位の品番（どの行の子品番にもなっていない親品番）を製品として扱います。
''' ・製品から下へ順にたどり（深さ優先）、累計員数 ＝ 上位の員数をすべて掛け合わせた値 を計算します。
''' ・循環参照（A→B→A など）を見つけたら、その経路の展開を止めてエラーに記録します。
''' ・展開結果から、製品ごと・品番ごとの集計も作ります。
''' </summary>
Public Class BomTreeExpander

    ''' <summary>
    ''' 展開行の上限です。Excelの1シートの行数上限（約104万行）を超えないようにするためのものです。
    ''' </summary>
    Public Property MaxLines As Integer = 1_000_000

    ''' <summary>
    ''' 部品表を多段階に展開します。
    ''' </summary>
    ''' <param name="rows">CsvBomReader で読み込んだ行（チェック済みのもの）</param>
    Public Function Expand(rows As IEnumerable(Of BomRow)) As BomExpansionResult
        Dim result As New BomExpansionResult()
        Dim rowList = rows.ToList()
        Dim context As New ExpandContext(rowList, result.Errors)

        ' 1) 最上位の品番（親品番として出てくるが、どこにも子品番として出てこないもの）を探す
        Dim topCodes = context.ParentOrder.Where(Function(p) Not context.ChildCodes.Contains(p)).ToList()
        If rowList.Count > 0 AndAlso topCodes.Count = 0 Then
            result.Errors.Add(New BomError(0, BomErrorKind.NoTopLevelItem,
                "最上位の品番（製品）が見つかりません。すべての品番が別の品番の子になっており、親子関係がループしている可能性があります。"))
        End If

        ' 2) 製品ごとに展開する
        For Each topCode In topCodes
            result.Products.Add(topCode)
            result.Lines.Add(New ExpandedLine With {
                .ProductCode = topCode,
                .Level = 0,
                .ItemCode = topCode,
                .Quantity = 1D,
                .CumulativeQuantity = 1D
            })

            ' 今たどっている経路（製品 → … → 現在の品番）。循環参照の検出に使います。
            Dim path As New List(Of String) From {topCode}
            ExpandChildren(topCode, topCode, 1, 1D, path, context, result.Lines)
        Next

        ' 3) 製品から到達できない部分にあるループも検出して記録する
        DetectUnreachableCycles(context)

        ' 4) 部品集計を作る
        result.Summary = Summarize(result.Lines, result.Products)
        Return result
    End Function

    ''' <summary>
    ''' 親品番の子を順番に出力し、さらにその子へ再帰的に下りていきます。
    ''' </summary>
    ''' <param name="productCode">展開中の製品</param>
    ''' <param name="parentCode">子を展開する親品番</param>
    ''' <param name="level">子のレベル</param>
    ''' <param name="parentCumulative">親の累計員数（製品1台あたりの親の数）</param>
    ''' <param name="path">製品から親品番までの経路</param>
    Private Sub ExpandChildren(productCode As String, parentCode As String, level As Integer,
                               parentCumulative As Decimal, path As List(Of String),
                               context As ExpandContext, lines As List(Of ExpandedLine))
        Dim children As List(Of BomRow) = Nothing
        If Not context.ChildrenOf.TryGetValue(parentCode, children) Then Return ' 子がない＝末端の部品

        For Each row In children
            ' 経路の中にすでに同じ品番があれば循環参照。この先は展開しません。
            Dim loopStart = path.IndexOf(row.ChildCode)
            If loopStart >= 0 Then
                context.ReportCycle(path.Skip(loopStart).ToList(), row)
                Continue For
            End If

            If lines.Count >= MaxLines Then
                Throw New InvalidOperationException(
                    $"展開結果が{MaxLines:#,0}行を超えたため処理を中止しました。部品表の構成を確認してください。")
            End If

            Dim cumulative = parentCumulative * row.Quantity
            lines.Add(New ExpandedLine With {
                .ProductCode = productCode,
                .Level = level,
                .ItemCode = row.ChildCode,
                .ItemName = row.ItemName,
                .Quantity = row.Quantity,
                .CumulativeQuantity = cumulative,
                .Unit = row.Unit,
                .Procurement = row.Procurement,
                .SourceLineNumber = row.LineNumber
            })
            context.Visited.Add(row.ChildCode)

            path.Add(row.ChildCode)
            ExpandChildren(productCode, row.ChildCode, level + 1, cumulative, path, context, lines)
            path.RemoveAt(path.Count - 1)
        Next
    End Sub

    ''' <summary>
    ''' 製品から一度もたどらなかった品番について、ループがないかを調べます。
    ''' （例：X→Y→X だけの組で、上に製品がない場合）
    ''' </summary>
    Private Sub DetectUnreachableCycles(context As ExpandContext)
        ' 0：未訪問、1：調査中（経路上にある）、2：調査済み
        Dim state As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
        Dim path As New List(Of String)

        For Each code In context.ParentOrder
            If context.Visited.Contains(code) OrElse state.ContainsKey(code) Then Continue For
            VisitForCycle(code, state, path, context)
        Next
    End Sub

    ''' <summary>
    ''' 深さ優先でたどり、調査中の品番に戻ってきたらループとして記録します。
    ''' </summary>
    Private Sub VisitForCycle(code As String, state As Dictionary(Of String, Integer),
                              path As List(Of String), context As ExpandContext)
        state(code) = 1
        path.Add(code)

        Dim children As List(Of BomRow) = Nothing
        If context.ChildrenOf.TryGetValue(code, children) Then
            For Each row In children
                Dim childState As Integer
                state.TryGetValue(row.ChildCode, childState)
                If childState = 1 Then
                    context.ReportCycle(path.Skip(path.IndexOf(row.ChildCode)).ToList(), row)
                ElseIf childState = 0 AndAlso Not context.Visited.Contains(row.ChildCode) Then
                    VisitForCycle(row.ChildCode, state, path, context)
                End If
            Next
        End If

        path.RemoveAt(path.Count - 1)
        state(code) = 2
    End Sub

    ''' <summary>
    ''' 展開結果から、製品ごとに同じ品番の累計員数を合算します。
    ''' 並び順：製品の順 → 手配区分（購入 → 内製）→ 品番
    ''' </summary>
    ''' <param name="lines">展開結果</param>
    ''' <param name="products">製品の並び順</param>
    Public Shared Function Summarize(lines As IEnumerable(Of ExpandedLine), products As IList(Of String)) As List(Of SummaryLine)
        Return lines.
            Where(Function(l) l.Level > 0).
            GroupBy(Function(l) (l.ProductCode, l.Procurement, l.ItemCode)).
            Select(Function(g) New SummaryLine With {
                .ProductCode = g.Key.ProductCode,
                .Procurement = g.Key.Procurement,
                .ItemCode = g.Key.ItemCode,
                .ItemName = g.First().ItemName,
                .Unit = g.First().Unit,
                .TotalQuantity = g.Sum(Function(l) l.CumulativeQuantity),
                .UsageCount = g.Count()
            }).
            OrderBy(Function(s) products.IndexOf(s.ProductCode)).
            ThenBy(Function(s) ProcurementSortOrder(s.Procurement)).
            ThenBy(Function(s) s.ItemCode, StringComparer.Ordinal).
            ToList()
    End Function

    ''' <summary>手配区分の並び順（購入が先、内製が後）</summary>
    Private Shared Function ProcurementSortOrder(procurement As String) As Integer
        Select Case procurement
            Case "購入" : Return 0
            Case "内製" : Return 1
            Case Else : Return 2
        End Select
    End Function

    ''' <summary>
    ''' 展開処理の間で共有するデータをまとめたクラスです。
    ''' </summary>
    Private Class ExpandContext

        ''' <summary>親品番 → 子の行の一覧（CSVの並び順）</summary>
        Public ReadOnly ChildrenOf As New Dictionary(Of String, List(Of BomRow))(StringComparer.Ordinal)

        ''' <summary>親品番がCSVに初めて出てきた順の一覧</summary>
        Public ReadOnly ParentOrder As New List(Of String)

        ''' <summary>子品番として出てくる品番の集合</summary>
        Public ReadOnly ChildCodes As New HashSet(Of String)(StringComparer.Ordinal)

        ''' <summary>製品からの展開でたどった品番</summary>
        Public ReadOnly Visited As New HashSet(Of String)(StringComparer.Ordinal)

        ''' <summary>記録済みのループ（同じループを何度も記録しないため）</summary>
        Private ReadOnly _reportedCycles As New HashSet(Of String)(StringComparer.Ordinal)

        Private ReadOnly _errors As List(Of BomError)

        Public Sub New(rows As List(Of BomRow), errors As List(Of BomError))
            _errors = errors
            For Each row In rows
                Dim children As List(Of BomRow) = Nothing
                If Not ChildrenOf.TryGetValue(row.ParentCode, children) Then
                    children = New List(Of BomRow)
                    ChildrenOf.Add(row.ParentCode, children)
                    ParentOrder.Add(row.ParentCode)
                End If
                children.Add(row)
                ChildCodes.Add(row.ChildCode)
            Next
        End Sub

        ''' <summary>
        ''' 循環参照をエラーに記録します。同じループはどこから見つけても1回だけ記録します。
        ''' </summary>
        ''' <param name="cycle">ループする品番の並び（例：B, C → B→C→B のループ）</param>
        ''' <param name="closingRow">ループを閉じている行（最後の品番から先頭の品番へ戻る行）</param>
        Public Sub ReportCycle(cycle As List(Of String), closingRow As BomRow)
            ' ループの開始位置が違っても同じループと判断できるよう、最小の品番から始まる形に並べ替えたものをキーにする
            Dim minIndex = 0
            For i = 1 To cycle.Count - 1
                If String.CompareOrdinal(cycle(i), cycle(minIndex)) < 0 Then minIndex = i
            Next
            Dim key = String.Join(vbTab, cycle.Skip(minIndex).Concat(cycle.Take(minIndex)))
            If Not _reportedCycles.Add(key) Then Return

            Dim route = String.Join(" → ", cycle.Append(cycle(0)))
            _errors.Add(New BomError(closingRow.LineNumber, BomErrorKind.CircularReference,
                $"循環参照を検出しました：{route}。この経路の展開を中止しました。",
                $"親品番={closingRow.ParentCode}, 子品番={closingRow.ChildCode}"))
        End Sub

    End Class

End Class
