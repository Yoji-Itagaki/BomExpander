''' <summary>
''' 多段階展開の結果から、部品ごとに「どの組品・どの製品に使われているか」を逆にたどる（逆展開する）クラスです。
'''
''' ・部品表に出てくるすべての品番（製品を除く）について、品番の順に作ります。
''' ・部品自身（レベル0）から親 → その親 … と上へたどり、製品に着いたところで止めます。
''' ・累計員数 ＝ 部品からその品番までの員数をすべて掛け合わせた値（その品番1個に使う部品の数）です。
'''   製品の行の累計員数は、その経路での「製品1台あたりの部品の数」になります。
''' ・親子関係は展開結果（製品からたどれた関係）だけを使うため、エラーで読み飛ばした行や、
'''   循環参照で展開を止めた関係は含まれません。
''' </summary>
Public Class WhereUsedBuilder

    ''' <summary>
    ''' 逆展開の行を作ります。
    ''' </summary>
    ''' <param name="lines">BomTreeExpander の展開結果</param>
    ''' <param name="maxLines">作る行の上限（Excelの行数上限を超えないようにするため）</param>
    Public Shared Function Build(lines As IEnumerable(Of ExpandedLine), maxLines As Integer) As List(Of WhereUsedLine)
        Dim context As New BuildContext(lines.ToList(), maxLines)
        Dim result As New List(Of WhereUsedLine)

        For Each partCode In context.PartCodes
            Dim part = context.ItemInfo(partCode)
            Dim partRow As New WhereUsedLine With {
                .PartCode = partCode,
                .PartName = part.ItemName,
                .Level = 0,
                .ItemCode = partCode,
                .ItemName = part.ItemName,
                .Quantity = Nothing,
                .CumulativeQuantity = 1D,
                .Unit = part.Unit,
                .Procurement = part.Procurement
            }
            context.CountLine()
            result.Add(partRow)

            Dim path As New List(Of String) From {partCode}
            result.AddRange(BuildParents(partRow, partCode, 1, 1D, path, context))
        Next

        Return result
    End Function

    ''' <summary>
    ''' 品番の親を順番にたどり、製品まで着いた経路の行だけを返します。
    ''' </summary>
    ''' <param name="part">レベル0の部品の行</param>
    ''' <param name="childCode">親を探す品番</param>
    ''' <param name="level">親のレベル</param>
    ''' <param name="childCumulative">childCode 1個に使う部品の数</param>
    ''' <param name="path">部品から childCode までの経路</param>
    Private Shared Function BuildParents(part As WhereUsedLine, childCode As String, level As Integer,
                                         childCumulative As Decimal, path As List(Of String),
                                         context As BuildContext) As List(Of WhereUsedLine)
        Dim rows As New List(Of WhereUsedLine)
        Dim parents As List(Of (ParentCode As String, Quantity As Decimal)) = Nothing
        If Not context.ParentsOf.TryGetValue(childCode, parents) Then Return rows

        For Each parent In parents
            ' 経路の中にすでにある品番に戻る場合はループなので、そこから先はたどらない
            ' （循環参照は展開時にエラーとして記録済みです）
            If path.Contains(parent.ParentCode) Then Continue For

            Dim isProduct = context.Products.Contains(parent.ParentCode)
            Dim info As ExpandedLine = Nothing
            context.ItemInfo.TryGetValue(parent.ParentCode, info)
            Dim cumulative = childCumulative * parent.Quantity
            Dim row As New WhereUsedLine With {
                .PartCode = part.PartCode,
                .PartName = part.PartName,
                .Level = level,
                .ItemCode = parent.ParentCode,
                .ItemName = If(isProduct OrElse info Is Nothing, "", info.ItemName),
                .Quantity = parent.Quantity,
                .CumulativeQuantity = cumulative,
                .Unit = part.Unit,
                .Procurement = If(isProduct OrElse info Is Nothing, "", info.Procurement),
                .IsProduct = isProduct
            }

            If isProduct Then
                context.CountLine()
                rows.Add(row)
                Continue For
            End If

            path.Add(parent.ParentCode)
            Dim upper = BuildParents(part, parent.ParentCode, level + 1, cumulative, path, context)
            path.RemoveAt(path.Count - 1)

            ' 製品まで着かなかった経路（ループで行き止まりになったもの）は出力しない
            If upper.Count = 0 Then Continue For
            context.CountLine()
            rows.Add(row)
            rows.AddRange(upper)
        Next

        Return rows
    End Function

    ''' <summary>
    ''' 逆展開の間で共有するデータをまとめたクラスです。
    ''' </summary>
    Private Class BuildContext

        ''' <summary>製品（最上位品番）の集合</summary>
        Public ReadOnly Products As New HashSet(Of String)(StringComparer.Ordinal)

        ''' <summary>品番 → その品番が最初に出てきた展開行（品名・単位・手配区分を取るため）</summary>
        Public ReadOnly ItemInfo As New Dictionary(Of String, ExpandedLine)(StringComparer.Ordinal)

        ''' <summary>子品番 → 親品番と員数の一覧（展開結果に最初に出てきた順）</summary>
        Public ReadOnly ParentsOf As New Dictionary(Of String, List(Of (ParentCode As String, Quantity As Decimal)))(StringComparer.Ordinal)

        ''' <summary>逆展開する部品（製品以外の品番）を品番の順に並べたもの</summary>
        Public ReadOnly PartCodes As List(Of String)

        Private ReadOnly _maxLines As Integer
        Private _lineCount As Integer

        Public Sub New(lines As List(Of ExpandedLine), maxLines As Integer)
            _maxLines = maxLines
            For Each line In lines
                If line.Level = 0 Then
                    Products.Add(line.ItemCode)
                    Continue For
                End If

                If Not ItemInfo.ContainsKey(line.ItemCode) Then ItemInfo.Add(line.ItemCode, line)

                Dim parents As List(Of (ParentCode As String, Quantity As Decimal)) = Nothing
                If Not ParentsOf.TryGetValue(line.ItemCode, parents) Then
                    parents = New List(Of (ParentCode As String, Quantity As Decimal))
                    ParentsOf.Add(line.ItemCode, parents)
                End If
                ' 同じ組品が複数の製品の下に展開されると同じ親子が何度も出てくるため、1回だけ登録する
                If Not parents.Any(Function(p) p.ParentCode = line.ParentCode) Then
                    parents.Add((line.ParentCode, line.Quantity))
                End If
            Next

            PartCodes = ItemInfo.Keys.OrderBy(Function(code) code, StringComparer.Ordinal).ToList()
        End Sub

        ''' <summary>作った行を数え、上限を超えたら処理を中止します。</summary>
        Public Sub CountLine()
            _lineCount += 1
            If _lineCount > _maxLines Then
                Throw New InvalidOperationException(
                    $"逆展開の結果が{_maxLines:#,0}行を超えたため処理を中止しました。部品表の構成を確認してください。")
            End If
        End Sub

    End Class

End Class
