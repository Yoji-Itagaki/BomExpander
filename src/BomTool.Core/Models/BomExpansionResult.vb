''' <summary>
''' BomTreeExpander の展開結果です。
''' </summary>
Public Class BomExpansionResult

    ''' <summary>最上位の品番（製品）の一覧。CSVに出てきた順です。</summary>
    Public ReadOnly Property Products As New List(Of String)

    ''' <summary>多段階展開した行（製品ごとに、レベル0の製品行から順に並びます）</summary>
    Public ReadOnly Property Lines As New List(Of ExpandedLine)

    ''' <summary>製品ごと・品番ごとの集計</summary>
    Public Property Summary As New List(Of SummaryLine)

    ''' <summary>部品ごとに製品までを逆にたどった結果（逆展開）</summary>
    Public Property WhereUsed As New List(Of WhereUsedLine)

    ''' <summary>展開時に見つかったエラー（循環参照など）</summary>
    Public ReadOnly Property Errors As New List(Of BomError)

    ''' <summary>部品行の数（レベル1以上の行の数。画面表示用）</summary>
    Public ReadOnly Property PartLineCount As Integer
        Get
            Return Lines.Where(Function(l) l.Level > 0).Count()
        End Get
    End Property

End Class
