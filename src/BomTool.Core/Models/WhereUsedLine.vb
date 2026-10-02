''' <summary>
''' 部品から製品までを逆にたどった部品表の1行です（Excelの「逆展開」シートの1行に対応します）。
''' 部品ごとに、レベル0の部品自身の行から始まり、親 → その親 … → 製品 の順に並びます。
''' </summary>
Public Class WhereUsedLine

    ''' <summary>たどり始めた部品の品番</summary>
    Public Property PartCode As String = ""

    ''' <summary>たどり始めた部品の品名</summary>
    Public Property PartName As String = ""

    ''' <summary>レベル（0が部品自身、その親が1、さらにその親が2…）</summary>
    Public Property Level As Integer

    ''' <summary>品番</summary>
    Public Property ItemCode As String = ""

    ''' <summary>品名（製品はCSVに品名がないため空です）</summary>
    Public Property ItemName As String = ""

    ''' <summary>員数（この品番1個に使う、1つ下のレベルの品番の数）。レベル0の部品自身は Nothing です。</summary>
    Public Property Quantity As Decimal?

    ''' <summary>累計員数（この品番1個に使う部品の数 ＝ 部品からここまでの員数をすべて掛け合わせた値）</summary>
    Public Property CumulativeQuantity As Decimal

    ''' <summary>部品の単位（累計員数の単位）</summary>
    Public Property Unit As String = ""

    ''' <summary>この品番の手配区分（製品は空です）</summary>
    Public Property Procurement As String = ""

    ''' <summary>この行が製品（最上位品番）なら True</summary>
    Public Property IsProduct As Boolean

End Class
