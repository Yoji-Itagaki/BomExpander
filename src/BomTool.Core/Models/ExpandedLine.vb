''' <summary>
''' 多段階展開した部品表の1行です（Excelの「BOM展開」シートの1行に対応します）。
''' </summary>
Public Class ExpandedLine

    ''' <summary>この行が属する製品（最上位品番）</summary>
    Public Property ProductCode As String = ""

    ''' <summary>レベル（0が製品、その下が1、2…）</summary>
    Public Property Level As Integer

    ''' <summary>直上の親の品番（レベル0の製品は空です）</summary>
    Public Property ParentCode As String = ""

    ''' <summary>品番</summary>
    Public Property ItemCode As String = ""

    ''' <summary>品名（レベル0の製品はCSVに品名がないため空です）</summary>
    Public Property ItemName As String = ""

    ''' <summary>員数（直上の親1個あたりの数）</summary>
    Public Property Quantity As Decimal

    ''' <summary>累計員数（製品1台あたりの数 ＝ 上位の員数をすべて掛け合わせた値）</summary>
    Public Property CumulativeQuantity As Decimal

    ''' <summary>単位</summary>
    Public Property Unit As String = ""

    ''' <summary>手配区分</summary>
    Public Property Procurement As String = ""

    ''' <summary>元になったCSVの行番号（レベル0の製品は0）</summary>
    Public Property SourceLineNumber As Integer

End Class
