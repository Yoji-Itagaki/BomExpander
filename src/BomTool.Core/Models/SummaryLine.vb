''' <summary>
''' 製品ごとに同じ品番の累計員数を合算した1行です（Excelの「部品集計」シートの1行に対応します）。
''' </summary>
Public Class SummaryLine

    ''' <summary>製品（最上位品番）</summary>
    Public Property ProductCode As String = ""

    ''' <summary>手配区分</summary>
    Public Property Procurement As String = ""

    ''' <summary>品番</summary>
    Public Property ItemCode As String = ""

    ''' <summary>品名</summary>
    Public Property ItemName As String = ""

    ''' <summary>合計員数（製品1台あたり。部品表の中で何か所に出てきても合算した値）</summary>
    Public Property TotalQuantity As Decimal

    ''' <summary>単位</summary>
    Public Property Unit As String = ""

    ''' <summary>部品表の中でこの品番が出てくる箇所の数</summary>
    Public Property UsageCount As Integer

End Class
