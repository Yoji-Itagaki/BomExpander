''' <summary>
''' CSVの1行（「親品番の下に子品番が員数分付く」という親子関係1件）を表すクラスです。
''' CsvBomReader がチェックを通過した行だけをこのクラスにして返します。
''' </summary>
Public Class BomRow

    ''' <summary>CSVファイル上の行番号（ヘッダーを1行目として数えます）。エラー表示に使います。</summary>
    Public Property LineNumber As Integer

    ''' <summary>親品番</summary>
    Public Property ParentCode As String = ""

    ''' <summary>子品番</summary>
    Public Property ChildCode As String = ""

    ''' <summary>子品番の品名</summary>
    Public Property ItemName As String = ""

    ''' <summary>員数（親1個あたりに使う子の数）。0より大きい値です。</summary>
    Public Property Quantity As Decimal

    ''' <summary>単位（個、本、m など）</summary>
    Public Property Unit As String = ""

    ''' <summary>手配区分（「購入」または「内製」）</summary>
    Public Property Procurement As String = ""

End Class
