''' <summary>
''' CsvBomReader の読み込み結果です。
''' </summary>
Public Class CsvReadResult

    ''' <summary>チェックを通過した行（展開に使う行）</summary>
    Public ReadOnly Property Rows As New List(Of BomRow)

    ''' <summary>読み込み時に見つかったエラー</summary>
    Public ReadOnly Property Errors As New List(Of BomError)

    ''' <summary>判定した文字コードの名前（画面表示用）</summary>
    Public Property EncodingName As String = ""

    ''' <summary>ヘッダーと空行を除いたデータ行の数</summary>
    Public Property DataLineCount As Integer

End Class
