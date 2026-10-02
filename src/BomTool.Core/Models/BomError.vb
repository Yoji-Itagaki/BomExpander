''' <summary>
''' エラーの種類です。Excelの「エラー」シートには KindText の日本語名で出力します。
''' </summary>
Public Enum BomErrorKind
    ''' <summary>列数が足りないなど、CSVの形式の問題</summary>
    FormatError
    ''' <summary>必須列が空</summary>
    MissingRequired
    ''' <summary>員数が数値でない、または0以下</summary>
    InvalidQuantity
    ''' <summary>手配区分が「購入」「内製」以外</summary>
    InvalidProcurement
    ''' <summary>同じ親子の組み合わせが重複している</summary>
    DuplicateRelation
    ''' <summary>親子関係がループしている</summary>
    CircularReference
    ''' <summary>最上位の品番（製品）が1つも見つからない</summary>
    NoTopLevelItem
End Enum

''' <summary>
''' 読み込み・展開の途中で見つかった問題1件を表すクラスです。
''' </summary>
Public Class BomError

    ''' <summary>エラーを作成します。</summary>
    ''' <param name="lineNumber">CSVの行番号。特定の行に結び付かない場合は0</param>
    ''' <param name="kind">エラーの種類</param>
    ''' <param name="message">エラーの内容（画面・Excelに表示する文章）</param>
    ''' <param name="rawText">問題のあったCSVの元の行など、参考情報</param>
    Public Sub New(lineNumber As Integer, kind As BomErrorKind, message As String, Optional rawText As String = "")
        Me.LineNumber = lineNumber
        Me.Kind = kind
        Me.Message = message
        Me.RawText = If(rawText, "")
    End Sub

    ''' <summary>CSVの行番号（0は「特定の行なし」）</summary>
    Public ReadOnly Property LineNumber As Integer

    ''' <summary>エラーの種類</summary>
    Public ReadOnly Property Kind As BomErrorKind

    ''' <summary>エラーの内容</summary>
    Public ReadOnly Property Message As String

    ''' <summary>元データ（CSVの行の文字列など）</summary>
    Public ReadOnly Property RawText As String

    ''' <summary>エラーの種類の日本語名を返します。</summary>
    Public ReadOnly Property KindText As String
        Get
            Select Case Kind
                Case BomErrorKind.FormatError : Return "CSV形式"
                Case BomErrorKind.MissingRequired : Return "必須列が空"
                Case BomErrorKind.InvalidQuantity : Return "員数が不正"
                Case BomErrorKind.InvalidProcurement : Return "手配区分が不正"
                Case BomErrorKind.DuplicateRelation : Return "親子の重複"
                Case BomErrorKind.CircularReference : Return "循環参照"
                Case BomErrorKind.NoTopLevelItem : Return "製品なし"
                Case Else : Return Kind.ToString()
            End Select
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return $"{LineNumber}行目 [{KindText}] {Message}"
    End Function

End Class
