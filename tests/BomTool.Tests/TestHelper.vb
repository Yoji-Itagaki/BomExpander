Imports BomTool.Core

''' <summary>
''' テストで使う共通の処理です。
''' </summary>
Friend Module TestHelper

    ''' <summary>テスト用CSVのヘッダー行</summary>
    Public Const Header As String = "親品番,子品番,品名,員数,単位,手配区分"

    ''' <summary>
    ''' ヘッダーを付けたCSV文字列を作ります。
    ''' </summary>
    ''' <param name="dataLines">2行目以降のデータ行</param>
    Public Function Csv(ParamArray dataLines As String()) As String
        Return Header & vbCrLf & String.Join(vbCrLf, dataLines) & vbCrLf
    End Function

    ''' <summary>
    ''' CSV文字列を読み込み、そのまま展開します。
    ''' </summary>
    Public Function ReadAndExpand(csvText As String) As (Read As CsvReadResult, Expansion As BomExpansionResult)
        Dim readResult = New CsvBomReader().ReadText(csvText)
        Dim expansion = New BomTreeExpander().Expand(readResult.Rows)
        Return (readResult, expansion)
    End Function

    ''' <summary>
    ''' 展開結果から、指定した製品・品番の行を探します（同じ品番が複数あれば最初の行）。
    ''' </summary>
    Public Function FindLine(expansion As BomExpansionResult, productCode As String, itemCode As String) As ExpandedLine
        Return expansion.Lines.First(Function(l) l.ProductCode = productCode AndAlso l.ItemCode = itemCode)
    End Function

End Module
