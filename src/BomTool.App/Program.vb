Imports System.Threading

''' <summary>
''' アプリケーションの開始位置です。
''' 想定外の例外が起きてもツールが落ちずにメッセージを出すよう、ここで例外の受け口を設定します。
''' </summary>
Module Program

    <STAThread>
    Sub Main()
        Application.SetHighDpiMode(HighDpiMode.PerMonitorV2)
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)

        ' 画面の処理中に起きた例外は ThreadException で受け取り、アプリを終了させない
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException)
        AddHandler Application.ThreadException, AddressOf OnThreadException
        AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf OnUnhandledException

        Application.Run(New MainForm())
    End Sub

    ''' <summary>画面の処理中に捕まえられなかった例外をメッセージで表示します。</summary>
    Private Sub OnThreadException(sender As Object, e As ThreadExceptionEventArgs)
        ShowFatalError(e.Exception)
    End Sub

    ''' <summary>画面以外のスレッドで起きた例外をメッセージで表示します。</summary>
    Private Sub OnUnhandledException(sender As Object, e As UnhandledExceptionEventArgs)
        ShowFatalError(TryCast(e.ExceptionObject, Exception))
    End Sub

    Private Sub ShowFatalError(ex As Exception)
        Dim detail = If(ex Is Nothing, "（詳細不明）", ex.Message)
        MessageBox.Show($"予期しないエラーが発生しました。{vbCrLf}{vbCrLf}{detail}",
                        "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

End Module
