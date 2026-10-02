<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MainForm
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        components = New System.ComponentModel.Container()
        lblCsvCaption = New Label()
        txtCsvPath = New TextBox()
        btnSelectCsv = New Button()
        btnExport = New Button()
        grpResult = New GroupBox()
        lblProductCountCaption = New Label()
        lblProductCountValue = New Label()
        lblPartLineCountCaption = New Label()
        lblPartLineCountValue = New Label()
        lblErrorCountCaption = New Label()
        lblErrorCountValue = New Label()
        lblEncodingCaption = New Label()
        lblEncodingValue = New Label()
        lblOutputCaption = New Label()
        lblOutputValue = New Label()
        toolTip = New ToolTip(components)
        grpResult.SuspendLayout()
        SuspendLayout()
        '
        ' lblCsvCaption
        '
        lblCsvCaption.AutoSize = True
        lblCsvCaption.Location = New Point(16, 16)
        lblCsvCaption.Name = "lblCsvCaption"
        lblCsvCaption.Size = New Size(104, 18)
        lblCsvCaption.TabIndex = 0
        lblCsvCaption.Text = "入力CSVファイル"
        '
        ' txtCsvPath
        '
        txtCsvPath.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        txtCsvPath.BackColor = SystemColors.Window
        txtCsvPath.Location = New Point(16, 40)
        txtCsvPath.Name = "txtCsvPath"
        txtCsvPath.PlaceholderText = "「CSVを選択」ボタンでファイルを選んでください"
        txtCsvPath.ReadOnly = True
        txtCsvPath.Size = New Size(480, 25)
        txtCsvPath.TabIndex = 1
        txtCsvPath.TabStop = False
        '
        ' btnSelectCsv
        '
        btnSelectCsv.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        btnSelectCsv.Location = New Point(504, 38)
        btnSelectCsv.Name = "btnSelectCsv"
        btnSelectCsv.Size = New Size(120, 30)
        btnSelectCsv.TabIndex = 2
        btnSelectCsv.Text = "CSVを選択..."
        btnSelectCsv.UseVisualStyleBackColor = True
        '
        ' btnExport
        '
        btnExport.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        btnExport.Enabled = False
        btnExport.Font = New Font("Yu Gothic UI", 11.0F, FontStyle.Bold)
        btnExport.Location = New Point(16, 82)
        btnExport.Name = "btnExport"
        btnExport.Size = New Size(608, 42)
        btnExport.TabIndex = 3
        btnExport.Text = "展開してExcelに出力"
        btnExport.UseVisualStyleBackColor = True
        '
        ' grpResult
        '
        grpResult.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        grpResult.Controls.Add(lblProductCountCaption)
        grpResult.Controls.Add(lblProductCountValue)
        grpResult.Controls.Add(lblPartLineCountCaption)
        grpResult.Controls.Add(lblPartLineCountValue)
        grpResult.Controls.Add(lblErrorCountCaption)
        grpResult.Controls.Add(lblErrorCountValue)
        grpResult.Controls.Add(lblEncodingCaption)
        grpResult.Controls.Add(lblEncodingValue)
        grpResult.Controls.Add(lblOutputCaption)
        grpResult.Controls.Add(lblOutputValue)
        grpResult.Location = New Point(16, 140)
        grpResult.Name = "grpResult"
        grpResult.Size = New Size(608, 176)
        grpResult.TabIndex = 4
        grpResult.TabStop = False
        grpResult.Text = "処理結果"
        '
        ' lblProductCountCaption
        '
        lblProductCountCaption.AutoSize = True
        lblProductCountCaption.Location = New Point(16, 30)
        lblProductCountCaption.Name = "lblProductCountCaption"
        lblProductCountCaption.Text = "製品数"
        '
        ' lblProductCountValue
        '
        lblProductCountValue.AutoSize = True
        lblProductCountValue.Font = New Font("Yu Gothic UI", 10.0F, FontStyle.Bold)
        lblProductCountValue.Location = New Point(140, 30)
        lblProductCountValue.Name = "lblProductCountValue"
        lblProductCountValue.Text = "－"
        '
        ' lblPartLineCountCaption
        '
        lblPartLineCountCaption.AutoSize = True
        lblPartLineCountCaption.Location = New Point(16, 58)
        lblPartLineCountCaption.Name = "lblPartLineCountCaption"
        lblPartLineCountCaption.Text = "部品行数"
        '
        ' lblPartLineCountValue
        '
        lblPartLineCountValue.AutoSize = True
        lblPartLineCountValue.Font = New Font("Yu Gothic UI", 10.0F, FontStyle.Bold)
        lblPartLineCountValue.Location = New Point(140, 58)
        lblPartLineCountValue.Name = "lblPartLineCountValue"
        lblPartLineCountValue.Text = "－"
        '
        ' lblErrorCountCaption
        '
        lblErrorCountCaption.AutoSize = True
        lblErrorCountCaption.Location = New Point(16, 86)
        lblErrorCountCaption.Name = "lblErrorCountCaption"
        lblErrorCountCaption.Text = "エラー件数"
        '
        ' lblErrorCountValue
        '
        lblErrorCountValue.AutoSize = True
        lblErrorCountValue.Font = New Font("Yu Gothic UI", 10.0F, FontStyle.Bold)
        lblErrorCountValue.Location = New Point(140, 86)
        lblErrorCountValue.Name = "lblErrorCountValue"
        lblErrorCountValue.Text = "－"
        '
        ' lblEncodingCaption
        '
        lblEncodingCaption.AutoSize = True
        lblEncodingCaption.Location = New Point(16, 114)
        lblEncodingCaption.Name = "lblEncodingCaption"
        lblEncodingCaption.Text = "CSVの文字コード"
        '
        ' lblEncodingValue
        '
        lblEncodingValue.AutoSize = True
        lblEncodingValue.Location = New Point(140, 114)
        lblEncodingValue.Name = "lblEncodingValue"
        lblEncodingValue.Text = "－"
        '
        ' lblOutputCaption
        '
        lblOutputCaption.AutoSize = True
        lblOutputCaption.Location = New Point(16, 142)
        lblOutputCaption.Name = "lblOutputCaption"
        lblOutputCaption.Text = "出力ファイル"
        '
        ' lblOutputValue
        '
        lblOutputValue.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        lblOutputValue.AutoEllipsis = True
        lblOutputValue.Location = New Point(140, 142)
        lblOutputValue.Name = "lblOutputValue"
        lblOutputValue.Size = New Size(452, 20)
        lblOutputValue.Text = "－"
        '
        ' MainForm
        '
        AutoScaleDimensions = New SizeF(96.0F, 96.0F)
        AutoScaleMode = AutoScaleMode.Dpi
        ClientSize = New Size(640, 332)
        Controls.Add(lblCsvCaption)
        Controls.Add(txtCsvPath)
        Controls.Add(btnSelectCsv)
        Controls.Add(btnExport)
        Controls.Add(grpResult)
        Font = New Font("Yu Gothic UI", 10.0F)
        FormBorderStyle = FormBorderStyle.FixedSingle
        MaximizeBox = False
        Name = "MainForm"
        StartPosition = FormStartPosition.CenterScreen
        Text = "多段階BOM展開ツール"
        grpResult.ResumeLayout(False)
        grpResult.PerformLayout()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents lblCsvCaption As Label
    Friend WithEvents txtCsvPath As TextBox
    Friend WithEvents btnSelectCsv As Button
    Friend WithEvents btnExport As Button
    Friend WithEvents grpResult As GroupBox
    Friend WithEvents lblProductCountCaption As Label
    Friend WithEvents lblProductCountValue As Label
    Friend WithEvents lblPartLineCountCaption As Label
    Friend WithEvents lblPartLineCountValue As Label
    Friend WithEvents lblErrorCountCaption As Label
    Friend WithEvents lblErrorCountValue As Label
    Friend WithEvents lblEncodingCaption As Label
    Friend WithEvents lblEncodingValue As Label
    Friend WithEvents lblOutputCaption As Label
    Friend WithEvents lblOutputValue As Label
    Friend WithEvents toolTip As ToolTip

End Class
