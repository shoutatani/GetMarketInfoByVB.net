<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class userForm
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.doCrawling = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.pgBar = New System.Windows.Forms.ProgressBar()
        Me.btnEnd = New System.Windows.Forms.Button()
        Me.cboDriver = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'doCrawling
        '
        Me.doCrawling.Location = New System.Drawing.Point(12, 67)
        Me.doCrawling.Name = "doCrawling"
        Me.doCrawling.Size = New System.Drawing.Size(486, 43)
        Me.doCrawling.TabIndex = 0
        Me.doCrawling.Text = "巡回開始"
        Me.doCrawling.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(306, 12)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "巡回先 : http://info.finance.yahoo.co.jp/fx/marketcalendar/"
        '
        'pgBar
        '
        Me.pgBar.Location = New System.Drawing.Point(12, 32)
        Me.pgBar.Name = "pgBar"
        Me.pgBar.Size = New System.Drawing.Size(552, 23)
        Me.pgBar.TabIndex = 2
        '
        'btnEnd
        '
        Me.btnEnd.Location = New System.Drawing.Point(504, 67)
        Me.btnEnd.Name = "btnEnd"
        Me.btnEnd.Size = New System.Drawing.Size(60, 43)
        Me.btnEnd.TabIndex = 3
        Me.btnEnd.Text = "終了"
        Me.btnEnd.UseVisualStyleBackColor = True
        '
        'cboDriver
        '
        Me.cboDriver.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDriver.FormattingEnabled = True
        Me.cboDriver.Location = New System.Drawing.Point(336, 6)
        Me.cboDriver.Name = "cboDriver"
        Me.cboDriver.Size = New System.Drawing.Size(228, 20)
        Me.cboDriver.TabIndex = 4
        '
        'userForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(576, 122)
        Me.Controls.Add(Me.cboDriver)
        Me.Controls.Add(Me.btnEnd)
        Me.Controls.Add(Me.pgBar)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.doCrawling)
        Me.MaximizeBox = False
        Me.Name = "userForm"
        Me.ShowIcon = False
        Me.Text = "経済指標巡回プログラム"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents doCrawling As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents pgBar As System.Windows.Forms.ProgressBar
    Friend WithEvents btnEnd As System.Windows.Forms.Button
    Friend WithEvents cboDriver As System.Windows.Forms.ComboBox
End Class
