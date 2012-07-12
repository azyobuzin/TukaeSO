<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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

    'Windows フォーム デザイナで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
    'Windows フォーム デザイナを使用して変更できます。  
    'コード エディタを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.SaveButton = New System.Windows.Forms.Button
        Me.OpenButton = New System.Windows.Forms.Button
        Me.HashtableEditGrid1 = New TukaeSO.pepetaroControl.HashtableEditGrid
        Me.SuspendLayout()
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'SaveButton
        '
        Me.SaveButton.Location = New System.Drawing.Point(93, 329)
        Me.SaveButton.Name = "SaveButton"
        Me.SaveButton.Size = New System.Drawing.Size(75, 23)
        Me.SaveButton.TabIndex = 0
        Me.SaveButton.Text = "保存"
        Me.SaveButton.UseVisualStyleBackColor = True
        '
        'OpenButton
        '
        Me.OpenButton.Location = New System.Drawing.Point(12, 329)
        Me.OpenButton.Name = "OpenButton"
        Me.OpenButton.Size = New System.Drawing.Size(75, 23)
        Me.OpenButton.TabIndex = 1
        Me.OpenButton.Text = "開く"
        Me.OpenButton.UseVisualStyleBackColor = True
        '
        'HashtableEditGrid1
        '
        Me.HashtableEditGrid1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.HashtableEditGrid1.Hashtable = CType(resources.GetObject("HashtableEditGrid1.Hashtable"), System.Collections.Hashtable)
        Me.HashtableEditGrid1.Location = New System.Drawing.Point(12, 12)
        Me.HashtableEditGrid1.Name = "HashtableEditGrid1"
        Me.HashtableEditGrid1.SelectedKey = Nothing
        Me.HashtableEditGrid1.SelectedKeyIndex = -1
        Me.HashtableEditGrid1.Size = New System.Drawing.Size(422, 311)
        Me.HashtableEditGrid1.TabIndex = 2
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(446, 364)
        Me.Controls.Add(Me.HashtableEditGrid1)
        Me.Controls.Add(Me.OpenButton)
        Me.Controls.Add(Me.SaveButton)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents SaveButton As System.Windows.Forms.Button
    Friend WithEvents OpenButton As System.Windows.Forms.Button
    Friend WithEvents HashtableEditGrid1 As TukaeSO.pepetaroControl.HashtableEditGrid

End Class
