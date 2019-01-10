<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgCatParam
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
        Me.components = New System.ComponentModel.Container()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.LbxCategories = New System.Windows.Forms.ListBox()
        Me.ChkbxParameters = New System.Windows.Forms.CheckedListBox()
        Me.cmsParams = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.mnuAllCheck = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuAllUnCheck = New System.Windows.Forms.ToolStripMenuItem()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.RbnInstance = New System.Windows.Forms.RadioButton()
        Me.RbnType = New System.Windows.Forms.RadioButton()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.BtnLoad = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.cmsParams.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(633, 489)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(243, 40)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Location = New System.Drawing.Point(5, 4)
        Me.OK_Button.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(111, 32)
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "OK"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(126, 4)
        Me.Cancel_Button.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(112, 32)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.Text = "Cancel"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(20, 18)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.LbxCategories)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.ChkbxParameters)
        Me.SplitContainer1.Size = New System.Drawing.Size(857, 417)
        Me.SplitContainer1.SplitterDistance = 281
        Me.SplitContainer1.SplitterWidth = 7
        Me.SplitContainer1.TabIndex = 1
        '
        'LbxCategories
        '
        Me.LbxCategories.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LbxCategories.FormattingEnabled = True
        Me.LbxCategories.ItemHeight = 18
        Me.LbxCategories.Location = New System.Drawing.Point(0, 0)
        Me.LbxCategories.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.LbxCategories.Name = "LbxCategories"
        Me.LbxCategories.Size = New System.Drawing.Size(281, 417)
        Me.LbxCategories.Sorted = True
        Me.LbxCategories.TabIndex = 0
        '
        'ChkbxParameters
        '
        Me.ChkbxParameters.CheckOnClick = True
        Me.ChkbxParameters.ContextMenuStrip = Me.cmsParams
        Me.ChkbxParameters.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ChkbxParameters.FormattingEnabled = True
        Me.ChkbxParameters.Location = New System.Drawing.Point(0, 0)
        Me.ChkbxParameters.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.ChkbxParameters.Name = "ChkbxParameters"
        Me.ChkbxParameters.Size = New System.Drawing.Size(569, 417)
        Me.ChkbxParameters.Sorted = True
        Me.ChkbxParameters.TabIndex = 0
        '
        'cmsParams
        '
        Me.cmsParams.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.cmsParams.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuAllCheck, Me.mnuAllUnCheck})
        Me.cmsParams.Name = "cmsParams"
        Me.cmsParams.Size = New System.Drawing.Size(244, 64)
        '
        'mnuAllCheck
        '
        Me.mnuAllCheck.Name = "mnuAllCheck"
        Me.mnuAllCheck.Size = New System.Drawing.Size(243, 30)
        Me.mnuAllCheck.Text = "すべてチェックする"
        '
        'mnuAllUnCheck
        '
        Me.mnuAllUnCheck.Name = "mnuAllUnCheck"
        Me.mnuAllUnCheck.Size = New System.Drawing.Size(243, 30)
        Me.mnuAllUnCheck.Text = "すべてチェック解除する"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.RbnInstance)
        Me.GroupBox1.Controls.Add(Me.RbnType)
        Me.GroupBox1.Location = New System.Drawing.Point(20, 466)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.GroupBox1.Size = New System.Drawing.Size(358, 74)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        '
        'RbnInstance
        '
        Me.RbnInstance.AutoSize = True
        Me.RbnInstance.Location = New System.Drawing.Point(165, 27)
        Me.RbnInstance.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.RbnInstance.Name = "RbnInstance"
        Me.RbnInstance.Size = New System.Drawing.Size(96, 22)
        Me.RbnInstance.TabIndex = 1
        Me.RbnInstance.Text = "&Instance"
        Me.RbnInstance.UseVisualStyleBackColor = True
        '
        'RbnType
        '
        Me.RbnType.AutoSize = True
        Me.RbnType.Location = New System.Drawing.Point(10, 27)
        Me.RbnType.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.RbnType.Name = "RbnType"
        Me.RbnType.Size = New System.Drawing.Size(71, 22)
        Me.RbnType.TabIndex = 0
        Me.RbnType.Text = "&Type"
        Me.RbnType.UseVisualStyleBackColor = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(20, 444)
        Me.ProgressBar1.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(857, 20)
        Me.ProgressBar1.TabIndex = 3
        '
        'BtnLoad
        '
        Me.BtnLoad.Location = New System.Drawing.Point(435, 492)
        Me.BtnLoad.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.BtnLoad.Name = "BtnLoad"
        Me.BtnLoad.Size = New System.Drawing.Size(125, 34)
        Me.BtnLoad.TabIndex = 4
        Me.BtnLoad.Text = "設定読込"
        Me.BtnLoad.UseVisualStyleBackColor = True
        '
        'dlgCatParam
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(10.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(897, 546)
        Me.Controls.Add(Me.BtnLoad)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dlgCatParam"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Select Category and Parameters"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.cmsParams.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents LbxCategories As System.Windows.Forms.ListBox
    Friend WithEvents ChkbxParameters As System.Windows.Forms.CheckedListBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RbnInstance As System.Windows.Forms.RadioButton
    Friend WithEvents RbnType As System.Windows.Forms.RadioButton
    Friend WithEvents cmsParams As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents mnuAllCheck As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuAllUnCheck As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents BtnLoad As System.Windows.Forms.Button

End Class
