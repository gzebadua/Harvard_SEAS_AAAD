<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UpgradingForms
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(UpgradingForms))
        Me.rtb = New System.Windows.Forms.RichTextBox()
        Me.inputFolder = New System.Windows.Forms.FolderBrowserDialog()
        Me.idFinder = New System.Windows.Forms.OpenFileDialog()
        Me.outputFolder = New System.Windows.Forms.FolderBrowserDialog()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.templateFinder = New System.Windows.Forms.OpenFileDialog()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.formatOneLinedXML = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'rtb
        '
        Me.rtb.Location = New System.Drawing.Point(12, 12)
        Me.rtb.Name = "rtb"
        Me.rtb.Size = New System.Drawing.Size(485, 214)
        Me.rtb.TabIndex = 0
        Me.rtb.Text = ""
        '
        'idFinder
        '
        Me.idFinder.Title = "Please select the Forms ID's file"
        '
        'btnExit
        '
        Me.btnExit.Enabled = False
        Me.btnExit.Location = New System.Drawing.Point(12, 232)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(75, 23)
        Me.btnExit.TabIndex = 1
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'templateFinder
        '
        Me.templateFinder.Title = "Please select the new version of the AAAD template"
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(382, 232)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(115, 23)
        Me.btnUpdate.TabIndex = 2
        Me.btnUpdate.Text = "Re-Run update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'formatOneLinedXML
        '
        Me.formatOneLinedXML.Location = New System.Drawing.Point(141, 232)
        Me.formatOneLinedXML.Name = "formatOneLinedXML"
        Me.formatOneLinedXML.Size = New System.Drawing.Size(201, 23)
        Me.formatOneLinedXML.TabIndex = 3
        Me.formatOneLinedXML.Text = "Format Comparable-One-Lined XML"
        Me.formatOneLinedXML.UseVisualStyleBackColor = True
        '
        'UpgradingForms
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(509, 262)
        Me.Controls.Add(Me.formatOneLinedXML)
        Me.Controls.Add(Me.btnUpdate)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.rtb)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "UpgradingForms"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "AAAD XML Forms Updater"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents rtb As System.Windows.Forms.RichTextBox
    Friend WithEvents inputFolder As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents idFinder As System.Windows.Forms.OpenFileDialog
    Friend WithEvents outputFolder As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents templateFinder As System.Windows.Forms.OpenFileDialog
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents formatOneLinedXML As System.Windows.Forms.Button

End Class
