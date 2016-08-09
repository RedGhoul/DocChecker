<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.Browse = New System.Windows.Forms.Button()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Process = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBoxResults = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.DateTimePickerEnd = New System.Windows.Forms.DateTimePicker()
        Me.DateTimePickerStart = New System.Windows.Forms.DateTimePicker()
        Me.Testlogsetup = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog2 = New System.Windows.Forms.FolderBrowserDialog()
        Me.NewFilesetup = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog3 = New System.Windows.Forms.FolderBrowserDialog()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.NewFilesetupText = New System.Windows.Forms.TextBox()
        Me.TestlogsetupText = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.CheckOpenOnEnd = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(7, 74)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(499, 28)
        Me.TextBox1.TabIndex = 0
        '
        'FolderBrowserDialog1
        '
        Me.FolderBrowserDialog1.SelectedPath = "\\cdiise1b\schedadfdocs$\Agreements\DGEN\"
        '
        'Browse
        '
        Me.Browse.Location = New System.Drawing.Point(514, 74)
        Me.Browse.Margin = New System.Windows.Forms.Padding(1)
        Me.Browse.Name = "Browse"
        Me.Browse.Size = New System.Drawing.Size(92, 28)
        Me.Browse.TabIndex = 1
        Me.Browse.Text = "Browse"
        Me.Browse.UseVisualStyleBackColor = True
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(11, 339)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(497, 30)
        Me.ListBox1.TabIndex = 2
        '
        'Process
        '
        Me.Process.Location = New System.Drawing.Point(514, 409)
        Me.Process.Name = "Process"
        Me.Process.Size = New System.Drawing.Size(90, 36)
        Me.Process.TabIndex = 3
        Me.Process.Text = "Process"
        Me.Process.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(5, 47)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(71, 24)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Source"
        '
        'TextBoxResults
        '
        Me.TextBoxResults.Enabled = False
        Me.TextBoxResults.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxResults.ForeColor = System.Drawing.Color.Red
        Me.TextBoxResults.Location = New System.Drawing.Point(11, 409)
        Me.TextBoxResults.Multiline = True
        Me.TextBoxResults.Name = "TextBoxResults"
        Me.TextBoxResults.Size = New System.Drawing.Size(497, 36)
        Me.TextBoxResults.TabIndex = 5
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(7, 312)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(262, 24)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Current File Being Processed "
        '
        'DateTimePickerEnd
        '
        Me.DateTimePickerEnd.Location = New System.Drawing.Point(322, 150)
        Me.DateTimePickerEnd.Name = "DateTimePickerEnd"
        Me.DateTimePickerEnd.Size = New System.Drawing.Size(225, 20)
        Me.DateTimePickerEnd.TabIndex = 7
        '
        'DateTimePickerStart
        '
        Me.DateTimePickerStart.Location = New System.Drawing.Point(7, 150)
        Me.DateTimePickerStart.Name = "DateTimePickerStart"
        Me.DateTimePickerStart.Size = New System.Drawing.Size(240, 20)
        Me.DateTimePickerStart.TabIndex = 8
        '
        'Testlogsetup
        '
        Me.Testlogsetup.Location = New System.Drawing.Point(514, 268)
        Me.Testlogsetup.Margin = New System.Windows.Forms.Padding(2)
        Me.Testlogsetup.Name = "Testlogsetup"
        Me.Testlogsetup.Size = New System.Drawing.Size(92, 30)
        Me.Testlogsetup.TabIndex = 9
        Me.Testlogsetup.Text = "Browser"
        Me.Testlogsetup.UseVisualStyleBackColor = True
        '
        'NewFilesetup
        '
        Me.NewFilesetup.Location = New System.Drawing.Point(514, 210)
        Me.NewFilesetup.Name = "NewFilesetup"
        Me.NewFilesetup.Size = New System.Drawing.Size(92, 29)
        Me.NewFilesetup.TabIndex = 10
        Me.NewFilesetup.Text = "Browser"
        Me.NewFilesetup.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(6, 108)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(178, 24)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Set your Dates Here"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(319, 132)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(70, 15)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "End Date:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(4, 132)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(75, 15)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "Start Date:"
        '
        'NewFilesetupText
        '
        Me.NewFilesetupText.Location = New System.Drawing.Point(10, 211)
        Me.NewFilesetupText.Multiline = True
        Me.NewFilesetupText.Name = "NewFilesetupText"
        Me.NewFilesetupText.Size = New System.Drawing.Size(499, 28)
        Me.NewFilesetupText.TabIndex = 14
        '
        'TestlogsetupText
        '
        Me.TestlogsetupText.Location = New System.Drawing.Point(11, 270)
        Me.TestlogsetupText.Multiline = True
        Me.TestlogsetupText.Name = "TestlogsetupText"
        Me.TestlogsetupText.Size = New System.Drawing.Size(499, 28)
        Me.TestlogsetupText.TabIndex = 15
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(6, 195)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(137, 13)
        Me.Label6.TabIndex = 16
        Me.Label6.Text = "Output File Destination"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(8, 254)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(120, 13)
        Me.Label7.TabIndex = 17
        Me.Label7.Text = "Log File Destination"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(7, 382)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(235, 24)
        Me.Label8.TabIndex = 18
        Me.Label8.Text = "Current Status of Operation"
        '
        'CheckOpenOnEnd
        '
        Me.CheckOpenOnEnd.AutoSize = True
        Me.CheckOpenOnEnd.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.CheckOpenOnEnd.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckOpenOnEnd.Location = New System.Drawing.Point(5, 12)
        Me.CheckOpenOnEnd.Name = "CheckOpenOnEnd"
        Me.CheckOpenOnEnd.Size = New System.Drawing.Size(292, 29)
        Me.CheckOpenOnEnd.TabIndex = 19
        Me.CheckOpenOnEnd.Text = "Open Log when Finished"
        Me.CheckOpenOnEnd.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(617, 468)
        Me.Controls.Add(Me.CheckOpenOnEnd)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.TestlogsetupText)
        Me.Controls.Add(Me.NewFilesetupText)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.NewFilesetup)
        Me.Controls.Add(Me.Testlogsetup)
        Me.Controls.Add(Me.DateTimePickerStart)
        Me.Controls.Add(Me.DateTimePickerEnd)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBoxResults)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Process)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.Browse)
        Me.Controls.Add(Me.TextBox1)
        Me.Name = "Form1"
        Me.Text = "Word Doc Checker"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Browse As System.Windows.Forms.Button
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents Process As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBoxResults As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DateTimePickerEnd As System.Windows.Forms.DateTimePicker
    Friend WithEvents DateTimePickerStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents Testlogsetup As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog2 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents NewFilesetup As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog3 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents NewFilesetupText As System.Windows.Forms.TextBox
    Friend WithEvents TestlogsetupText As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents CheckOpenOnEnd As System.Windows.Forms.CheckBox

End Class
