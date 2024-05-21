<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form3
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.butImport = New System.Windows.Forms.Button()
        Me.butUpdate = New System.Windows.Forms.Button()
        Me.lblCount = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.Label2 = New System.Windows.Forms.Label()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'butImport
        '
        Me.butImport.Location = New System.Drawing.Point(66, 40)
        Me.butImport.Name = "butImport"
        Me.butImport.Size = New System.Drawing.Size(234, 23)
        Me.butImport.TabIndex = 5
        Me.butImport.Text = "Import Yahoo historical data (*.csv)"
        Me.butImport.UseVisualStyleBackColor = True
        '
        'butUpdate
        '
        Me.butUpdate.Location = New System.Drawing.Point(66, 79)
        Me.butUpdate.Name = "butUpdate"
        Me.butUpdate.Size = New System.Drawing.Size(216, 23)
        Me.butUpdate.TabIndex = 6
        Me.butUpdate.Text = "Update database"
        Me.butUpdate.UseVisualStyleBackColor = True
        '
        'lblCount
        '
        Me.lblCount.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.lblCount.Location = New System.Drawing.Point(78, 161)
        Me.lblCount.Name = "lblCount"
        Me.lblCount.Size = New System.Drawing.Size(80, 25)
        Me.lblCount.TabIndex = 9
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(78, 137)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(67, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Import Count"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Menu
        Me.Label2.Location = New System.Drawing.Point(32, 218)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(367, 63)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Rename the CSV files before importing if you want a different ticker symbol name " &
    "in the database. For example, if you want BRK-B.csv to be BRK.B.csv or BRK/B.csv" &
    "."
        '
        'Form3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.ClientSize = New System.Drawing.Size(460, 360)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblCount)
        Me.Controls.Add(Me.butUpdate)
        Me.Controls.Add(Me.butImport)
        Me.Name = "Form3"
        Me.Text = "Import CSV Fie"
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents butImport As Button
    Friend WithEvents butUpdate As Button
    Friend WithEvents lblCount As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents ErrorProvider1 As ErrorProvider
    Friend WithEvents Label2 As Label
End Class
