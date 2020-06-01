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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.LblShipName = New System.Windows.Forms.Label()
        Me.LblEmbarkDate = New System.Windows.Forms.Label()
        Me.LblNumberOfDays = New System.Windows.Forms.Label()
        Me.DtpEmbarkation = New System.Windows.Forms.DateTimePicker()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.CboShipPicker = New System.Windows.Forms.ComboBox()
        Me.CboNumberOfDays = New System.Windows.Forms.ComboBox()
        Me.lblVoyageNumber = New System.Windows.Forms.Label()
        Me.txtVoyageNumber = New System.Windows.Forms.TextBox()
        Me.lblVersionNumber = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'LblShipName
        '
        Me.LblShipName.AutoSize = True
        Me.LblShipName.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblShipName.Location = New System.Drawing.Point(12, 9)
        Me.LblShipName.Name = "LblShipName"
        Me.LblShipName.Size = New System.Drawing.Size(41, 20)
        Me.LblShipName.TabIndex = 0
        Me.LblShipName.Text = "Ship"
        '
        'LblEmbarkDate
        '
        Me.LblEmbarkDate.AutoSize = True
        Me.LblEmbarkDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblEmbarkDate.Location = New System.Drawing.Point(13, 47)
        Me.LblEmbarkDate.Name = "LblEmbarkDate"
        Me.LblEmbarkDate.Size = New System.Drawing.Size(138, 20)
        Me.LblEmbarkDate.TabIndex = 1
        Me.LblEmbarkDate.Text = "Embarkation Date"
        '
        'LblNumberOfDays
        '
        Me.LblNumberOfDays.AutoSize = True
        Me.LblNumberOfDays.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblNumberOfDays.Location = New System.Drawing.Point(13, 82)
        Me.LblNumberOfDays.Name = "LblNumberOfDays"
        Me.LblNumberOfDays.Size = New System.Drawing.Size(123, 20)
        Me.LblNumberOfDays.TabIndex = 2
        Me.LblNumberOfDays.Text = "Number of Days"
        '
        'DtpEmbarkation
        '
        Me.DtpEmbarkation.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DtpEmbarkation.Location = New System.Drawing.Point(172, 42)
        Me.DtpEmbarkation.Name = "DtpEmbarkation"
        Me.DtpEmbarkation.Size = New System.Drawing.Size(391, 26)
        Me.DtpEmbarkation.TabIndex = 5
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(16, 149)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(547, 54)
        Me.Button1.TabIndex = 6
        Me.Button1.Text = "Create Itinerary File on Desktop"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'CboShipPicker
        '
        Me.CboShipPicker.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.CboShipPicker.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.CboShipPicker.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboShipPicker.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboShipPicker.FormattingEnabled = True
        Me.CboShipPicker.Location = New System.Drawing.Point(172, 6)
        Me.CboShipPicker.Name = "CboShipPicker"
        Me.CboShipPicker.Size = New System.Drawing.Size(391, 28)
        Me.CboShipPicker.TabIndex = 7
        '
        'CboNumberOfDays
        '
        Me.CboNumberOfDays.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append
        Me.CboNumberOfDays.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.CboNumberOfDays.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboNumberOfDays.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboNumberOfDays.FormattingEnabled = True
        Me.CboNumberOfDays.Location = New System.Drawing.Point(173, 74)
        Me.CboNumberOfDays.Name = "CboNumberOfDays"
        Me.CboNumberOfDays.Size = New System.Drawing.Size(389, 28)
        Me.CboNumberOfDays.TabIndex = 8
        '
        'lblVoyageNumber
        '
        Me.lblVoyageNumber.AutoSize = True
        Me.lblVoyageNumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVoyageNumber.Location = New System.Drawing.Point(13, 118)
        Me.lblVoyageNumber.Name = "lblVoyageNumber"
        Me.lblVoyageNumber.Size = New System.Drawing.Size(123, 20)
        Me.lblVoyageNumber.TabIndex = 9
        Me.lblVoyageNumber.Text = "Voyage Number"
        '
        'txtVoyageNumber
        '
        Me.txtVoyageNumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVoyageNumber.Location = New System.Drawing.Point(172, 112)
        Me.txtVoyageNumber.MaxLength = 5
        Me.txtVoyageNumber.Name = "txtVoyageNumber"
        Me.txtVoyageNumber.Size = New System.Drawing.Size(390, 26)
        Me.txtVoyageNumber.TabIndex = 10
        '
        'lblVersionNumber
        '
        Me.lblVersionNumber.AutoSize = True
        Me.lblVersionNumber.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.lblVersionNumber.Location = New System.Drawing.Point(14, 219)
        Me.lblVersionNumber.Name = "lblVersionNumber"
        Me.lblVersionNumber.Size = New System.Drawing.Size(0, 13)
        Me.lblVersionNumber.TabIndex = 11
        Me.lblVersionNumber.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(585, 241)
        Me.Controls.Add(Me.lblVersionNumber)
        Me.Controls.Add(Me.txtVoyageNumber)
        Me.Controls.Add(Me.lblVoyageNumber)
        Me.Controls.Add(Me.CboNumberOfDays)
        Me.Controls.Add(Me.CboShipPicker)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.DtpEmbarkation)
        Me.Controls.Add(Me.LblNumberOfDays)
        Me.Controls.Add(Me.LblEmbarkDate)
        Me.Controls.Add(Me.LblShipName)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(601, 280)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(601, 280)
        Me.Name = "Form1"
        Me.Text = "IQWF - COVID-19 Itinerary Generator (CIG)"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents LblShipName As Label
    Friend WithEvents LblEmbarkDate As Label
    Friend WithEvents LblNumberOfDays As Label
    Friend WithEvents DtpEmbarkation As DateTimePicker
    Friend WithEvents Button1 As Button
    Friend WithEvents CboShipPicker As ComboBox
    Friend WithEvents CboNumberOfDays As ComboBox
    Friend WithEvents lblVoyageNumber As Label
    Friend WithEvents txtVoyageNumber As TextBox
    Friend WithEvents lblVersionNumber As Label
End Class
