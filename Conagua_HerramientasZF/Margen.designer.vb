<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Margen
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
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

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Margen))
        Me.cbMargen = New System.Windows.Forms.ComboBox()
        Me.btnCrearMargen = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'cbMargen
        '
        Me.cbMargen.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbMargen.FormattingEnabled = True
        Me.cbMargen.Items.AddRange(New Object() {"ZF IZQUIERDA", "ZF DERECHA", "NAMO"})
        Me.cbMargen.Location = New System.Drawing.Point(12, 12)
        Me.cbMargen.Name = "cbMargen"
        Me.cbMargen.Size = New System.Drawing.Size(258, 24)
        Me.cbMargen.TabIndex = 0
        '
        'btnCrearMargen
        '
        Me.btnCrearMargen.Location = New System.Drawing.Point(12, 42)
        Me.btnCrearMargen.Name = "btnCrearMargen"
        Me.btnCrearMargen.Size = New System.Drawing.Size(258, 23)
        Me.btnCrearMargen.TabIndex = 1
        Me.btnCrearMargen.Text = "Crear Margen"
        Me.btnCrearMargen.UseVisualStyleBackColor = True
        '
        'Margen
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(282, 79)
        Me.Controls.Add(Me.btnCrearMargen)
        Me.Controls.Add(Me.cbMargen)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Margen"
        Me.Text = "Margen"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cbMargen As System.Windows.Forms.ComboBox
    Friend WithEvents btnCrearMargen As System.Windows.Forms.Button
End Class
