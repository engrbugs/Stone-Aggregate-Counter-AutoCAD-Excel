Public Class Form1
    Inherits System.Windows.Forms.Form


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents T As System.Windows.Forms.TextBox
    Friend WithEvents P As System.Windows.Forms.ProgressBar
    Friend WithEvents txtFile As System.Windows.Forms.TextBox
    Friend WithEvents O As System.Windows.Forms.OpenFileDialog
    Friend WithEvents B As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
        Me.T = New System.Windows.Forms.TextBox
        Me.B = New System.Windows.Forms.Button
        Me.P = New System.Windows.Forms.ProgressBar
        Me.txtFile = New System.Windows.Forms.TextBox
        Me.O = New System.Windows.Forms.OpenFileDialog
        Me.SuspendLayout()
        '
        'T
        '
        Me.T.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.T.Location = New System.Drawing.Point(0, 48)
        Me.T.Multiline = True
        Me.T.Name = "T"
        Me.T.ReadOnly = True
        Me.T.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.T.Size = New System.Drawing.Size(480, 200)
        Me.T.TabIndex = 2
        Me.T.Text = ""
        '
        'B
        '
        Me.B.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.B.Location = New System.Drawing.Point(424, 0)
        Me.B.Name = "B"
        Me.B.Size = New System.Drawing.Size(56, 24)
        Me.B.TabIndex = 0
        Me.B.Text = "&Browse"
        '
        'P
        '
        Me.P.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.P.Location = New System.Drawing.Point(0, 24)
        Me.P.Name = "P"
        Me.P.Size = New System.Drawing.Size(480, 24)
        Me.P.TabIndex = 3
        '
        'txtFile
        '
        Me.txtFile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFile.Location = New System.Drawing.Point(0, 1)
        Me.txtFile.Name = "txtFile"
        Me.txtFile.ReadOnly = True
        Me.txtFile.Size = New System.Drawing.Size(424, 21)
        Me.txtFile.TabIndex = 1
        Me.txtFile.Text = "None"
        '
        'O
        '
        Me.O.Multiselect = True
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
        Me.ClientSize = New System.Drawing.Size(480, 246)
        Me.Controls.Add(Me.txtFile)
        Me.Controls.Add(Me.P)
        Me.Controls.Add(Me.B)
        Me.Controls.Add(Me.T)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ISIP BATO by BUGS"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        


    End Sub
    Public Function splineLength(ByVal spline As AutoCAD.AcadSpline) As Double
        Dim pt1(1) As Double
        Dim pt2(1) As Double

        Dim i As Integer
        Dim j As Integer

        splineLength = 0

        Dim reder As Double

        For i = 0 To spline.NumberOfFitPoints - 1
            For j = 0 To spline.NumberOfFitPoints - 1
                pt1 = spline.GetFitPoint(i)
                pt2 = spline.GetFitPoint(j)
                reder = (((pt1(0) - pt2(0)) ^ 2) + ((pt1(1) - pt2(1)) ^ 2)) ^ 0.5
                If reder > splineLength Then splineLength = reder
            Next
        Next

    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles B.Click
        O.Title = "Open Autocad File"
        O.Filter = "Autocad File(*.dwg)|*.dwg"
        If O.ShowDialog() = DialogResult.OK Then
            Dim file As String
            B.Enabled = False
            txtFile.Text = ""
            For Each file In O.FileNames
                txtFile.Text &= file & "; "
            Next

            T.Text = ""


            Dim acApp As AutoCAD.AcadApplication
            Dim Dwg As AutoCAD.AcadDocument
            Dim pebble As Integer = 1

            Const cntspl = "AcDbSpline"

            Try
                acApp = GetObject(, "AutoCAD.Application")
            Catch ex As Exception
                acApp = CreateObject("AutoCAD.Application")
            End Try
            acApp.Visible = False


            Dim no As Integer = 1
            For Each file In O.FileNames
                Dwg = acApp.Documents.Open(file)
                Dim looper As Integer
                Dim spl As AutoCAD.AcadSpline
                Dim pt(1) As Double
                T.Text &= file & vbNewLine
                T.Text &= "pebble, Length, Area" & vbNewLine
                P.Value = 0
                Me.Text = "ISIP BATO by BUGS - " & no & "/" & O.FileNames.Length
                P.Refresh()
                Application.DoEvents()
                pebble = 1
                For looper = 0 To Dwg.ModelSpace.Count - 1
                    P.Value = (looper + 1) * 100 / Dwg.ModelSpace.Count
                    Me.Text = "ISIP BATO by BUGS - " & no & "/" & O.FileNames.Length
                    P.Refresh()
                    Application.DoEvents()


                    If Dwg.ModelSpace.Item(looper).ObjectName = cntspl Then

                        spl = Dwg.ModelSpace.Item(looper)
                        T.Text &= pebble & ", " & Format(splineLength(spl), "0.00") & ", " & Format(spl.Area, "0.00") & vbNewLine
                        pebble += 1
                    End If
                Next
                no += 1
            Next
            Dwg.Close()
            acApp.Quit()
            MsgBox("ISIP BATO is Finished.")
            B.Enabled = True
        End If
    End Sub

    Private Sub T_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles T.TextChanged

    End Sub
End Class
