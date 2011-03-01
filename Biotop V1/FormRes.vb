

Public Class FormRes

    Private zed As New ZedGraph.ZedGraphControl
    Private myPane As ZedGraph.GraphPane

    Private propRes As String
    Private SaldoVerlhalbZ As String

    Public Property prop1() As String
        Get
            ' The Get property procedure is called when the value
            ' of a property is retrieved.
            Return propRes
        End Get
        Set(ByVal value As String)
            ' The Set property procedure is called when the value 
            ' of a property is modified.  The value to be assigned
            ' is passed in the argument to Set.
            propRes = value
        End Set
    End Property

    Public Property SaldoVerlaufhalbZero() As String
        Get
            ' The Get property procedure is called when the value
            ' of a property is retrieved.
            Return SaldoVerlhalbZ
        End Get
        Set(ByVal value As String)
            ' The Set property procedure is called when the value 
            ' of a property is modified.  The value to be assigned
            ' is passed in the argument to Set.
            SaldoVerlhalbZ = value
        End Set
    End Property

    Private Sub FormRes_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.TextBox1.Text = propRes
        Button1.Focus()
        'Set the cursor to the end of the textbox.


        zed = New ZedGraph.ZedGraphControl
        zed.Parent = Me.Panel1
        zed.Location = New Point(0, 0)

        zed.Size = New Size(Panel1.Width, Panel1.Height - 10)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        myPane = Nothing
        zed.GraphPane.CurveList.Clear()
        zed = Nothing
        Me.Dispose()

    End Sub

    Private Sub FormRes_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        TextBox1.SelectionStart = TextBox1.TextLength
        TextBox1.ScrollToCaret()
        CreateGraph()

    End Sub

    Private Sub CreateGraph()


        myPane = zed.GraphPane

        myPane.Title.Text = "Saldo Verlauf"
        myPane.XAxis.Title.Text = "Satz"
        myPane.YAxis.Title.Text = "Saldo"

        Dim x As Integer = 0
        Dim y1 As Integer
        Dim list1 As New ZedGraph.PointPairList
        Dim CntS As Integer = 0
        Dim x3 As Integer = 0
        Dim y3 As Integer
        Dim list3 As New ZedGraph.PointPairList
        Dim CntS3 As Integer = 0


        If Not SaldoVerlhalbZ = Nothing Then
            Dim len As Integer = SaldoVerlhalbZ.Length
            If SaldoVerlhalbZ <> "" Then
                Do While CntS < len

                    If SaldoVerlhalbZ.Substring(CntS, 1) = "+" Then
                        y1 = y1 + 1
                        x = x + 1
                        list1.Add(x, y1)
                    ElseIf SaldoVerlhalbZ.Substring(CntS, 1) = "-" Then
                        y1 = y1 - 1
                        x = x + 1
                        list1.Add(x, y1)
                    End If
                    CntS = CntS + 1
                Loop
            End If

            Dim len3 As Integer = SaldoVerlhalbZ.Length
            If SaldoVerlhalbZ <> "" Then
                Do While CntS3 < len3
                    If SaldoVerlhalbZ.Substring(CntS3, 1) = "+" Then
                        y3 = y3 + 1
                        x3 = x3 + 1
                        list3.Add(x3, y3)
                    ElseIf SaldoVerlhalbZ.Substring(CntS3, 1) = "-" Then
                        y3 = y3 - 1
                        x3 = x3 + 1
                        list3.Add(x3, y3)
                    ElseIf SaldoVerlhalbZ.Substring(CntS3, 1) = "Z" Then
                        y3 = y3 - 0.5
                        x3 = x3 + 1
                        list3.Add(x3, y3)
                    End If
                    CntS3 = CntS3 + 1
                Loop
            End If

            Dim myCurve1 As ZedGraph.LineItem
            Dim myCurve2 As ZedGraph.LineItem
            myCurve1 = myPane.AddCurve("Saldo mit Zero", list3, Color.Orange, ZedGraph.SymbolType.None)
            myCurve2 = myPane.AddCurve("Saldo ohne Zero", list1, Color.Green, ZedGraph.SymbolType.None)
            zed.AxisChange()
        End If
    End Sub

    Private Sub HScrollBar1_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles HScrollBar1.Scroll
        zed.Size = New Size(Panel1.Width + (Panel1.Width * (HScrollBar1.Value / 10)), Panel1.Height - 10)
        zed.AxisChange()
    End Sub
End Class