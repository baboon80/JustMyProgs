

Public Class FormRes

    Private zed As New ZedGraph.ZedGraphControl

    Private propRes As String
    Private SaldoVerl1 As String
    Private SaldoVerlNoZ As String

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

    Public Property SaldoVerlauf1() As String
        Get
            ' The Get property procedure is called when the value
            ' of a property is retrieved.
            Return SaldoVerl1
        End Get
        Set(ByVal value As String)
            ' The Set property procedure is called when the value 
            ' of a property is modified.  The value to be assigned
            ' is passed in the argument to Set.
            SaldoVerl1 = value
        End Set
    End Property

    Public Property SaldoVerlaufNoZero() As String
        Get
            ' The Get property procedure is called when the value
            ' of a property is retrieved.
            Return SaldoVerlNoZ
        End Get
        Set(ByVal value As String)
            ' The Set property procedure is called when the value 
            ' of a property is modified.  The value to be assigned
            ' is passed in the argument to Set.
            SaldoVerlNoZ = value
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
        Call Close()

    End Sub

    Private Sub FormRes_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        TextBox1.SelectionStart = TextBox1.TextLength
        TextBox1.ScrollToCaret()
        CreateGraph()

    End Sub

    Private Sub CreateGraph()
        Dim myPane As ZedGraph.GraphPane

        myPane = zed.GraphPane

        myPane.Title.Text = "Saldo Verlauf"
        myPane.XAxis.Title.Text = "Satz"
        myPane.YAxis.Title.Text = "Saldo"

        Dim x As Integer = 0
        Dim y1 As Integer
        Dim list1 As New ZedGraph.PointPairList
        Dim CntS As Integer = 0
        Dim x2 As Integer = 0
        Dim y2 As Integer
        Dim list2 As New ZedGraph.PointPairList
        Dim CntS2 As Integer = 0


        Dim len As Integer = SaldoVerl1.Length

        If SaldoVerl1 <> "" Then
            Do While CntS < len

                If SaldoVerl1.Substring(CntS, 1) = "+" Then
                    y1 = y1 + 1
                    x = x + 1
                    list1.Add(x, y1)
                ElseIf SaldoVerl1.Substring(CntS, 1) = "-" Then
                    y1 = y1 - 1
                    x = x + 1
                    list1.Add(x, y1)
                End If
                CntS = CntS + 1
            Loop
        End If

        Dim len2 As Integer = SaldoVerlNoZ.Length
        If SaldoVerlNoZ <> "" Then
            Do While CntS2 < len2
                If SaldoVerlNoZ.Substring(CntS2, 1) = "+" Then
                    y2 = y2 + 1
                    x2 = x2 + 1
                    list2.Add(x2, y2)
                ElseIf SaldoVerlNoZ.Substring(CntS2, 1) = "-" Then
                    y2 = y2 - 1
                    x2 = x2 + 1
                    list2.Add(x2, y2)
                End If
                CntS2 = CntS2 + 1
            Loop
        End If

        Dim myCurve As ZedGraph.LineItem
        Dim myCurve2 As ZedGraph.LineItem
        myCurve = myPane.AddCurve("Saldo inkl. Zero", list1, Color.Blue, ZedGraph.SymbolType.None)
        myCurve2 = myPane.AddCurve("Saldo ohne Zero", list2, Color.Green, ZedGraph.SymbolType.None)
        zed.AxisChange()

    End Sub

    Private Sub HScrollBar1_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles HScrollBar1.Scroll
        zed.Size = New Size(Panel1.Width + (Panel1.Width * (HScrollBar1.Value / 10)), Panel1.Height - 10)
        zed.AxisChange()
    End Sub
End Class