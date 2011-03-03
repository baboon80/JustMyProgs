Public Class FormGrid
    Private arr As ArrayList
    Private SaldoVerlhalbZ As String
    Private zed As New ZedGraph.ZedGraphControl
    Private myPane As ZedGraph.GraphPane
    Private loaded As Boolean = False

    Public Property CoupData() As ArrayList
        Get
            ' The Get property procedure is called when the value
            ' of a property is retrieved.
            Return arr
        End Get
        Set(ByVal value As ArrayList)
            ' The Set property procedure is called when the value 
            ' of a property is modified.  The value to be assigned
            ' is passed in the argument to Set.
            arr = value
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

    Private Enum ColNum As Integer
        eLdf = 0
        eCoup = 1
        eTP = 2
        eTM = 3
        eTPS = 4
        eTPI = 5
        eTPR = 6
        eTPR7 = 7
        eTPRS = 8
        eTMS = 9
        eTMI = 10
        eTMR = 11
        eTMR7 = 12
        eTMRS = 13
        eS = 14
        eR = 15
        eSS = 16
        eSI = 17
        eSR = 18
        eSR7 = 19
        eSRS = 20
        eRS = 21
        eRI = 22
        eRR = 23
        eRR7 = 24
        eRRS = 25
    End Enum

    Private Sub FormGrid_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Call WriteINIColSize()
        Me.Dispose()
    End Sub

    Private Sub FormGrid_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With DataGridView1
            .EditMode = False
            .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Raised
            .CellBorderStyle = DataGridViewCellBorderStyle.Single
            .GridColor = SystemColors.ActiveBorder
            .RowHeadersVisible = False

            .Columns.Add("No", "Nr.")
            .Columns.Add("Coup", "Zahl")
            .Columns.Add("TransfPlus", "+")
            .Columns.Add("TransfMinus", "-")
            .Columns.Add("TransfPlusSerie", "Ser.")
            .Columns.Add("TransfPlusIntermit", "Int.")
            .Columns.Add("TransfPlusRap", "Rap.")
            .Columns.Add("TransfPlusRap5", "+Rap.(last 3-7)")
            .Columns.Add("TransfPlusRapS", "Satz")
            .Columns.Add("TransfMinusSerie", "Ser.")
            .Columns.Add("TransfMinusIntermit", "Int.")
            .Columns.Add("TransfMinusRap", "Rap.")
            .Columns.Add("TransfMinusRap5", "-Rap.(last 3-7)")
            .Columns.Add("TransfMinusRapS", "Satz")
            .Columns.Add("Schwarz", "S")
            .Columns.Add("Rot", "R")
            .Columns.Add("SchwarzSerie", "Ser.")
            .Columns.Add("SchwarzIntermit", "Int.")
            .Columns.Add("SchwarzRap", "Rap.")
            .Columns.Add("SchwarzRap5", "+Rap.(last 3-7)")
            .Columns.Add("SchwarzRapS", "Satz")
            .Columns.Add("RotSerie", "Ser.")
            .Columns.Add("RotIntermit", "Int.")
            .Columns.Add("RotRap", "Rap.")
            .Columns.Add("RotRap5", "-Rap.(last 3-7)")
            .Columns.Add("RotRapS", "Satz")
            .Columns(ColNum.eLdf).SortMode = DataGridViewColumnSortMode.Programmatic
            .Columns(ColNum.eCoup).SortMode = DataGridViewColumnSortMode.Programmatic
            .Columns(ColNum.eTP).SortMode = DataGridViewColumnSortMode.Programmatic
            .Columns(ColNum.eTM).SortMode = DataGridViewColumnSortMode.Programmatic
            .Columns(ColNum.eTPS).SortMode = DataGridViewColumnSortMode.Programmatic
            .Columns(ColNum.eTPI).SortMode = DataGridViewColumnSortMode.Programmatic
            .Columns(ColNum.eTPR).SortMode = DataGridViewColumnSortMode.Programmatic
            .Columns(ColNum.eTPR7).SortMode = DataGridViewColumnSortMode.Programmatic
            .Columns(ColNum.eTPRS).SortMode = DataGridViewColumnSortMode.Programmatic
            .Columns(ColNum.eTMS).SortMode = DataGridViewColumnSortMode.Programmatic
            .Columns(ColNum.eTMI).SortMode = DataGridViewColumnSortMode.Programmatic
            .Columns(ColNum.eTMR).SortMode = DataGridViewColumnSortMode.Programmatic
            .Columns(ColNum.eTMR7).SortMode = DataGridViewColumnSortMode.Programmatic
            .Columns(ColNum.eTMRS).SortMode = DataGridViewColumnSortMode.Programmatic
            .Columns(ColNum.eS).SortMode = DataGridViewColumnSortMode.Programmatic
            .Columns(ColNum.eR).SortMode = DataGridViewColumnSortMode.Programmatic
            .Columns(ColNum.eSS).SortMode = DataGridViewColumnSortMode.Programmatic
            .Columns(ColNum.eSI).SortMode = DataGridViewColumnSortMode.Programmatic
            .Columns(ColNum.eSR).SortMode = DataGridViewColumnSortMode.Programmatic
            .Columns(ColNum.eSR7).SortMode = DataGridViewColumnSortMode.Programmatic
            .Columns(ColNum.eSRS).SortMode = DataGridViewColumnSortMode.Programmatic
            .Columns(ColNum.eRS).SortMode = DataGridViewColumnSortMode.Programmatic
            .Columns(ColNum.eRI).SortMode = DataGridViewColumnSortMode.Programmatic
            .Columns(ColNum.eRR).SortMode = DataGridViewColumnSortMode.Programmatic
            .Columns(ColNum.eRR7).SortMode = DataGridViewColumnSortMode.Programmatic
            .Columns(ColNum.eRRS).SortMode = DataGridViewColumnSortMode.Programmatic

            .MultiSelect = False
            .BackgroundColor = Color.Honeydew
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        End With

        Call ReadINIColSize()
        Call SetSizeLabelColPos()
        Label8.BackColor = System.Drawing.Color.FromArgb(255, 162, 15, 15)

        If arr.Count > 0 Then
            Dim R1 As DataGridViewRow
            For I As Integer = 0 To arr.Count - 1

                Dim localInstanceofMyClass = CType(arr(I), clsCoupData)
                With Me.DataGridView1.Rows
                    .Add(localInstanceofMyClass.Ldf, localInstanceofMyClass.Coup _
                         , localInstanceofMyClass.TP, localInstanceofMyClass.TM, localInstanceofMyClass.TPS, localInstanceofMyClass.TPI, localInstanceofMyClass.TPR, localInstanceofMyClass.TPR7, localInstanceofMyClass.TPRS _
                         , localInstanceofMyClass.TmS, localInstanceofMyClass.TmI, localInstanceofMyClass.TmR, localInstanceofMyClass.TmR7, localInstanceofMyClass.TmRS _
                         , localInstanceofMyClass.S, localInstanceofMyClass.R, localInstanceofMyClass.SS, localInstanceofMyClass.SI, localInstanceofMyClass.SR, localInstanceofMyClass.SR7, localInstanceofMyClass.SRS _
                         , localInstanceofMyClass.RS, localInstanceofMyClass.RI, localInstanceofMyClass.RR, localInstanceofMyClass.RR7, localInstanceofMyClass.RRS)
                End With

                R1 = DataGridView1.Rows(I)

                R1.Cells(ColNum.eCoup).Style.ForeColor = Color.Black
                If localInstanceofMyClass.Coup = 0 Then
                    R1.Cells(ColNum.eCoup).Style.BackColor = Color.Green
                    R1.Cells(ColNum.eCoup).Style.ForeColor = Color.White
                End If

                R1.Cells(ColNum.eS).Style.ForeColor = Color.Black
                R1.Cells(ColNum.eR).Style.ForeColor = System.Drawing.Color.FromArgb(255, 162, 15, 15)

                R1.Cells(ColNum.eTP).Style.BackColor = Label2.BackColor
                R1.Cells(ColNum.eTM).Style.BackColor = Label2.BackColor

                R1.Cells(ColNum.eTPS).Style.BackColor = Label3.BackColor
                R1.Cells(ColNum.eTPI).Style.BackColor = Label3.BackColor
                R1.Cells(ColNum.eTPR).Style.BackColor = Label3.BackColor
                R1.Cells(ColNum.eTPR7).Style.BackColor = Label3.BackColor
                If localInstanceofMyClass.TPRS <> "" Then
                    R1.Cells(ColNum.eTPRS).Style.BackColor = Label3.BackColor
                End If

                R1.Cells(ColNum.eTMS).Style.BackColor = Label4.BackColor
                R1.Cells(ColNum.eTMI).Style.BackColor = Label4.BackColor
                R1.Cells(ColNum.eTMR).Style.BackColor = Label4.BackColor
                R1.Cells(ColNum.eTMR7).Style.BackColor = Label4.BackColor
                If localInstanceofMyClass.TmRS <> "" Then
                    R1.Cells(ColNum.eTMRS).Style.BackColor = Label4.BackColor
                End If

                R1.Cells(ColNum.eS).Style.BackColor = Label6.BackColor
                R1.Cells(ColNum.eR).Style.BackColor = Label6.BackColor

                R1.Cells(ColNum.eSS).Style.BackColor = Label7.BackColor
                R1.Cells(ColNum.eSI).Style.BackColor = Label7.BackColor
                R1.Cells(ColNum.eSR).Style.BackColor = Label7.BackColor
                R1.Cells(ColNum.eSR7).Style.BackColor = Label7.BackColor
                R1.Cells(ColNum.eSS).Style.ForeColor = Color.White
                R1.Cells(ColNum.eSI).Style.ForeColor = Color.White
                R1.Cells(ColNum.eSR).Style.ForeColor = Color.White
                R1.Cells(ColNum.eSR7).Style.ForeColor = Color.White
                If localInstanceofMyClass.SRS <> "" Then
                    R1.Cells(ColNum.eSRS).Style.BackColor = Label7.BackColor
                    R1.Cells(ColNum.eSRS).Style.ForeColor = Color.White
                End If

                R1.Cells(ColNum.eRS).Style.BackColor = Label8.BackColor
                R1.Cells(ColNum.eRI).Style.BackColor = Label8.BackColor
                R1.Cells(ColNum.eRR).Style.BackColor = Label8.BackColor
                R1.Cells(ColNum.eRR7).Style.BackColor = Label8.BackColor
                R1.Cells(ColNum.eRS).Style.ForeColor = Color.White
                R1.Cells(ColNum.eRI).Style.ForeColor = Color.White
                R1.Cells(ColNum.eRR).Style.ForeColor = Color.White
                R1.Cells(ColNum.eRR7).Style.ForeColor = Color.White
                If localInstanceofMyClass.RRS <> "" Then
                    R1.Cells(ColNum.eRRS).Style.BackColor = Label8.BackColor
                    R1.Cells(ColNum.eRRS).Style.ForeColor = Color.White
                End If

            Next I
        End If

        zed = New ZedGraph.ZedGraphControl
        zed.Parent = Me.Panel2
        zed.Location = New Point(0, 0)

        zed.Size = New Size(Panel2.Width - 8, Panel2.Height - 8)

        CreateGraph()
        loaded = True

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub DataGridView1_ColumnWidthChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles DataGridView1.ColumnWidthChanged
        If loaded = True Then Call SetSizeLabelColPos()
    End Sub

    Private Sub FormGrid_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        DataGridView1.Width = Me.Width - DataGridView1.Left - 20
        DataGridView1.Height = Me.Height - DataGridView1.Top - 60 - Panel2.Height
        Panel2.Top = DataGridView1.Top + DataGridView1.Height + 10
        Panel2.Width = DataGridView1.Width
        zed.Size = New Size(Panel2.Width - 8, Panel2.Height - 8)
    End Sub

    Private Sub SetSizeLabelColPos()
        DataGridView1.Update()

        Label9.Left = DataGridView1.GetCellDisplayRectangle(ColNum.eLdf, -1, True).Left + 10
        Label9.Top = DataGridView1.Top - Label3.Height
        Label9.Width = DataGridView1.Columns.Item(ColNum.eLdf).Width + DataGridView1.Columns.Item(ColNum.eCoup).Width

        Label2.Left = DataGridView1.GetCellDisplayRectangle(ColNum.eTP, -1, True).Left + 10
        Label2.Top = DataGridView1.Top - Label2.Height
        Label2.Width = DataGridView1.Columns.Item(ColNum.eTP).Width + DataGridView1.Columns.Item(ColNum.eTM).Width

        Label3.Left = DataGridView1.GetCellDisplayRectangle(ColNum.eTPS, -1, True).Left + 10
        Label3.Top = DataGridView1.Top - Label3.Height
        Label3.Width = DataGridView1.Columns.Item(ColNum.eTPS).Width + DataGridView1.Columns.Item(ColNum.eTPI).Width + DataGridView1.Columns.Item(ColNum.eTPR).Width + DataGridView1.Columns.Item(ColNum.eTPR7).Width + DataGridView1.Columns.Item(ColNum.eTPRS).Width

        Label4.Left = DataGridView1.GetCellDisplayRectangle(ColNum.eTMS, -1, True).Left + 10
        Label4.Top = DataGridView1.Top - Label4.Height
        Label4.Width = DataGridView1.Columns.Item(ColNum.eTMS).Width + DataGridView1.Columns.Item(ColNum.eTMI).Width + DataGridView1.Columns.Item(ColNum.eTMR).Width + DataGridView1.Columns.Item(ColNum.eTMR7).Width + DataGridView1.Columns.Item(ColNum.eTMRS).Width


        Label6.Left = DataGridView1.GetCellDisplayRectangle(ColNum.eS, -1, True).Left + 10
        Label6.Top = DataGridView1.Top - Label2.Height
        Label6.Width = DataGridView1.Columns.Item(ColNum.eS).Width + DataGridView1.Columns.Item(ColNum.eR).Width

        Label7.Left = DataGridView1.GetCellDisplayRectangle(ColNum.eSS, -1, True).Left + 10
        Label7.Top = DataGridView1.Top - Label3.Height
        Label7.Width = DataGridView1.Columns.Item(ColNum.eSS).Width + DataGridView1.Columns.Item(ColNum.eSI).Width + DataGridView1.Columns.Item(ColNum.eSR).Width + DataGridView1.Columns.Item(ColNum.eSR7).Width + DataGridView1.Columns.Item(ColNum.eSRS).Width

        Label8.Left = DataGridView1.GetCellDisplayRectangle(ColNum.eRS, -1, True).Left + 10
        Label8.Top = DataGridView1.Top - Label4.Height
        Label8.Width = DataGridView1.Columns.Item(ColNum.eRS).Width + DataGridView1.Columns.Item(ColNum.eRI).Width + DataGridView1.Columns.Item(ColNum.eRR).Width + DataGridView1.Columns.Item(ColNum.eRR7).Width + DataGridView1.Columns.Item(ColNum.eRRS).Width

        Me.Width = DataGridView1.Width + 30
    End Sub

    Private Sub FormGrid_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

    End Sub

    Private Sub ReadINIColSize()
        Dim INI As New ClsINI

        INI.Pfad = System.AppDomain.CurrentDomain.BaseDirectory & "Option.ini"
        With DataGridView1
            .Columns(ColNum.eLdf).Width = INI.WertLesen("ColSize", "Ldf", "49")
            .Columns(ColNum.eCoup).Width = INI.WertLesen("ColSize", "Coup", "26")

            .Columns(ColNum.eTP).Width = INI.WertLesen("ColSize", "TP", "38")
            .Columns(ColNum.eTM).Width = INI.WertLesen("ColSize", "TM", "38")

            .Columns(ColNum.eTPS).Width = INI.WertLesen("ColSize", "TPS", "30")
            .Columns(ColNum.eTPI).Width = INI.WertLesen("ColSize", "TPI", "30")
            .Columns(ColNum.eTPR).Width = INI.WertLesen("ColSize", "TPR", "35")
            .Columns(ColNum.eTPR7).Width = INI.WertLesen("ColSize", "TPR6", "65")
            .Columns(ColNum.eTPRS).Width = INI.WertLesen("ColSize", "TPRS", "75")

            .Columns(ColNum.eTMS).Width = INI.WertLesen("ColSize", "TMS", "30")
            .Columns(ColNum.eTMI).Width = INI.WertLesen("ColSize", "TMI", "30")
            .Columns(ColNum.eTMR).Width = INI.WertLesen("ColSize", "TMR", "35")
            .Columns(ColNum.eTMR7).Width = INI.WertLesen("ColSize", "TMR6", "65")
            .Columns(ColNum.eTMRS).Width = INI.WertLesen("ColSize", "TMRS", "75")


            .Columns(ColNum.eS).Width = INI.WertLesen("ColSize", "S", "38")
            .Columns(ColNum.eR).Width = INI.WertLesen("ColSize", "R", "38")

            .Columns(ColNum.eSS).Width = INI.WertLesen("ColSize", "SS", "30")
            .Columns(ColNum.eSI).Width = INI.WertLesen("ColSize", "SI", "30")
            .Columns(ColNum.eSR).Width = INI.WertLesen("ColSize", "SR", "35")
            .Columns(ColNum.eSR7).Width = INI.WertLesen("ColSize", "SR6", "65")
            .Columns(ColNum.eSRS).Width = INI.WertLesen("ColSize", "SRS", "75")

            .Columns(ColNum.eRS).Width = INI.WertLesen("ColSize", "RS", "30")
            .Columns(ColNum.eRI).Width = INI.WertLesen("ColSize", "RI", "30")
            .Columns(ColNum.eRR).Width = INI.WertLesen("ColSize", "RR", "35")
            .Columns(ColNum.eRR7).Width = INI.WertLesen("ColSize", "RR6", "65")
            .Columns(ColNum.eRRS).Width = INI.WertLesen("ColSize", "RRS", "75")
        End With
    End Sub

    Private Sub WriteINIColSize()
        Dim INI As New ClsINI

        INI.Pfad = System.AppDomain.CurrentDomain.BaseDirectory & "Option.ini"
        With DataGridView1
            INI.WertSchreiben("ColSize", "Ldf", .Columns(ColNum.eLdf).Width)
            INI.WertSchreiben("ColSize", "Coup", .Columns(ColNum.eCoup).Width)

            INI.WertSchreiben("ColSize", "TP", .Columns(ColNum.eTP).Width)
            INI.WertSchreiben("ColSize", "TM", .Columns(ColNum.eTM).Width)

            INI.WertSchreiben("ColSize", "TPS", .Columns(ColNum.eTPS).Width)
            INI.WertSchreiben("ColSize", "TPI", .Columns(ColNum.eTPI).Width)
            INI.WertSchreiben("ColSize", "TPR", .Columns(ColNum.eTPR).Width)
            INI.WertSchreiben("ColSize", "TPR6", .Columns(ColNum.eTPR7).Width)
            INI.WertSchreiben("ColSize", "TPRS", .Columns(ColNum.eTPRS).Width)

            INI.WertSchreiben("ColSize", "TMS", .Columns(ColNum.eTMS).Width)
            INI.WertSchreiben("ColSize", "TMI", .Columns(ColNum.eTMI).Width)
            INI.WertSchreiben("ColSize", "TMR", .Columns(ColNum.eTMR).Width)
            INI.WertSchreiben("ColSize", "TMR6", .Columns(ColNum.eTMR7).Width)
            INI.WertSchreiben("ColSize", "TMRS", .Columns(ColNum.eTMRS).Width)

            INI.WertSchreiben("ColSize", "S", .Columns(ColNum.eS).Width)
            INI.WertSchreiben("ColSize", "R", .Columns(ColNum.eR).Width)

            INI.WertSchreiben("ColSize", "SS", .Columns(ColNum.eSS).Width)
            INI.WertSchreiben("ColSize", "SI", .Columns(ColNum.eSI).Width)
            INI.WertSchreiben("ColSize", "SR", .Columns(ColNum.eSR).Width)
            INI.WertSchreiben("ColSize", "SR6", .Columns(ColNum.eSR7).Width)
            INI.WertSchreiben("ColSize", "SRS", .Columns(ColNum.eSRS).Width)

            INI.WertSchreiben("ColSize", "RS", .Columns(ColNum.eRS).Width)
            INI.WertSchreiben("ColSize", "RI", .Columns(ColNum.eRI).Width)
            INI.WertSchreiben("ColSize", "RR", .Columns(ColNum.eRR).Width)
            INI.WertSchreiben("ColSize", "RR6", .Columns(ColNum.eRR7).Width)
            INI.WertSchreiben("ColSize", "RRS", .Columns(ColNum.eRRS).Width)
        End With
    End Sub

    Private Function GetColWidth() As Integer
        Dim width As Integer = DataGridView1.Columns.Item(ColNum.eLdf).Width + _
        DataGridView1.Columns.Item(ColNum.eCoup).Width + _
        DataGridView1.Columns.Item(ColNum.eTP).Width + _
        DataGridView1.Columns.Item(ColNum.eTM).Width + _
        DataGridView1.Columns.Item(ColNum.eTPS).Width + _
        DataGridView1.Columns.Item(ColNum.eTPI).Width + _
        DataGridView1.Columns.Item(ColNum.eTPR).Width + _
        DataGridView1.Columns.Item(ColNum.eTPR7).Width + _
        DataGridView1.Columns.Item(ColNum.eTPRS).Width + _
        DataGridView1.Columns.Item(ColNum.eTMS).Width + _
        DataGridView1.Columns.Item(ColNum.eTMI).Width + _
        DataGridView1.Columns.Item(ColNum.eTMR).Width + _
        DataGridView1.Columns.Item(ColNum.eTMR7).Width + _
        DataGridView1.Columns.Item(ColNum.eTMRS).Width + _
        DataGridView1.Columns.Item(ColNum.eS).Width + _
        DataGridView1.Columns.Item(ColNum.eR).Width + _
        DataGridView1.Columns.Item(ColNum.eSS).Width + _
        DataGridView1.Columns.Item(ColNum.eSI).Width + _
        DataGridView1.Columns.Item(ColNum.eSR).Width + _
        DataGridView1.Columns.Item(ColNum.eSR7).Width + _
        DataGridView1.Columns.Item(ColNum.eSRS).Width + _
        DataGridView1.Columns.Item(ColNum.eRS).Width + _
        DataGridView1.Columns.Item(ColNum.eRI).Width + _
        DataGridView1.Columns.Item(ColNum.eRR).Width + _
        DataGridView1.Columns.Item(ColNum.eRR7).Width + _
        DataGridView1.Columns.Item(ColNum.eRRS).Width
        GetColWidth = width + 20
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Width = Me.Width
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

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Width = Me.Width
    End Sub
End Class