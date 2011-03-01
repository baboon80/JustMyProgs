Public Class FormGrid
    Private cCoupData As clsCoupData

    Public Property CoupData() As clsCoupData
        Get
            ' The Get property procedure is called when the value
            ' of a property is retrieved.
            Return cCoupData
        End Get
        Set(ByVal value As clsCoupData)
            ' The Set property procedure is called when the value 
            ' of a property is modified.  The value to be assigned
            ' is passed in the argument to Set.
            cCoupData = value
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

        If CoupData.myarraylist.Count > 0 Then
            For I As Integer = 0 To CoupData.myarraylist.Count - 1
                Dim localInstanceofMyClass As clsCoupData.myStruct = CType(cCoupData.myarraylist(I), clsCoupData.myStruct)

                With Me.DataGridView1.Rows
                    .Add(localInstanceofMyClass.dLdf, localInstanceofMyClass.dCoup _
                         , localInstanceofMyClass.dTP, localInstanceofMyClass.dTM, localInstanceofMyClass.dTPS, localInstanceofMyClass.dTPI, localInstanceofMyClass.dTPR, localInstanceofMyClass.dTPR7, localInstanceofMyClass.dTPRS _
                         , localInstanceofMyClass.dTMS, localInstanceofMyClass.dTMI, localInstanceofMyClass.dTMR, localInstanceofMyClass.dTMR7, localInstanceofMyClass.dTMRS _
                         , localInstanceofMyClass.dS, localInstanceofMyClass.dR, localInstanceofMyClass.dSS, localInstanceofMyClass.dSI, localInstanceofMyClass.dSR, localInstanceofMyClass.dSR7, localInstanceofMyClass.dSRS _
                         , localInstanceofMyClass.dRS, localInstanceofMyClass.dRI, localInstanceofMyClass.dRR, localInstanceofMyClass.dRR7, localInstanceofMyClass.dRRS)
                End With
            Next I
        End If

        Label8.BackColor = System.Drawing.Color.FromArgb(255, 162, 15, 15)

    End Sub

    Private Sub FormGrid_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        DataGridView1.Width = Me.Width - DataGridView1.Left - 30
        DataGridView1.Height = Me.Height - DataGridView1.Top - 70
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

        'Me.Width = DataGridView1.Width + 30
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
End Class