Imports System.IO
Imports System
Imports System.ComponentModel
Imports System.Threading
Imports System.Windows.Forms


Public Class Form1
    Private PermFile As String
    Private RowCount As Integer
    Private SatzCount As Integer
    Private loaded As Boolean = False
    Private SatzPlus As String
    Private SatzMinus As String
    Private SatzSchwarz As String
    Private SatzRot As String
    Private saldo As Double
    Private SatzCnt As Integer
    Private testCnt As Double
    Private zeroCnt As Double
    Private Maxloss As Double
    Private GesMaxloss As Double
    Private CntGewinn As Double
    Private CntVerlust As Integer
    Private MaxSaldo As Double
    Private MinSaldo As Double
    Private lastsaldo As Double
    Private datum As String
    Private ResTxt As String    
    Private InvertRegel As Boolean = False
    Private gesSaldoVerlaufhalbZero As String = ""

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

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing    
        Call WriteINIColSize()
    End Sub

    Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Windows.Forms.Keys.F Then
            If e.Control = True Then

                'FormFilter.Datum() = FilterDatum
                'FormFilter.Invert() = InvertRegel
                'Call FormFilter.ShowDialog()
                'FilterDatum = FormFilter.Datum()
                'InvertRegel = FormFilter.Invert()

            End If
        End If
    End Sub


    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

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

        ToolStripStatusLabel1.Text = "Bitte wählen Sie zum auswerten eine Permanenz Datei aus, oder geben Sie manuell Zahlen ein!"

        Label8.BackColor = System.Drawing.Color.FromArgb(255, 162, 15, 15)

        ToolStripProgressBar1.Visible = False


        loaded = True
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


    Private Sub StartAuswertung()

        Dim NewCoup As String
        Dim lNewCoup As Long
        Dim line As String
        Dim nextDay As Boolean
        Dim TmpMaxSaldo As Double = 0
        Dim GesSaldo As Double = 0
        Dim GesZero As Double = 0
        Dim GesSatzCnt As Integer = 0
        Dim GlobalGesMaxloss As Integer = 0
        Dim filterfound As Boolean = True
        Dim saldoverlauf As String = ""

        If File.Exists(PermFile) = False Then
            Exit Sub
        End If

        RowCount = 0
        SatzCount = 0
        saldo = 0
        SatzPlus = ""
        SatzMinus = ""
        SatzSchwarz = ""
        SatzRot = ""        
        SatzCnt = 0
        testCnt = 0
        zeroCnt = 0
        GesMaxloss = 0
        Maxloss = 0
        CntGewinn = 0
        CntVerlust = 0
        MaxSaldo = 0
        MinSaldo = 0     
        datum = ""
        nextDay = False
        ResTxt = ""
        GesSaldo = 0
        lastsaldo = 0
        gesSaldoVerlaufhalbZero = ""

        DataGridView1.Rows.Clear()
        Using r As StreamReader = New StreamReader(TextBox1.Text)
                        line = r.ReadLine
            Do While (Not line Is Nothing)
                NewCoup = Trim(line)
                If IsNumeric(NewCoup) = True And (Len(NewCoup) = 2 Or Len(NewCoup) = 1) And nextDay = False Then
                    If filterfound = True Then
                        lNewCoup = NewCoup
                        Call InsertNewCoup(NewCoup)
                        

                        If saldo > lastsaldo Then
                            gesSaldoVerlaufhalbZero = gesSaldoVerlaufhalbZero & "+"
                            saldoverlauf = saldoverlauf & "+"
                        ElseIf saldo < lastsaldo Then
                            If lNewCoup > 0 Then
                                gesSaldoVerlaufhalbZero = gesSaldoVerlaufhalbZero & "-"
                                saldoverlauf = saldoverlauf & "-"
                            Else
                                saldo = saldo + 1 'zero Steuer wieder drauf, nur nötig damit erkannt wird das es Verlust gab
                                saldoverlauf = saldoverlauf & "Z"
                                gesSaldoVerlaufhalbZero = gesSaldoVerlaufhalbZero & "Z"
                            End If
                        End If

                        If (saldo = 2 Or saldo = 2.5) And TmpMaxSaldo >= 3 Then
                            nextDay = True
                        ElseIf (saldo = 3 Or saldo = 3.5) And TmpMaxSaldo >= 5 Then
                            nextDay = True
                        ElseIf (saldo = 5 Or saldo = 5.5) And TmpMaxSaldo >= 8 Then
                            nextDay = True
                        ElseIf (saldo = 8 Or saldo = 8.5) And TmpMaxSaldo >= 11 Then
                            nextDay = True
                        ElseIf (saldo = 11 Or saldo = 11.5) And TmpMaxSaldo >= 14 Then
                            nextDay = True
                        ElseIf (saldo = 14 Or saldo = 14.5) And TmpMaxSaldo >= 17 Then
                            nextDay = True
                        ElseIf (saldo = 17 Or saldo = 17.5) And TmpMaxSaldo >= 20 Then
                            nextDay = True
                        ElseIf saldo = 20 Or saldo = 20.5 Then
                            nextDay = True
                        ElseIf saldo = -6 Then
                            nextDay = True
                        End If

                        If TmpMaxSaldo < saldo Then TmpMaxSaldo = saldo

                        lastsaldo = saldo                        
                    End If
                Else
                    If InStr(line, "Datum") Then
                        If datum = "" Then
                            datum = Trim(line.Substring(6))

                            If TextBox2.Text <> "" Then
                                If datum = TextBox2.Text Then
                                    filterfound = True
                                Else
                                    filterfound = False
                                    datum = ""
                                End If
                            End If
                        Else
                            If TextBox2.Text <> "" Then
                                If datum = TextBox2.Text And nextDay = True Then
                                    Exit Do
                                End If
                            End If
                        End If

                        If nextDay = True Then
                            GesSaldo = GesSaldo + saldo

                            ResTxt = ResTxt & datum & ": Saldo = " & saldo & " | Zero Verlust = " & (zeroCnt / 2) & " | Anz. Sätze = " & SatzCnt & " | Max. Verlustserie = " & GesMaxloss & vbNewLine
                            ResTxt = ResTxt & "Saldoverlauf: " & vbNewLine & saldoverlauf & vbNewLine & vbNewLine

                            If GlobalGesMaxloss < GesMaxloss Then
                                GlobalGesMaxloss = GesMaxloss
                            End If
                            GesZero = GesZero + (zeroCnt / 2)
                            GesSatzCnt = GesSatzCnt + SatzCnt

                            saldo = 0
                            lastsaldo = 0
                            saldoverlauf = ""
                            SatzPlus = ""
                            SatzMinus = ""
                            SatzSchwarz = ""
                            SatzRot = ""
                            SatzCnt = 0
                            testCnt = 0
                            zeroCnt = 0
                            GesMaxloss = 0
                            Maxloss = 0
                            CntGewinn = 0
                            CntVerlust = 0
                            MaxSaldo = 0
                            MinSaldo = 0

                            datum = ""
                            nextDay = False
                            TmpMaxSaldo = 0
                            GesMaxloss = 0
                            DataGridView1.DataSource = "null"
                            DataGridView1.DataSource = Nothing
                            RowCount = 0

                            datum = Trim(line.Substring(6))
                        End If
                    End If
                End If

                line = r.ReadLine
            Loop

            If SatzCnt > 0 Then
                GesSaldo = GesSaldo + saldo
                If datum = "" Then
                    datum = Now.Date
                End If
                ResTxt = ResTxt & datum & ": Saldo = " & saldo & " | Zero Verlust = " & (zeroCnt / 2) & " | Anz. Sätze = " & SatzCnt & " | Max. Verlustserie = " & GesMaxloss & vbNewLine
                ResTxt = ResTxt & "Saldoverlauf: " & vbNewLine & saldoverlauf & vbNewLine & vbNewLine
            End If
            GesZero = GesZero + (zeroCnt / 2)
            GesSatzCnt = GesSatzCnt + SatzCnt
            If GlobalGesMaxloss < GesMaxloss Then
                GlobalGesMaxloss = GesMaxloss
            End If

            If filterfound = False Then
                ResTxt = ResTxt & "Datum wurde in Permanenz nicht gefunden"
            Else
                ResTxt = ResTxt & vbNewLine & "Saldo Ohne Zero: " & GesSaldo & vbNewLine
                ResTxt = ResTxt & "Saldo mit Zero: " & GesSaldo - GesZero & vbNewLine
                ResTxt = ResTxt & "Zero Verlust: " & GesZero & vbNewLine
                ResTxt = ResTxt & "Gesamt Anzahl Sätze: " & GesSatzCnt & vbNewLine
                ResTxt = ResTxt & "Größte Verlust Serie (Ohne Zero): " & GlobalGesMaxloss & vbNewLine
            End If
        End Using

        If RowCount > 0 Then
            DataGridView1.FirstDisplayedScrollingRowIndex = RowCount - 1
            Me.DataGridView1.Rows(RowCount - 1).Selected = True
        End If
        ToolStripProgressBar1.Visible = False
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub InsertNewCoup(ByVal NewCoup As String, Optional ByVal dummySet As Boolean = False)

        Dim R1 As DataGridViewRow
        Dim lNewCoup As Integer
        Dim Col As Integer


        lNewCoup = NewCoup
        Col = CheckCollong(lNewCoup) 'prüft die Farbe

        With Me.DataGridView1.Rows
            .Add(RowCount + 1, NewCoup)
        End With

        R1 = DataGridView1.Rows(RowCount)

        If lNewCoup = 0 Then
            R1.Cells(ColNum.eCoup).Style.ForeColor = Color.White
            R1.Cells(ColNum.eCoup).Style.BackColor = Color.Green
        Else
            If Col = 1 Then
                R1.Cells(ColNum.eCoup).Style.ForeColor = Color.Black
                R1.Cells(ColNum.eS).Value = "S"
                R1.Cells(ColNum.eS).Style.ForeColor = Color.Black
            Else
                R1.Cells(ColNum.eCoup).Style.ForeColor = System.Drawing.Color.FromArgb(255, 162, 15, 15)
                R1.Cells(ColNum.eR).Value = "R"
                R1.Cells(ColNum.eR).Style.ForeColor = System.Drawing.Color.FromArgb(255, 162, 15, 15)
            End If

            Call SetTransformator(R1)
            Call fillSelektorPlus()
            Call fillSelektorMinus()
            Call fillSelektorPlusRap()
            Call fillSelektorMinusRap()
            Call fillSelektorSchwarz()
            Call fillSelektorRot()
            Call fillSelektorSchwarzRap()
            Call fillSelektorRotRap()

            If dummySet = True Then Exit Sub

            R1.Cells(ColNum.eTPRS).Value = FillSatz(ColNum.eTPR7)
            R1.Cells(ColNum.eTMRS).Value = FillSatz(ColNum.eTMR7)
            R1.Cells(ColNum.eSRS).Value = FillSatz(ColNum.eSR7)
            R1.Cells(ColNum.eRRS).Value = FillSatz(ColNum.eRR7)
        End If

        If dummySet = True Then Exit Sub

        Call CanISetPlus()
        Call CanISetMinus()
        Call CanISetSchwarz()
        Call CanISetRot()

        Call CalkSaldo()

        R1.Cells(ColNum.eTP).Style.BackColor = Label2.BackColor
        R1.Cells(ColNum.eTM).Style.BackColor = Label2.BackColor

        R1.Cells(ColNum.eTPS).Style.BackColor = Label3.BackColor
        R1.Cells(ColNum.eTPI).Style.BackColor = Label3.BackColor
        R1.Cells(ColNum.eTPR).Style.BackColor = Label3.BackColor
        R1.Cells(ColNum.eTPR7).Style.BackColor = Label3.BackColor

        R1.Cells(ColNum.eTMS).Style.BackColor = Label4.BackColor
        R1.Cells(ColNum.eTMI).Style.BackColor = Label4.BackColor
        R1.Cells(ColNum.eTMR).Style.BackColor = Label4.BackColor
        R1.Cells(ColNum.eTMR7).Style.BackColor = Label4.BackColor

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

        R1.Cells(ColNum.eRS).Style.BackColor = Label8.BackColor
        R1.Cells(ColNum.eRI).Style.BackColor = Label8.BackColor
        R1.Cells(ColNum.eRR).Style.BackColor = Label8.BackColor
        R1.Cells(ColNum.eRR7).Style.BackColor = Label8.BackColor
        R1.Cells(ColNum.eRS).Style.ForeColor = Color.White
        R1.Cells(ColNum.eRI).Style.ForeColor = Color.White
        R1.Cells(ColNum.eRR).Style.ForeColor = Color.White
        R1.Cells(ColNum.eRR7).Style.ForeColor = Color.White

        RowCount = RowCount + 1
    End Sub

    Private Sub SetTransformator(ByRef R1 As DataGridViewRow)
        Dim Erg As Integer
        Erg = SetTransformatorRegel1()
        If Erg = 1 Then
            R1.Cells(ColNum.eTP).Value = "+ (R1)"
        ElseIf Erg = 2 Then
            R1.Cells(ColNum.eTM).Value = "- (R1)"
        End If

        If Erg = 0 Then
            Erg = SetTransformatorRegel2()
            If Erg = 1 Then
                R1.Cells(ColNum.eTP).Value = "+ (R2)"
            ElseIf Erg = 2 Then
                R1.Cells(ColNum.eTM).Value = "- (R2)"
            End If
        End If

        If Erg = 0 Then
            Erg = SetTransformatorRegel3()
            If Erg = 1 Then
                R1.Cells(ColNum.eTP).Value = "+ (R3)"
            ElseIf Erg = 2 Then
                R1.Cells(ColNum.eTM).Value = "- (R3)"
            End If
        End If

        If Erg = 0 Then
            Erg = SetTransformatorRegel4()
            If Erg = 1 Then
                R1.Cells(ColNum.eTP).Value = "+ (R4)"
            ElseIf Erg = 2 Then
                R1.Cells(ColNum.eTM).Value = "- (R4)"
            End If
        End If

        If Erg = 0 And RowCount > 3 Then
            Erg = SetTransformatorRegel5()
            If Erg = 1 Then
                R1.Cells(ColNum.eTP).Value = "+ (R5)"
            ElseIf Erg = 2 Then
                R1.Cells(ColNum.eTM).Value = "- (R5)"
            End If
        End If

        If Erg = 0 Then
            Erg = SetTransformatorRegel5_1()
            If Erg = 1 Then
                R1.Cells(ColNum.eTP).Value = "+ (R5)"
            ElseIf Erg = 2 Then
                R1.Cells(ColNum.eTM).Value = "- (R5)"
            End If
        End If

    End Sub

    Private Sub CalkSaldo()
        Dim SetTo As String = ""
        Dim Res As String = ""
        Dim tmpsaldo As Integer = 0

        Dim stmp1 As String = ""
        Dim stmp2 As String = ""
        Dim stmp3 As String = ""
        Dim stmp4 As String = ""
        Dim stmp5 As String = ""
        Dim cnt As Integer = 0
        Dim play As Boolean = True
        Dim played As Boolean = False

        If RowCount = 0 Then Exit Sub
        tmpsaldo = saldo


        stmp1 = DataGridView1.Rows(RowCount - 1).Cells(ColNum.eTPRS).Value
        stmp2 = DataGridView1.Rows(RowCount - 1).Cells(ColNum.eTMRS).Value
        stmp3 = DataGridView1.Rows(RowCount - 1).Cells(ColNum.eSRS).Value
        stmp4 = DataGridView1.Rows(RowCount - 1).Cells(ColNum.eRRS).Value

        If stmp1 <> "" Then
            If stmp1.Substring(0, 3) = "Sat" Then
                stmp1 = stmp1.Substring(InStr(stmp1, ")") - 2, 1)
                cnt = cnt + 1
            End If
        End If

        If stmp2 <> "" Then
            If stmp2.Substring(0, 3) = "Sat" Then
                stmp2 = stmp2.Substring(InStr(stmp2, ")") - 2, 1)
                cnt = cnt + 1
            End If
        End If

        If stmp3 <> "" Then
            If stmp3.Substring(0, 3) = "Sat" Then
                stmp3 = stmp3.Substring(InStr(stmp3, ")") - 2, 1)
                cnt = cnt + 1
            End If
        End If

        If stmp4 <> "" Then
            If stmp4.Substring(0, 3) = "Sat" Then
                stmp4 = stmp4.Substring(InStr(stmp4, ")") - 2, 1)
                cnt = cnt + 1
            End If
        End If

        If cnt > 1 Then 'überprüfen ob alle auf das gleich sätzen wollen.
            stmp1 = stmp1 & stmp2 & stmp3 & stmp4

            For k As Integer = 0 To stmp1.Length - 1
                If k = 0 Then
                    stmp5 = stmp1.Substring(0, 1)
                Else
                    If stmp5 <> stmp1.Substring(k, 1) Then
                        play = False
                        Exit For
                    End If
                End If

            Next k
        End If


        SetTo = DataGridView1.Rows(RowCount - 1).Cells(ColNum.eTPRS).Value
        If SetTo <> "" Then
            If SetTo.Substring(0, 3) = "Sat" Then

                If play = False Then
                    DataGridView1.Rows(RowCount - 1).Cells(ColNum.eTPRS).Value = "Kein Satz(S&R)"
                Else
                    If played = False Then
                        played = True
                        If DataGridView1.Rows(RowCount).Cells(ColNum.eTPR).Value = "" Then 'zero Verlust
                            zeroCnt = zeroCnt + 1
                        End If

                        If SetTo.Substring(InStr(SetTo, "("), 1) = DataGridView1.Rows(RowCount).Cells(ColNum.eTPR).Value Then
                            saldo = saldo + 1
                            DataGridView1.Rows(RowCount - 1).Cells(ColNum.eTPRS).Value = SetTo & " +1/" & saldo
                            Maxloss = 0
                        Else
                            saldo = saldo - 1
                            Maxloss = Maxloss + 1
                            If DataGridView1.Rows(RowCount).Cells(ColNum.eTPR).Value = "" Then
                                DataGridView1.Rows(RowCount - 1).Cells(ColNum.eTPRS).Value = SetTo & " -0.5/" & saldo
                            Else
                                DataGridView1.Rows(RowCount - 1).Cells(ColNum.eTPRS).Value = SetTo & " -1/" & saldo
                            End If
                        End If
                        SatzCnt = SatzCnt + 1
                    Else
                        DataGridView1.Rows(RowCount - 1).Cells(ColNum.eTPRS).Value = "bereits gesetzt"
                    End If
                End If

            End If
        End If


        SetTo = DataGridView1.Rows(RowCount - 1).Cells(ColNum.eTMRS).Value
        If SetTo <> "" Then
            If SetTo.Substring(0, 3) = "Sat" Then
                If play = False Then
                    DataGridView1.Rows(RowCount - 1).Cells(ColNum.eTMRS).Value = "Kein Satz(S&R)"
                Else
                    If played = False Then
                        played = True
                        If DataGridView1.Rows(RowCount).Cells(ColNum.eTMR).Value = "" Then 'zero Verlust
                            zeroCnt = zeroCnt + 1
                        End If
                        If SetTo.Substring(InStr(SetTo, "("), 1) = DataGridView1.Rows(RowCount).Cells(ColNum.eTMR).Value Then
                            saldo = saldo + 1
                            DataGridView1.Rows(RowCount - 1).Cells(ColNum.eTMRS).Value = SetTo & " +1/" & saldo                        
                            Maxloss = 0
                        Else
                            saldo = saldo - 1                            
                            Maxloss = Maxloss + 1
                            If DataGridView1.Rows(RowCount).Cells(ColNum.eTMR).Value = "" Then
                                DataGridView1.Rows(RowCount - 1).Cells(ColNum.eTMRS).Value = SetTo & " -0.5/" & saldo
                            Else
                                DataGridView1.Rows(RowCount - 1).Cells(ColNum.eTMRS).Value = SetTo & " -1/" & saldo
                            End If
                        End If
                        SatzCnt = SatzCnt + 1
                    Else
                        DataGridView1.Rows(RowCount - 1).Cells(ColNum.eTMRS).Value = "bereits gesetzt"
                    End If
                End If
            End If
        End If

        SetTo = DataGridView1.Rows(RowCount - 1).Cells(ColNum.eSRS).Value
        If SetTo <> "" Then
            If SetTo.Substring(0, 3) = "Sat" Then
                If play = False Then
                    DataGridView1.Rows(RowCount - 1).Cells(ColNum.eSRS).Value = "Kein Satz(S&R)"
                Else
                    If played = False Then
                        played = True
                        If DataGridView1.Rows(RowCount).Cells(ColNum.eSR).Value = "" Then 'zero Verlust
                            zeroCnt = zeroCnt + 1
                        End If
                        If SetTo.Substring(InStr(SetTo, "("), 1) = DataGridView1.Rows(RowCount).Cells(ColNum.eSR).Value Then
                            saldo = saldo + 1
                            DataGridView1.Rows(RowCount - 1).Cells(ColNum.eSRS).Value = SetTo & " +1/" & saldo
                            Maxloss = 0
                        Else
                            saldo = saldo - 1
                            Maxloss = Maxloss + 1
                            If DataGridView1.Rows(RowCount).Cells(ColNum.eSR).Value = "" Then
                                DataGridView1.Rows(RowCount - 1).Cells(ColNum.eSRS).Value = SetTo & " -0.5/" & saldo
                            Else
                                DataGridView1.Rows(RowCount - 1).Cells(ColNum.eSRS).Value = SetTo & " -1/" & saldo
                            End If
                        End If
                        SatzCnt = SatzCnt + 1
                    Else
                        DataGridView1.Rows(RowCount - 1).Cells(ColNum.eSRS).Value = "bereits gesetzt"
                    End If
                End If
            End If
        End If

        SetTo = DataGridView1.Rows(RowCount - 1).Cells(ColNum.eRRS).Value
        If SetTo <> "" Then
            If SetTo.Substring(0, 3) = "Sat" Then
                If play = False Then
                    DataGridView1.Rows(RowCount - 1).Cells(ColNum.eRRS).Value = "Kein Satz(S&R)"
                Else
                    If played = False Then
                        played = True
                        If DataGridView1.Rows(RowCount).Cells(ColNum.eRR).Value = "" Then 'zero Verlust
                            zeroCnt = zeroCnt + 1
                        End If
                        If SetTo.Substring(InStr(SetTo, "("), 1) = DataGridView1.Rows(RowCount).Cells(ColNum.eRR).Value Then
                            saldo = saldo + 1
                            DataGridView1.Rows(RowCount - 1).Cells(ColNum.eRRS).Value = SetTo & " +1/" & saldo
                            Maxloss = 0
                        Else
                            saldo = saldo - 1                            
                            Maxloss = Maxloss + 1
                            If DataGridView1.Rows(RowCount).Cells(ColNum.eRR).Value = "" Then
                                DataGridView1.Rows(RowCount - 1).Cells(ColNum.eRRS).Value = SetTo & " -0.5/" & saldo
                            Else
                                DataGridView1.Rows(RowCount - 1).Cells(ColNum.eRRS).Value = SetTo & " -1/" & saldo
                            End If
                        End If
                        SatzCnt = SatzCnt + 1
                    Else
                        DataGridView1.Rows(RowCount - 1).Cells(ColNum.eRRS).Value = "bereits gesetzt"
                    End If
                End If
            End If
        End If

        If Maxloss > GesMaxloss Then
            GesMaxloss = Maxloss
            If GesMaxloss = 9 Then
                GesMaxloss = GesMaxloss
            End If
        End If

        If saldo > MaxSaldo Then
            MaxSaldo = saldo
        ElseIf saldo < MinSaldo Then
            MinSaldo = saldo
        End If

    End Sub

    Private Sub CanISetPlus()
        Dim Rap As String
        Dim Ser As String
        Dim Int As String
        Dim SerInt As String = ""
        Dim SerIntCnt As Integer = 0
        Dim sAkt As String = ""
        Dim sLast As String = ""
        Dim coup As String = ""

        If SatzPlus <> "" Then
            Rap = DataGridView1.Rows(RowCount).Cells(ColNum.eTPR).Value
            If Rap <> "" Then Exit Sub

            For I As Integer = 0 To RowCount
                Ser = DataGridView1.Rows(RowCount - I).Cells(ColNum.eTPS).Value
                Int = DataGridView1.Rows(RowCount - I).Cells(ColNum.eTPI).Value

                If Ser <> "" Then
                    SerInt = "S" + SerInt
                    SerIntCnt = SerIntCnt + 1
                End If

                If Int <> "" Then
                    SerInt = "I" + SerInt
                    SerIntCnt = SerIntCnt + 1
                End If

                If SerIntCnt = 3 Then Exit For
            Next

            If SerInt = "SII" Or SerInt = "ISS" Then 'now set?
                sAkt = DataGridView1.Rows(RowCount).Cells(ColNum.eTP).Value
                For I As Integer = 1 To RowCount
                    coup = DataGridView1.Rows(RowCount - I).Cells(ColNum.eCoup).Value
                    If coup <> "00" And coup <> "0" Then
                        sLast = DataGridView1.Rows(RowCount - I).Cells(ColNum.eTP).Value
                        If sAkt <> "" And sLast = "" Then 'And coup <> "0" Then                            
                            RowCount = RowCount + 1
                            InsertNewCoup("1", True)
                            If DataGridView1.Rows(RowCount).Cells(ColNum.eTPR).Value = SatzPlus Then
                                'red                                
                                DataGridView1.Rows(RowCount - 1).Cells(ColNum.eTPRS).Value = "Satz auf (" & SatzPlus & "/R)"
                            Else
                                'black                                
                                DataGridView1.Rows(RowCount - 1).Cells(ColNum.eTPRS).Value = "Satz auf (" & SatzPlus & "/S)"
                            End If
                            DeletelastDummyCoup()

                            SatzPlus = ""
                        End If
                        Exit For
                    Else
                        testCnt = testCnt + 1
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub DeletelastDummyCoup()
        DataGridView1.Rows.Remove(DataGridView1.Rows.Item(RowCount))
        RowCount = RowCount - 1
    End Sub

    Private Sub CanISetMinus()
        Dim Rap As String
        Dim Ser As String
        Dim Int As String
        Dim SerInt As String = ""
        Dim SerIntCnt As Integer = 0
        Dim sAkt As String = ""
        Dim sLast As String = ""
        Dim coup As String = ""

        If SatzMinus <> "" Then
            Rap = DataGridView1.Rows(RowCount).Cells(ColNum.eTMR).Value
            If Rap <> "" Then Exit Sub

            For I As Integer = 0 To RowCount
                Ser = DataGridView1.Rows(RowCount - I).Cells(ColNum.eTMS).Value
                Int = DataGridView1.Rows(RowCount - I).Cells(ColNum.eTMI).Value

                If Ser <> "" Then
                    SerInt = "S" + SerInt
                    SerIntCnt = SerIntCnt + 1
                End If

                If Int <> "" Then
                    SerInt = "I" + SerInt
                    SerIntCnt = SerIntCnt + 1
                End If

                If SerIntCnt = 3 Then Exit For
            Next

            If SerInt = "SII" Or SerInt = "ISS" Then 'now set?
                sAkt = DataGridView1.Rows(RowCount).Cells(ColNum.eTM).Value
                For I As Integer = 1 To RowCount
                    coup = DataGridView1.Rows(RowCount - I).Cells(ColNum.eCoup).Value
                    If coup <> "00" And coup <> "0" Then
                        sLast = DataGridView1.Rows(RowCount - I).Cells(ColNum.eTM).Value
                        If sAkt <> "" And sLast = "" Then 'And coup <> "0" Then
                            RowCount = RowCount + 1
                            InsertNewCoup("1", True)
                            If DataGridView1.Rows(RowCount).Cells(ColNum.eTMR).Value = SatzMinus Then
                                'red
                                DataGridView1.Rows(RowCount - 1).Cells(ColNum.eTMRS).Value = "Satz auf (" & SatzMinus & "/R)"
                            Else
                                'black
                                DataGridView1.Rows(RowCount - 1).Cells(ColNum.eTMRS).Value = "Satz auf (" & SatzMinus & "/S)"
                            End If
                            DeletelastDummyCoup()
                            SatzMinus = ""
                        End If
                        Exit For
                    Else
                        testCnt = testCnt + 1
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub CanISetSchwarz()
        Dim Rap As String
        Dim Ser As String
        Dim Int As String
        Dim SerInt As String = ""
        Dim SerIntCnt As Integer = 0
        Dim sAkt As String = ""
        Dim sLast As String = ""
        Dim coup As String = ""

        If SatzSchwarz <> "" Then
            Rap = DataGridView1.Rows(RowCount).Cells(ColNum.eSR).Value
            If Rap <> "" Then Exit Sub

            For I As Integer = 0 To RowCount
                Ser = DataGridView1.Rows(RowCount - I).Cells(ColNum.eSS).Value
                Int = DataGridView1.Rows(RowCount - I).Cells(ColNum.eSI).Value

                If Ser <> "" Then
                    SerInt = "S" + SerInt
                    SerIntCnt = SerIntCnt + 1
                End If

                If Int <> "" Then
                    SerInt = "I" + SerInt
                    SerIntCnt = SerIntCnt + 1
                End If

                If SerIntCnt = 3 Then Exit For
            Next

            If SerInt = "SII" Or SerInt = "ISS" Then 'now set?
                sAkt = DataGridView1.Rows(RowCount).Cells(ColNum.eS).Value
                For I As Integer = 1 To RowCount
                    coup = DataGridView1.Rows(RowCount - I).Cells(ColNum.eCoup).Value
                    If coup <> "00" And coup <> "0" Then
                        sLast = DataGridView1.Rows(RowCount - I).Cells(ColNum.eS).Value
                        If sAkt <> "" And sLast = "" Then 'And coup <> "0" Then
                            RowCount = RowCount + 1
                            InsertNewCoup("1", True)
                            If DataGridView1.Rows(RowCount).Cells(ColNum.eSR).Value = SatzSchwarz Then
                                'red
                                DataGridView1.Rows(RowCount - 1).Cells(ColNum.eSRS).Value = "Satz auf (" & SatzSchwarz & "/R)"
                            Else
                                'black
                                DataGridView1.Rows(RowCount - 1).Cells(ColNum.eSRS).Value = "Satz auf (" & SatzSchwarz & "/S)"
                            End If
                            DeletelastDummyCoup()
                            SatzSchwarz = ""
                        End If
                        Exit For
                    Else
                        testCnt = testCnt + 1
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub CanISetRot()
        Dim Rap As String
        Dim Ser As String
        Dim Int As String
        Dim SerInt As String = ""
        Dim SerIntCnt As Integer = 0
        Dim sAkt As String = ""
        Dim sLast As String = ""
        Dim coup As String = ""

        If SatzRot <> "" Then
            Rap = DataGridView1.Rows(RowCount).Cells(ColNum.eRR).Value
            If Rap <> "" Then Exit Sub

            For I As Integer = 0 To RowCount
                Ser = DataGridView1.Rows(RowCount - I).Cells(ColNum.eRS).Value
                Int = DataGridView1.Rows(RowCount - I).Cells(ColNum.eRI).Value

                If Ser <> "" Then
                    SerInt = "S" + SerInt
                    SerIntCnt = SerIntCnt + 1
                End If

                If Int <> "" Then
                    SerInt = "I" + SerInt
                    SerIntCnt = SerIntCnt + 1
                End If

                If SerIntCnt = 3 Then Exit For
            Next

            If SerInt = "SII" Or SerInt = "ISS" Then 'now set?
                sAkt = DataGridView1.Rows(RowCount).Cells(ColNum.eR).Value
                For I As Integer = 1 To RowCount
                    coup = DataGridView1.Rows(RowCount - I).Cells(ColNum.eCoup).Value
                    If coup <> "00" And coup <> "0" Then
                        sLast = DataGridView1.Rows(RowCount - I).Cells(ColNum.eR).Value
                        If sAkt <> "" And sLast = "" Then 'And coup <> "0" Then
                            RowCount = RowCount + 1
                            InsertNewCoup("1", True)
                            If DataGridView1.Rows(RowCount).Cells(ColNum.eRR).Value = SatzRot Then
                                'red
                                DataGridView1.Rows(RowCount - 1).Cells(ColNum.eRRS).Value = "Satz auf (" & SatzRot & "/R)"
                            Else
                                'black
                                DataGridView1.Rows(RowCount - 1).Cells(ColNum.eRRS).Value = "Satz auf (" & SatzRot & "/S)"
                            End If
                            DeletelastDummyCoup()
                            SatzRot = ""
                        End If
                        Exit For
                    Else
                        testCnt = testCnt + 1
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub fillSelektorPlus(Optional ByVal fiktiv As Boolean = False)

        Dim Plus1 As String = ""
        Dim Plus2 As String = ""
        Dim Plus3 As String = ""
        Dim PlusTmp As String = ""
        Dim countPlus As Integer = 1
        Dim coup As Integer

        For I As Integer = 0 To RowCount
            coup = DataGridView1.Rows(RowCount - I).Cells(ColNum.eCoup).Value
            If coup <> 0 Then
                PlusTmp = DataGridView1.Rows(RowCount - I).Cells(ColNum.eTP).Value
                Select Case countPlus
                    Case 1
                        Plus1 = PlusTmp
                        countPlus = 2
                    Case 2
                        Plus2 = PlusTmp
                        countPlus = 3
                    Case 3
                        Plus3 = PlusTmp
                        countPlus = 4
                End Select
            End If
            If countPlus = 4 Then Exit For
        Next

        If Plus1 <> "" And Plus2 <> "" And Plus3 = "" Then 'Serie
            DataGridView1.Rows(RowCount).Cells(ColNum.eTPS).Value = "X"
        End If

        If Plus1 = "" And Plus2 <> "" And Plus3 = "" Then 'Intermittenz
            DataGridView1.Rows(RowCount).Cells(ColNum.eTPI).Value = "X"
        End If
    End Sub


    Private Sub fillSelektorSchwarz()

        Dim Schwarz1 As String = ""
        Dim Schwarz2 As String = ""
        Dim Schwarz3 As String = ""
        Dim SchwarzTmp As String = ""
        Dim countSchwarz As Integer = 1
        Dim coup As Integer

        For I As Integer = 0 To RowCount
            coup = DataGridView1.Rows(RowCount - I).Cells(ColNum.eCoup).Value
            If coup <> 0 Then
                SchwarzTmp = DataGridView1.Rows(RowCount - I).Cells(ColNum.eS).Value
                Select Case countSchwarz
                    Case 1
                        Schwarz1 = SchwarzTmp
                        countSchwarz = 2
                    Case 2
                        Schwarz2 = SchwarzTmp
                        countSchwarz = 3
                    Case 3
                        Schwarz3 = SchwarzTmp
                        countSchwarz = 4
                End Select
            End If
            If countSchwarz = 4 Then Exit For
        Next

        If Schwarz1 <> "" And Schwarz2 <> "" And Schwarz3 = "" Then 'Serie
            DataGridView1.Rows(RowCount).Cells(ColNum.eSS).Value = "X"
        End If

        If Schwarz1 = "" And Schwarz2 <> "" And Schwarz3 = "" Then 'Intermittenz
            DataGridView1.Rows(RowCount).Cells(ColNum.eSI).Value = "X"
        End If
    End Sub

    Private Sub fillSelektorRot()

        Dim Rot1 As String = ""
        Dim Rot2 As String = ""
        Dim Rot3 As String = ""
        Dim RotTmp As String = ""
        Dim countRot As Integer = 1
        Dim coup As Integer

        For I As Integer = 0 To RowCount
            coup = DataGridView1.Rows(RowCount - I).Cells(ColNum.eCoup).Value
            If coup <> 0 Then
                RotTmp = DataGridView1.Rows(RowCount - I).Cells(ColNum.eR).Value
                Select Case countRot
                    Case 1
                        Rot1 = RotTmp
                        countRot = 2
                    Case 2
                        Rot2 = RotTmp
                        countRot = 3
                    Case 3
                        Rot3 = RotTmp
                        countRot = 4
                End Select
            End If
            If countRot = 4 Then Exit For
        Next

        If Rot1 <> "" And Rot2 <> "" And Rot3 = "" Then 'Serie
            DataGridView1.Rows(RowCount).Cells(ColNum.eRS).Value = "X"
        End If

        If Rot1 = "" And Rot2 <> "" And Rot3 = "" Then 'Intermittenz
            DataGridView1.Rows(RowCount).Cells(ColNum.eRI).Value = "X"
        End If
    End Sub

    Private Sub fillSelektorPlusRap()

        Dim PlusSerie As String = ""
        Dim PlusIntermit As String = ""
        Dim PlusRap1 As String = ""
        Dim PlusRap2 As String = ""
        Dim PlusRap3 As String = ""
        Dim PlusRap4 As String = ""
        Dim PlusSum As String = ""
        Dim coup As Integer
        Dim Rap As String = ""

        PlusSerie = DataGridView1.Rows(RowCount).Cells(ColNum.eTPS).Value
        PlusIntermit = DataGridView1.Rows(RowCount).Cells(ColNum.eTPI).Value

        If PlusSerie = "" And PlusIntermit = "" Then Exit Sub

        For I As Integer = 0 To RowCount
            coup = DataGridView1.Rows(RowCount - I).Cells(ColNum.eCoup).Value
            If coup <> 0 Then
                PlusSerie = DataGridView1.Rows(RowCount - I).Cells(ColNum.eTPS).Value
                PlusIntermit = DataGridView1.Rows(RowCount - I).Cells(ColNum.eTPI).Value
                If PlusSerie <> "" Then
                    If PlusRap4 = "" Then
                        PlusRap4 = "S"
                    Else
                        If PlusRap3 = "" Then
                            PlusRap3 = "S"
                        Else
                            If PlusRap2 = "" Then
                                PlusRap2 = "S"
                            Else
                                If PlusRap1 = "" Then
                                    PlusRap1 = "S"
                                End If
                            End If
                        End If
                    End If
                ElseIf PlusIntermit <> "" Then
                    If PlusRap4 = "" Then
                        PlusRap4 = "I"
                    Else
                        If PlusRap3 = "" Then
                            PlusRap3 = "I"
                        Else
                            If PlusRap2 = "" Then
                                PlusRap2 = "I"
                            Else
                                If PlusRap1 = "" Then
                                    PlusRap1 = "I"
                                End If
                            End If
                        End If
                    End If
                    If PlusRap1 <> "" And PlusRap2 <> "" And PlusRap3 <> "" Then Exit For
                End If
                If PlusRap1 <> "" And PlusRap2 <> "" And PlusRap3 <> "" And PlusRap4 <> "" Then Exit For
            End If
        Next

        PlusSum = PlusRap1 & PlusRap2 & PlusRap3 & PlusRap4

        Select Case PlusSum
            Case "ISSI"
                Rap = "O"
            Case "SIIS"
                Rap = "O"
            Case "SIII"
                Rap = "X"
            Case "ISSS"
                Rap = "X"
        End Select

        DataGridView1.Rows(RowCount).Cells(ColNum.eTPR).Value = Rap

        If Rap <> "" Then
            DataGridView1.Rows(RowCount).Cells(ColNum.eTPR7).Value = GetRapLast7(ColNum.eTPR)
        End If
    End Sub

    Private Sub fillSelektorSchwarzRap()

        Dim SchwarzSerie As String = ""
        Dim SchwarzIntermit As String = ""
        Dim SchwarzRap1 As String = ""
        Dim SchwarzRap2 As String = ""
        Dim SchwarzRap3 As String = ""
        Dim SchwarzRap4 As String = ""
        Dim SchwarzSum As String = ""
        Dim coup As Integer
        Dim Rap As String = ""

        SchwarzSerie = DataGridView1.Rows(RowCount).Cells(ColNum.eSS).Value
        SchwarzIntermit = DataGridView1.Rows(RowCount).Cells(ColNum.eSI).Value

        If SchwarzSerie = "" And SchwarzIntermit = "" Then Exit Sub

        For I As Integer = 0 To RowCount
            coup = DataGridView1.Rows(RowCount - I).Cells(ColNum.eCoup).Value
            If coup <> 0 Then
                SchwarzSerie = DataGridView1.Rows(RowCount - I).Cells(ColNum.eSS).Value
                SchwarzIntermit = DataGridView1.Rows(RowCount - I).Cells(ColNum.eSI).Value
                If SchwarzSerie <> "" Then
                    If SchwarzRap4 = "" Then
                        SchwarzRap4 = "S"
                    Else
                        If SchwarzRap3 = "" Then
                            SchwarzRap3 = "S"
                        Else
                            If SchwarzRap2 = "" Then
                                SchwarzRap2 = "S"
                            Else
                                If SchwarzRap1 = "" Then
                                    SchwarzRap1 = "S"
                                End If
                            End If
                        End If
                    End If
                ElseIf SchwarzIntermit <> "" Then
                    If SchwarzRap4 = "" Then
                        SchwarzRap4 = "I"
                    Else
                        If SchwarzRap3 = "" Then
                            SchwarzRap3 = "I"
                        Else
                            If SchwarzRap2 = "" Then
                                SchwarzRap2 = "I"
                            Else
                                If SchwarzRap1 = "" Then
                                    SchwarzRap1 = "I"
                                End If
                            End If
                        End If
                    End If
                    If SchwarzRap1 <> "" And SchwarzRap2 <> "" And SchwarzRap3 <> "" Then Exit For
                End If
                If SchwarzRap1 <> "" And SchwarzRap2 <> "" And SchwarzRap3 <> "" And SchwarzRap4 <> "" Then Exit For
            End If
        Next

        SchwarzSum = SchwarzRap1 & SchwarzRap2 & SchwarzRap3 & SchwarzRap4

        Select Case SchwarzSum
            Case "ISSI"
                Rap = "O"
            Case "SIIS"
                Rap = "O"
            Case "SIII"
                Rap = "X"
            Case "ISSS"
                Rap = "X"
        End Select

        DataGridView1.Rows(RowCount).Cells(ColNum.eSR).Value = Rap

        If Rap <> "" Then
            DataGridView1.Rows(RowCount).Cells(ColNum.eSR7).Value = GetRapLast7(ColNum.eSR)
        End If
    End Sub

    Private Sub fillSelektorRotRap()

        Dim RotSerie As String = ""
        Dim RotIntermit As String = ""
        Dim RotRap1 As String = ""
        Dim RotRap2 As String = ""
        Dim RotRap3 As String = ""
        Dim RotRap4 As String = ""
        Dim RotSum As String = ""
        Dim coup As Integer
        Dim Rap As String = ""

        RotSerie = DataGridView1.Rows(RowCount).Cells(ColNum.eRS).Value
        RotIntermit = DataGridView1.Rows(RowCount).Cells(ColNum.eRI).Value

        If RotSerie = "" And RotIntermit = "" Then Exit Sub

        For I As Integer = 0 To RowCount
            coup = DataGridView1.Rows(RowCount - I).Cells(ColNum.eCoup).Value
            If coup <> 0 Then
                RotSerie = DataGridView1.Rows(RowCount - I).Cells(ColNum.eRS).Value
                RotIntermit = DataGridView1.Rows(RowCount - I).Cells(ColNum.eRI).Value
                If RotSerie <> "" Then
                    If RotRap4 = "" Then
                        RotRap4 = "S"
                    Else
                        If RotRap3 = "" Then
                            RotRap3 = "S"
                        Else
                            If RotRap2 = "" Then
                                RotRap2 = "S"
                            Else
                                If RotRap1 = "" Then
                                    RotRap1 = "S"
                                End If
                            End If
                        End If
                    End If
                ElseIf RotIntermit <> "" Then
                    If RotRap4 = "" Then
                        RotRap4 = "I"
                    Else
                        If RotRap3 = "" Then
                            RotRap3 = "I"
                        Else
                            If RotRap2 = "" Then
                                RotRap2 = "I"
                            Else
                                If RotRap1 = "" Then
                                    RotRap1 = "I"
                                End If
                            End If
                        End If
                    End If
                    If RotRap1 <> "" And RotRap2 <> "" And RotRap3 <> "" Then Exit For
                End If
                If RotRap1 <> "" And RotRap2 <> "" And RotRap3 <> "" And RotRap4 <> "" Then Exit For
            End If
        Next

        RotSum = RotRap1 & RotRap2 & RotRap3 & RotRap4

        Select Case RotSum
            Case "ISSI"
                Rap = "O"
            Case "SIIS"
                Rap = "O"
            Case "SIII"
                Rap = "X"
            Case "ISSS"
                Rap = "X"
        End Select

        DataGridView1.Rows(RowCount).Cells(ColNum.eRR).Value = Rap

        If Rap <> "" Then
            DataGridView1.Rows(RowCount).Cells(ColNum.eRR7).Value = GetRapLast7(ColNum.eRR)
        End If
    End Sub

    Private Function GetRapLast7(ByVal SelektorCol As Integer) As String
        Dim Rap As String = ""
        Dim RapTmp As String = ""
        Dim count As Integer = 0

        For I As Integer = 0 To RowCount
            RapTmp = DataGridView1.Rows(RowCount - I).Cells(SelektorCol).Value

            If RapTmp <> "" Then
                count = count + 1
                Rap = RapTmp + Rap
            End If

            If count = 7 Then Exit For
        Next

        If count >= 3 Then
            GetRapLast7 = Rap
        Else
            GetRapLast7 = ""
        End If
    End Function

    Private Function FillSatz(ByVal SelektorCol As Integer) As String
        Dim Rap As String = ""
        Dim RapTmp As String = ""
        Dim count As Integer = 0
        Dim I As Integer = 0

        Dim s1 As String = ""
        Dim s2 As String = ""
        Dim s3 As String = ""
        Dim s4 As String = ""
        Dim s5 As String = ""
        Dim s6 As String = ""
        Dim s7 As String = ""
        Dim ReturnStr As String = ""
        Dim Satz As String = ""



        RapTmp = DataGridView1.Rows(RowCount).Cells(SelektorCol).Value
        If RapTmp <> "" Then
            I = 7
            Do Until I = 0
                Select Case I
                    Case 1
                        s1 = RapTmp.Substring(RapTmp.Length - 7, 1)
                    Case 2
                        s2 = RapTmp.Substring(RapTmp.Length - 6, 1)
                    Case 3
                        s3 = RapTmp.Substring(RapTmp.Length - 5, 1)
                    Case 4
                        s4 = RapTmp.Substring(RapTmp.Length - 4, 1)
                    Case 5
                        s5 = RapTmp.Substring(RapTmp.Length - 3, 1)
                    Case 6
                        s6 = RapTmp.Substring(RapTmp.Length - 2, 1)
                    Case 7
                        s7 = RapTmp.Substring(RapTmp.Length - 1, 1)
                End Select

                I = I - 1

                If 7 - I = RapTmp.Length Then Exit Do
            Loop

            ReturnStr = ""

            'Figuren isoliert: -160
            '    ''Regel10 XXX(X) | OOO(O)
            '    If (s7 = s6 And s7 = s5) Then

            '        If RapTmp.Length < 4 Then
            '            FillSatz = ""
            '            Exit Function
            '        End If

            '        If ((s7 <> s4) Or RapTmp.Length = 4) Then
            '            ReturnStr = "R10: " & s5 & s6 & s7 & "(" & s7 & ")"
            '            SatzCount = SatzCount + 1
            '            Call SetSatz(s7, SelektorCol)
            '        Else
            '            ReturnStr = "R10: Not Aborted"
            '        End If
            '        FillSatz = ReturnStr
            '        Exit Function
            '    End If

            '    ''Is R10 aborted?
            '    If s2 = s3 And s3 = s4 And s4 = s5 Then
            '        FillSatz = ""
            '        Exit Function
            '    End If

            '    ''Regel11 OO XOX(O) | XX OXO(X) 
            '    If (s7 <> s6 And s7 = s5 And s5 <> s4) Then

            '        If RapTmp.Length < 5 Then
            '            FillSatz = ""
            '            Exit Function
            '        End If

            '        If (s6 = s4 And s4 = s3) Then
            '            ReturnStr = "R11: " & s5 & s6 & s7 & "(" & s6 & ")"
            '            SatzCount = SatzCount + 1
            '            Call SetSatz(s6, SelektorCol)
            '        Else
            '            ReturnStr = "R11: Not Aborted"
            '        End If

            '        FillSatz = ReturnStr
            '        Exit Function
            '    End If

            '    'XXOOX
            '    ''Regel12 XXO(O) | OOX(X) XXOOX
            '    If (s7 <> s6 And s6 = s5 And s5 <> s4) Then

            '        If RapTmp.Length < 5 Then
            '            FillSatz = ""
            '            Exit Function
            '        End If

            '        If Not ((s4 = s7) And (s3 = s4)) Then
            '            ReturnStr = "R12: " & s5 & s6 & s7 & "(" & s7 & ")"
            '            SatzCount = SatzCount + 1
            '            Call SetSatz(s7, SelektorCol)
            '        Else
            '            If RapTmp.Length = 5 Then
            '                ReturnStr = "R12: " & s5 & s6 & s7 & "(" & s7 & ")"
            '                SatzCount = SatzCount + 1
            '                Call SetSatz(s7, SelektorCol)
            '            Else
            '                ReturnStr = "R12: Not Aborted"
            '            End If

            '        End If

            '        FillSatz = ReturnStr
            '        Exit Function
            '    End If

            '    If RapTmp.Length = 7 Then
            '        ''If RapTmp.Length = 5 Then
            '        'Regel13 OXXOX(X) | XOOXO(O)
            '        If (s7 <> s6 And s6 <> s5 And s5 = s4 And s4 <> s3 And (s2 <> s3)) Then
            '            ReturnStr = "R13: " & s3 & s4 & s5 & s6 & s7 & "(" & s7 & ")"
            '            SatzCount = SatzCount + 1
            '            FillSatz = ReturnStr
            '            Call SetSatz(s7, SelektorCol)
            '            Exit Function
            '            'XXOXX(O) | OOXOO(X)
            '        ElseIf (s7 = s6 And s6 <> s5 And s5 <> s4 And s4 = s3 And (s2 <> s3)) Then
            '            ReturnStr = "R13: " & s3 & s4 & s5 & s6 & s7 & "(" & s5 & ")"
            '            SatzCount = SatzCount + 1
            '            FillSatz = ReturnStr
            '            Call SetSatz(s5, SelektorCol)
            '            Exit Function
            '        End If
            '    Else
            '        FillSatz = ""
            '        Exit Function
            '    End If

            If RapTmp.Length < 6 Then
                FillSatz = ""
                Exit Function
            End If

            ''Regel10 XXX(X) | OOO(O)
            If (s7 = s6 And s7 = s5) Then
                If (s7 <> s4) Then
                    ReturnStr = "R10: " & s7 & s6 & s5 & "(" & s7 & ")"
                    SatzCount = SatzCount + 1
                    Call SetSatz(s7, SelektorCol)
                Else
                    ReturnStr = "R10: Not Aborted"
                End If
                FillSatz = ReturnStr
                Exit Function
            End If

            ''Is R10 aborted?
            If s3 <> "" And s4 <> "" And s5 <> "" Then
                If s3 = s4 And s4 = s5 Then
                    FillSatz = ""
                    Exit Function
                End If
            End If

            ''Regel11 OO XOX(O) | XX OXO(X) 
            If (s7 <> s6 And s7 = s5 And s5 <> s4) Then
                If s6 = s4 And s4 = s3 Then
                    ReturnStr = "R11: " & s5 & s6 & s7 & "(" & s6 & ")"
                    SatzCount = SatzCount + 1
                    Call SetSatz(s6, SelektorCol)
                Else
                    ReturnStr = "R11: Not Aborted"
                End If

                FillSatz = ReturnStr
                Exit Function
            End If

            ''Regel12 XXO(O) | OOX(X)
            If (s7 <> s6 And s6 = s5) Then
                If Not ((s4 = s7) And (s3 = s4)) Then
                    ReturnStr = "R12: " & s5 & s6 & s7 & "(" & s7 & ")"
                    SatzCount = SatzCount + 1
                    Call SetSatz(s7, SelektorCol)
                Else
                    ReturnStr = "R12: Not Aborted"
                End If

                FillSatz = ReturnStr
                Exit Function
            End If

            ''If RapTmp.Length = 5 Then
            'Regel13 OXXOX(X) | XOOXO(O)
            If (s7 <> s6 And s6 <> s5 And s5 = s4 And s4 <> s3) Then
                ReturnStr = "R13: " & s3 & s4 & s5 & s6 & s7 & "(" & s7 & ")"
                SatzCount = SatzCount + 1
                FillSatz = ReturnStr
                Call SetSatz(s7, SelektorCol)
                Exit Function
                'XXOXX(O) | OOXOO(X)
            ElseIf (s7 = s6 And s6 <> s5 And s5 <> s4 And s4 = s3) Then
                ReturnStr = "R13: " & s3 & s4 & s5 & s6 & s7 & "(" & s5 & ")"
                SatzCount = SatzCount + 1
                FillSatz = ReturnStr
                Call SetSatz(s5, SelektorCol)
                Exit Function
            End If
        End If

        FillSatz = ReturnStr

    End Function

    Private Sub SetSatz(ByVal s As String, ByVal col As Integer)
        If s <> "" Then
            Select Case col
                Case ColNum.eTPR7
                    SatzPlus = s
                Case ColNum.eTMR7
                    SatzMinus = s
                Case ColNum.eSR7
                    SatzSchwarz = s
                Case ColNum.eRR7
                    SatzRot = s
            End Select
        End If
    End Sub

    Private Sub fillSelektorMinusRap()
        Dim MinusSerie As String = ""
        Dim MinusIntermit As String = ""
        Dim MinusRap1 As String = ""
        Dim MinusRap2 As String = ""
        Dim MinusRap3 As String = ""
        Dim MinusRap4 As String = ""
        Dim MinusSum As String = ""
        Dim coup As Integer
        Dim Rap As String = ""

        MinusSerie = DataGridView1.Rows(RowCount).Cells(ColNum.eTMS).Value
        MinusIntermit = DataGridView1.Rows(RowCount).Cells(ColNum.eTMI).Value

        If MinusSerie = "" And MinusIntermit = "" Then Exit Sub

        For I As Integer = 0 To RowCount
            coup = DataGridView1.Rows(RowCount - I).Cells(ColNum.eCoup).Value
            If coup <> 0 Then
                MinusSerie = DataGridView1.Rows(RowCount - I).Cells(ColNum.eTMS).Value
                MinusIntermit = DataGridView1.Rows(RowCount - I).Cells(ColNum.eTMI).Value
                If MinusSerie <> "" Then
                    If MinusRap4 = "" Then
                        MinusRap4 = "S"
                    Else
                        If MinusRap3 = "" Then
                            MinusRap3 = "S"
                        Else
                            If MinusRap2 = "" Then
                                MinusRap2 = "S"
                            Else
                                If MinusRap1 = "" Then
                                    MinusRap1 = "S"
                                End If
                            End If
                        End If
                    End If
                ElseIf MinusIntermit <> "" Then
                    If MinusRap4 = "" Then
                        MinusRap4 = "I"
                    Else
                        If MinusRap3 = "" Then
                            MinusRap3 = "I"
                        Else
                            If MinusRap2 = "" Then
                                MinusRap2 = "I"
                            Else
                                If MinusRap1 = "" Then
                                    MinusRap1 = "I"
                                End If
                            End If
                        End If
                    End If
                    If MinusRap1 <> "" And MinusRap2 <> "" And MinusRap3 <> "" Then Exit For
                End If
                If MinusRap1 <> "" And MinusRap2 <> "" And MinusRap3 <> "" And MinusRap4 <> "" Then Exit For
            End If
        Next

        MinusSum = MinusRap1 & MinusRap2 & MinusRap3 & MinusRap4

        Select Case MinusSum
            Case "ISSI"
                Rap = "O"
            Case "SIIS"
                Rap = "O"
            Case "SIII"
                Rap = "X"
            Case "ISSS"
                Rap = "X"
        End Select

        DataGridView1.Rows(RowCount).Cells(ColNum.eTMR).Value = Rap

        If Rap <> "" Then
            DataGridView1.Rows(RowCount).Cells(ColNum.eTMR7).Value = GetRapLast7(ColNum.eTMR)
        End If

    End Sub

    Private Sub fillSelektorMinus()

        Dim Minus1 As String = ""
        Dim Minus2 As String = ""
        Dim Minus3 As String = ""
        Dim MinusTmp As String = ""
        Dim countMinus As Integer = 1
        Dim coup As Integer

        For I As Integer = 0 To RowCount
            coup = DataGridView1.Rows(RowCount - I).Cells(ColNum.eCoup).Value
            If coup <> 0 Then
                MinusTmp = DataGridView1.Rows(RowCount - I).Cells(ColNum.eTM).Value
                Select Case countMinus
                    Case 1
                        Minus1 = MinusTmp
                        countMinus = 2
                    Case 2
                        Minus2 = MinusTmp
                        countMinus = 3
                    Case 3
                        Minus3 = MinusTmp
                        countMinus = 4
                End Select
            End If
            If countMinus = 4 Then Exit For
        Next

        If Minus1 <> "" And Minus2 <> "" And Minus3 = "" Then 'Serie
            DataGridView1.Rows(RowCount).Cells(ColNum.eTMS).Value = "X"
        End If

        If Minus1 = "" And Minus2 <> "" And Minus3 = "" Then 'Intermittenz
            DataGridView1.Rows(RowCount).Cells(ColNum.eTMI).Value = "X"
        End If
    End Sub

    Private Function GetNextCol(ByRef pos As Integer, Optional ByVal AddCol As Boolean = True, Optional ByRef coup As Integer = 0) As Integer
        Dim i As Integer

        GetNextCol = 0

        coup = DataGridView1.Rows(RowCount - pos).Cells(ColNum.eCoup).Value
        If coup <> 0 Then
            GetNextCol = CheckCollong(coup)
        End If

        If AddCol = True Then pos = i + 1 'damit gleich der nächst coup ausgewählt ist
    End Function

    Private Function GetPlus(ByRef pos As Integer, Optional ByVal AddCol As Boolean = True, Optional ByRef coup As Integer = 0) As String
        Dim i As Integer

        GetPlus = ""

        coup = DataGridView1.Rows(RowCount - pos).Cells(ColNum.eCoup).Value
        If coup <> 0 Then
            GetPlus = DataGridView1.Rows(RowCount - pos).Cells(ColNum.eTM).Value
        End If

        If AddCol = True Then pos = i + 1 'damit gleich der nächst coup ausgewählt ist
    End Function

    Private Function SetTransformatorRegel1() As Integer
        Dim IsCol As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim col1 As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim tmpcol As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer

        Dim IsBreaked As Boolean
        Dim LastcolFound As Boolean = False
        Dim pos As Integer = 1
        Dim coup As Integer
        Dim break As Boolean

        If RowCount < 3 Then 'Regel 1 kann erst fühstens ab Coup 4 eintreten!
            SetTransformatorRegel1 = 0
            Exit Function
        End If

        For I As Integer = pos To RowCount
            tmpcol = GetNextCol(I, False, coup)
            If tmpcol <> 0 Then
                If IsCol = 0 Then
                    IsCol = tmpcol
                Else
                    If col1 = 0 Then
                        col1 = tmpcol
                    End If
                End If
            End If
            If IsCol <> 0 And col1 <> 0 Then Exit For
        Next

        If (IsCol = col1) Then 'hier generelle Abfrage ob Regel2 überhaupt in Kraft treten kann
            'Coup != letzter coup 
            SetTransformatorRegel1 = 0
            Exit Function
        End If

        IsBreaked = CheckIsBreaked()
        For I As Integer = pos To RowCount
            tmpcol = GetNextCol(I, False, coup)
            If LastcolFound = True Then
                If tmpcol <> 0 Then
                    If tmpcol <> IsCol Then
                        'Regel 1: Vorherige Farbe war Intermittenz: 
                        'aktuelle Farbe abgebrochen -> +
                        'aktuelle Farbe weitergelaufen -> -
                        If IsBreaked = True Then
                            SetTransformatorRegel1 = 1 '+
                        Else
                            SetTransformatorRegel1 = 2 '-
                        End If
                        Exit Function
                    Else
                        'Regel 1: nicht in Kraft getreten, Vorherige Farbe war keine Intermittenz
                        SetTransformatorRegel1 = 0
                        Exit Function
                    End If
                End If
            Else
                If tmpcol = IsCol And break = True Then
                    LastcolFound = True
                Else
                    If tmpcol <> 0 And break = False Then
                        If tmpcol <> IsCol Then
                            break = True
                        End If
                    End If
                End If
            End If
        Next I

    End Function

    Private Function SetTransformatorRegel2() As Integer
        Dim IsCol As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim col1 As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim col2 As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim tmpcol As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer

        Dim IsBreaked As Boolean
        Dim LastcolFound As Boolean = False
        Dim pos As Integer = 1
        Dim coup As Integer
        Dim break As Boolean

        If RowCount < 5 Then 'Regel 2 kann erst fühstens ab Coup 6 eintreten!
            SetTransformatorRegel2 = 0
            Exit Function
        End If

        If RowCount = 6 Then
            RowCount = RowCount
        End If

        For I As Integer = pos To RowCount
            tmpcol = GetNextCol(I, False, coup)
            If tmpcol <> 0 Then
                If IsCol = 0 Then
                    IsCol = tmpcol
                Else
                    If col1 = 0 Then
                        col1 = tmpcol
                    Else
                        If col2 = 0 Then
                            col2 = tmpcol
                        End If
                    End If
                End If
            End If
            If IsCol <> 0 And col1 <> 0 And col2 <> 0 Then Exit For
        Next

        If Not ((IsCol = col1) And (col1 <> col2)) Then 'hier generelle Abfrage ob Regel2 überhaupt in Kraft treten kann
            'letzer Coup = vorletzer coup und vorvorletzter coup != letzter coup
            SetTransformatorRegel2 = 0
            Exit Function
        End If

        IsBreaked = CheckIsBreaked()
        For I As Integer = pos To RowCount
            tmpcol = GetNextCol(I, False, coup)
            If LastcolFound = True Then
                If tmpcol <> 0 Then
                    If tmpcol <> IsCol Then
                        'Regel 1: Vorherige Farbe war Intermittenz: 
                        'aktuelle Farbe abgebrochen -> +
                        'aktuelle Farbe weitergelaufen -> -
                        If IsBreaked = True Then
                            SetTransformatorRegel2 = 1 '+
                        Else
                            SetTransformatorRegel2 = 2 '-
                        End If
                        Exit Function
                    Else
                        'Regel 1: nicht in Kraft getreten, Vorherige Farbe war keine Intermittenz
                        SetTransformatorRegel2 = 0
                        Exit Function
                    End If
                End If
            Else
                If tmpcol = IsCol And break = True Then
                    LastcolFound = True
                Else
                    If tmpcol <> 0 And break = False Then
                        If tmpcol <> IsCol Then
                            break = True
                        End If
                    End If
                End If
            End If
        Next I

    End Function

    Private Function SetTransformatorRegel3() As Integer
        Dim IsCol As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim col1 As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim col2 As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim tmpcol As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer

        Dim IsBreaked As Boolean
        Dim LastcolFound As Boolean = False
        Dim pos As Integer = 1
        Dim coup As Integer
        Dim break As Boolean

        If RowCount < 5 Then 'Regel 3 kann erst fühstens ab Coup 6 eintreten!
            SetTransformatorRegel3 = 0
            Exit Function
        End If

        For I As Integer = pos To RowCount
            tmpcol = GetNextCol(I, False, coup)
            If tmpcol <> 0 Then
                If IsCol = 0 Then
                    IsCol = tmpcol
                Else
                    If col1 = 0 Then
                        col1 = tmpcol
                    End If
                End If
            End If
            If IsCol <> 0 And col1 <> 0 Then Exit For
        Next

        If (IsCol = col1) Then 'hier generelle Abfrage ob Regel3 überhaupt in Kraft treten kann
            'Coup != letzter coup 
            SetTransformatorRegel3 = 0
            Exit Function
        End If

        col1 = 0

        IsBreaked = CheckIsBreaked()
        For I As Integer = pos To RowCount
            tmpcol = GetNextCol(I, False, coup)
            If LastcolFound = True Then
                If tmpcol <> 0 Then
                    If col1 = 0 Then
                        col1 = tmpcol
                    Else
                        col2 = tmpcol
                        If ((IsCol = col1) And (col1 <> col2)) Then
                            'Regel 3: Vorherige Farbe war 2er: 
                            'aktuelle Farbe abgebrochen -> -
                            'aktuelle Farbe weitergelaufen -> +
                            If IsBreaked = True Then
                                SetTransformatorRegel3 = 2 '-
                            Else
                                SetTransformatorRegel3 = 1 '+
                            End If
                            Exit Function
                        Else
                            'Regel 1: nicht in Kraft getreten, Vorherige Farbe war keine Intermittenz
                            SetTransformatorRegel3 = 0
                            Exit Function
                        End If
                    End If
                End If
            Else
                If tmpcol = IsCol And break = True Then
                    LastcolFound = True
                Else
                    If tmpcol <> 0 And break = False Then
                        If tmpcol <> IsCol Then
                            break = True
                        End If
                    End If
                End If
            End If
        Next I

    End Function

    Private Function SetTransformatorRegel4() As Integer
        Dim IsCol As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim col1 As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim col2 As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim tmpcol As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer

        Dim IsBreaked As Boolean
        Dim LastcolFound As Boolean = False
        Dim pos As Integer = 1
        Dim coup As Integer
        Dim break As Boolean

        If RowCount < 6 Then 'Regel 4 kann erst fühstens ab Coup 7 eintreten!
            SetTransformatorRegel4 = 0
            Exit Function
        End If

        For I As Integer = pos To RowCount
            tmpcol = GetNextCol(I, False, coup)
            If tmpcol <> 0 Then
                If IsCol = 0 Then
                    IsCol = tmpcol
                Else
                    If col1 = 0 Then
                        col1 = tmpcol
                    Else
                        If col2 = 0 Then
                            col2 = tmpcol
                        End If
                    End If
                End If
            End If
            If IsCol <> 0 And col1 <> 0 And col2 <> 0 Then Exit For
        Next

        If Not ((IsCol = col1) And (col1 <> col2)) Then 'hier generelle Abfrage ob Regel2 überhaupt in Kraft treten kann
            'letzer Coup = vorletzer coup und vorvorletzter coup != letzter coup
            SetTransformatorRegel4 = 0
            Exit Function
        End If

        col1 = 0

        IsBreaked = CheckIsBreaked()
        For I As Integer = pos To RowCount
            tmpcol = GetNextCol(I, False, coup)
            If LastcolFound = True Then
                If tmpcol <> 0 Then
                    If col1 = 0 Then
                        col1 = tmpcol
                    Else
                        col2 = tmpcol
                        If ((IsCol = col1) And (col1 <> col2)) Then
                            'Regel 3: Vorherige Farbe war 2er: 
                            'aktuelle Farbe abgebrochen -> -
                            'aktuelle Farbe weitergelaufen -> +
                            If IsBreaked = True Then
                                SetTransformatorRegel4 = 1 '+
                            Else
                                SetTransformatorRegel4 = 2 '-
                            End If
                            Exit Function
                        Else
                            'Regel 1: nicht in Kraft getreten, Vorherige Farbe war keine Intermittenz
                            SetTransformatorRegel4 = 0
                            Exit Function
                        End If
                    End If
                End If
            Else
                If tmpcol = IsCol And break = True Then
                    LastcolFound = True
                Else
                    If tmpcol <> 0 And break = False Then
                        If tmpcol <> IsCol Then
                            break = True
                        End If
                    End If
                End If
            End If
        Next I

    End Function

    Private Function SetTransformatorRegel5() As Integer
        Dim IsCol As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim col1 As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim col2 As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim tmpcol As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer

        Dim IsBreaked As Boolean
        Dim pos As Integer = 1
        Dim coup As Integer

        If RowCount < 3 Then 'Regel 5 kann erst fühstens ab Coup 4 eintreten!
            SetTransformatorRegel5 = 0
            Exit Function
        End If

        If RowCount = 11 Then
            RowCount = RowCount
        End If

        For I As Integer = pos To RowCount
            tmpcol = GetNextCol(I, False, coup)
            If tmpcol <> 0 Then
                If IsCol = 0 Then
                    IsCol = tmpcol
                Else
                    If col1 = 0 Then
                        col1 = tmpcol
                    Else
                        If col2 = 0 Then
                            col2 = tmpcol
                        End If
                    End If
                End If
            End If
            If IsCol <> 0 And col1 <> 0 And col2 <> 0 Then Exit For
        Next

        If Not ((IsCol = col1) And (col1 = col2)) Then 'hier generelle Abfrage ob Regel5 überhaupt in Kraft treten kann
            'letzer Coup = vorletzer coup = vorvorletzercoup
            SetTransformatorRegel5 = 0
            Exit Function
        End If

        col1 = 0

        IsBreaked = CheckIsBreaked()
        If IsBreaked = True Then
            SetTransformatorRegel5 = 2 '-
        Else
            SetTransformatorRegel5 = 1 '+
        End If

    End Function

    Private Function SetTransformatorRegel5_1() As Integer
        Dim IsCol As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim col1 As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim col2 As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim tmpcol As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer

        Dim IsBreaked As Boolean
        Dim LastcolFound As Boolean = False
        Dim pos As Integer = 1
        Dim coup As Integer
        Dim break As Boolean = False

        If RowCount < 6 Then 'Regel 5_1 kann erst fühstens ab Coup 6 eintreten!
            SetTransformatorRegel5_1 = 0
            Exit Function
        End If

        For I As Integer = pos To RowCount
            tmpcol = GetNextCol(I, False, coup)
            If tmpcol <> 0 Then
                If IsCol = 0 Then
                    IsCol = tmpcol
                Else
                    If col1 = 0 Then
                        col1 = tmpcol
                    End If
                End If
            End If
            If IsCol <> 0 And col1 <> 0 Then Exit For
        Next

        col1 = 0
        IsBreaked = CheckIsBreaked()

        For I As Integer = pos To RowCount
            tmpcol = GetNextCol(I, False, coup)
            If LastcolFound = True Then
                If tmpcol <> 0 Then
                    If col1 = 0 Then
                        col1 = tmpcol
                    Else
                        col2 = tmpcol
                        If ((IsCol = col1) And (col1 = col2)) Then
                            'Regel 5_1: Vorherige Farbe war mindestens 3er: 
                            'aktuelle Farbe abgebrochen -> -
                            'aktuelle Farbe weitergelaufen -> +
                            If IsBreaked = True Then
                                SetTransformatorRegel5_1 = 2 '-
                            Else
                                SetTransformatorRegel5_1 = 1 '+
                            End If
                            Exit Function
                        Else
                            'Regel 1: nicht in Kraft getreten, Vorherige Farbe war keine Intermittenz
                            SetTransformatorRegel5_1 = 0
                            Exit Function
                        End If
                    End If
                End If
            Else
                If tmpcol = IsCol And break = True Then
                    LastcolFound = True
                Else
                    If tmpcol <> 0 And break = False Then
                        If tmpcol <> IsCol Then
                            break = True
                        End If
                    End If
                End If
            End If
        Next I

    End Function

    Private Function CheckIsBreaked()
        If GetNextCol(0) = GetNextCol(1) Then
            CheckIsBreaked = False
        Else
            CheckIsBreaked = True
        End If
    End Function

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

    Private Function CheckCollong(ByVal xNewCoup As Long) As Long 'Prüft die Farbe

        Dim I As Long

        On Error Resume Next

        If xNewCoup = 2 Or xNewCoup = 4 Or xNewCoup = 6 Or xNewCoup = 8 Or xNewCoup = 10 Or xNewCoup = 11 Or xNewCoup = 13 Or xNewCoup = 15 Or xNewCoup = 17 Or xNewCoup = 20 Or _
        xNewCoup = 22 Or xNewCoup = 24 Or xNewCoup = 26 Or xNewCoup = 28 Or xNewCoup = 29 Or xNewCoup = 31 Or xNewCoup = 33 Or xNewCoup = 35 Then
            CheckCollong = 1 '1 für schwarz
        Else
            CheckCollong = -1 '-1 für rot
        End If

        If xNewCoup = 0 Then CheckCollong = 0 '0 für Null

    End Function

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        With OpenFileDialog1
            .Filter = _
            "Text Datei (*.txt)|*.txt|Alle Dateien (*.*)|*.*"
            .FilterIndex = 1
            .InitialDirectory = Application.StartupPath
            .Title = "Auswahl Permanenz Datei"
            .FileName = PermFile
        End With

        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            PermFile = OpenFileDialog1.FileName
            TextBox1.Text = PermFile
        End If
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Auswerten()
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        Auswerten()
    End Sub

    Private Sub TextBox1_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "" Then
            Button2.Enabled = False
        Else
            Button2.Enabled = True
        End If
    End Sub

    Private Sub TextBox3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call InsertManuell()
        End If
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text = "" Then
            Button6.Enabled = False
        Else
            Button6.Enabled = True
        End If
    End Sub

    Private Sub InsertManuell()
        Dim key As String = TextBox3.Text
        If key <> "" Then
            TextBox3.Text = ""
            If CheckKey(key) = False Then
                Call MsgBox(key & " ist keine gültige Roulett Zahl!", MsgBoxStyle.Critical, "Achtung")
                Exit Sub
            Else
                Call InsertNewCoup(key)
                DataGridView1.FirstDisplayedScrollingRowIndex = RowCount - 1
                Me.DataGridView1.Rows(RowCount - 1).Selected = True
                Call SetToolTipStatus()
                Me.Update()
            End If
        End If
    End Sub

    Private Sub SetToolTipStatus()
        'ToolStripStatusLabel1.Text = "Anzahl Coups: " & RowCount & " | Anzahl Satzsignale: " & SatzCount & " (" & FormatNumber((100 * SatzCount) / RowCount, 2) & " %)"
        'ToolStripStatusLabel1.Text = ToolStripStatusLabel1.Text & " | Anz. Sätze: " & SatzCnt & " | Saldo: " & saldo
        'ToolStripStatusLabel1.Text = ToolStripStatusLabel1.Text & " | Zero Verlust: " & zeroCnt
        ToolStripStatusLabel1.Text = ""
    End Sub


    Private Function GetResult() As String
        Dim stxt As String = ""

        stxt = "Anzahl Coups: " & RowCount & vbNewLine
        stxt = stxt & "Anzahl Satzsignale: " & SatzCount & " (" & FormatNumber((100 * SatzCount) / RowCount, 2) & " %)" & vbNewLine
        stxt = stxt & "Anz. Sätze: " & SatzCnt & vbNewLine
        stxt = stxt & "gewonnene Sätze: " & CntGewinn & vbNewLine
        stxt = stxt & "verlorene Sätze: " & CntVerlust & vbNewLine
        stxt = stxt & "Zero Verlust: " & (zeroCnt / 2) & vbNewLine
        stxt = stxt & "trefferchance: " & FormatNumber((CInt(CntGewinn) * 100) / CInt(CntGewinn + CntVerlust), 2) & " %" & vbNewLine

        If saldo > 0 Then
            stxt = stxt & "Saldo: " & saldo & " (+" & FormatNumber(((CInt(saldo) * 100) / (CInt(SatzCnt))), 2) & " %)" & vbNewLine
        Else
            stxt = stxt & "Saldo: " & saldo & " (-" & FormatNumber(((CInt(saldo) * 100) / (CInt(SatzCnt))), 2) & " %)" & vbNewLine
        End If

        stxt = stxt & "größter Saldo: " & MaxSaldo & vbNewLine
        stxt = stxt & "kleinster Saldo: " & MinSaldo & vbNewLine
        stxt = stxt & "größte Verluststrecke: " & GesMaxloss & vbNewLine
        stxt = stxt & vbNewLine

        GetResult = stxt
    End Function

    Private Function CheckKey(ByVal key As String) As Boolean
        CheckKey = False
        If key = "0" Then Return True
        If key = "1" Then Return True
        If key = "2" Then Return True
        If key = "3" Then Return True
        If key = "4" Then Return True
        If key = "5" Then Return True
        If key = "6" Then Return True
        If key = "7" Then Return True
        If key = "8" Then Return True
        If key = "9" Then Return True
        If key = "10" Then Return True
        If key = "11" Then Return True
        If key = "12" Then Return True
        If key = "13" Then Return True
        If key = "14" Then Return True
        If key = "15" Then Return True
        If key = "16" Then Return True
        If key = "17" Then Return True
        If key = "18" Then Return True
        If key = "19" Then Return True
        If key = "20" Then Return True
        If key = "21" Then Return True
        If key = "22" Then Return True
        If key = "23" Then Return True
        If key = "24" Then Return True
        If key = "25" Then Return True
        If key = "26" Then Return True
        If key = "27" Then Return True
        If key = "28" Then Return True
        If key = "29" Then Return True
        If key = "30" Then Return True
        If key = "31" Then Return True
        If key = "32" Then Return True
        If key = "33" Then Return True
        If key = "34" Then Return True
        If key = "35" Then Return True
        If key = "36" Then Return True
    End Function

    Private Sub DataGridView1_ColumnWidthChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles DataGridView1.ColumnWidthChanged
        If loaded = True Then Call SetSizeLabelColPos()
    End Sub

    Private Sub Label8_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label8.DoubleClick
        Dim R1 As DataGridViewRow

        If ColorDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Label8.BackColor = ColorDialog1.Color
            For I As Integer = 0 To RowCount - 1

                R1 = DataGridView1.Rows(I)
                With DataGridView1
                    R1.Cells(ColNum.eRS).Style.BackColor = Label8.BackColor
                    R1.Cells(ColNum.eRI).Style.BackColor = Label8.BackColor
                    R1.Cells(ColNum.eRR).Style.BackColor = Label8.BackColor
                    R1.Cells(ColNum.eRR7).Style.BackColor = Label8.BackColor
                End With
            Next I
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Call InsertManuell()
    End Sub

    Private Sub DataGridView1_Scroll(ByVal sender As Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles DataGridView1.Scroll
        Call SetSizeLabelColPos()
    End Sub

    Private Sub Form1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        DataGridView1.Width = Me.Width - DataGridView1.Left - 30
        DataGridView1.Height = Me.Height - DataGridView1.Top - 70
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        
    End Sub

    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            Auswerten()
        End If
    End Sub

    Private Sub Auswerten()
        ToolStripStatusLabel1.Text = "Daten werden eingelesen und verarbeitet, bitte warten...."
        Me.Update()
        Me.Cursor = Cursors.WaitCursor

        SatzPlus = ""
        SatzMinus = ""
        SatzSchwarz = ""
        SatzRot = ""
        Call StartAuswertung()

        Call SetToolTipStatus()
        'FormRes.prop1() = GetResult()
        FormRes.prop1() = ResTxt
        FormRes.SaldoVerlaufhalbZero = gesSaldoVerlaufhalbZero
        Call FormRes.ShowDialog()
        Me.Update()
    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        DataGridView1.DataSource = "null"
        DataGridView1.DataSource = Nothing
        RowCount = 0

        Me.TextBox1.Focus()

        ToolStripStatusLabel1.Text = "Bitte wählen Sie zum auswerten eine Permanenz Datei aus, oder geben Sie manuell Zahlen ein!"
        Me.Update()
    End Sub

    Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        FormRes.prop1() = ResTxt
        FormRes.SaldoVerlaufhalbZero = gesSaldoVerlaufhalbZero
        Call FormRes.ShowDialog()
    End Sub
End Class
