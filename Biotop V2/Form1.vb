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

    Private SatzPlusRegel As String
    Private SatzMinusRegel As String
    Private SatzSchwarzRegel As String
    Private SatzRotRegel As String

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
    Public arrayPartie As New ArrayList
    Public arrayGlobal As New ArrayList
    Private cCoupData As clsCoupData
    Private Sel1Verlust As Integer
    Private Sel2Verlust As Integer
    Private Sel3Verlust As Integer
    Private Sel4Verlust As Integer
    Private SaldoWithOutZ As Double

    Private zed As ZedGraph.ZedGraphControl
    Private myPane As ZedGraph.GraphPane

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

        ToolStripStatusLabel1.Text = "Bitte Permanenz Datei für die System-Auswertung eingeben!"
        ToolStripProgressBar1.Visible = False
        loaded = True

        Me.Text = "Biotop V2 Version: 2.10"
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
        Dim PartieParam As Integer = 0

        If File.Exists(PermFile) = False Then
            Exit Sub
        End If

        RowCount = 0
        SatzCount = 0
        saldo = 0
        SaldoWithOutZ = 0
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
        arrayPartie.Clear()
        arrayGlobal.Clear()

        Sel1Verlust = 0
        Sel2Verlust = 0
        Sel3Verlust = 0
        Sel4Verlust = 0

        ListBox1.Items.RemoveAt(0)

        Using r As StreamReader = New StreamReader(TextBox1.Text)
            line = r.ReadLine
            Do While (Not line Is Nothing)
                NewCoup = Trim(line)
                If IsNumeric(NewCoup) = True And (Len(NewCoup) = 2 Or Len(NewCoup) = 1) And nextDay = False Then
                    If filterfound = True Then
                        lNewCoup = NewCoup
                        Call InsertNewCoup(NewCoup)

                        If saldo > lastsaldo Then
                            gesSaldoVerlaufhalbZero = gesSaldoVerlaufhalbZero & "+ "
                            saldoverlauf = saldoverlauf & "+ "
                        ElseIf saldo < lastsaldo Then
                            If lNewCoup > 0 Then
                                gesSaldoVerlaufhalbZero = gesSaldoVerlaufhalbZero & "- "
                                saldoverlauf = saldoverlauf & "- "
                            Else
                                saldoverlauf = saldoverlauf & "Z "
                                gesSaldoVerlaufhalbZero = gesSaldoVerlaufhalbZero & "Z "
                                zeroCnt = zeroCnt + 1
                            End If
                        End If

                        If CheckBox1.Checked = False Then
                            If SaldoWithOutZ = 2 And TmpMaxSaldo >= 3 Then
                                nextDay = True
                            ElseIf SaldoWithOutZ = 3 And TmpMaxSaldo >= 5 Then
                                nextDay = True
                            ElseIf SaldoWithOutZ = 5 And TmpMaxSaldo >= 8 Then
                                nextDay = True
                            ElseIf SaldoWithOutZ = 8 And TmpMaxSaldo >= 11 Then
                                nextDay = True
                            ElseIf SaldoWithOutZ = 11 And TmpMaxSaldo >= 14 Then
                                nextDay = True
                            ElseIf SaldoWithOutZ = 14 And TmpMaxSaldo >= 17 Then
                                nextDay = True
                            ElseIf SaldoWithOutZ = 17 And TmpMaxSaldo >= 20 Then
                                nextDay = True
                            ElseIf SaldoWithOutZ = 20 Then
                                nextDay = True
                            ElseIf SaldoWithOutZ = -6 Then
                                nextDay = True
                            End If

                            If TmpMaxSaldo < SaldoWithOutZ Then TmpMaxSaldo = SaldoWithOutZ
                        End If
                        lastsaldo = saldo
                    End If
                Else
                    If InStr(line, "Datum") And CheckBox1.Checked = False Then
                        If datum = "" Then
                            datum = Trim(line.Substring(6))
                        End If

                        If nextDay = True Then
                            GesSaldo = GesSaldo + saldo

                            Me.ListBox1.Items.Add(datum & ": Saldo = " & saldo & vbTab & "Zero Verlust = " & (zeroCnt / 2) & vbTab & "Anz. Sätze = " & SatzCnt & vbTab & "Max. Verlustserie = " & GesMaxloss & vbTab & _
                                                  "Saldoverlauf: " & saldoverlauf)

                            If GlobalGesMaxloss < GesMaxloss Then
                                GlobalGesMaxloss = GesMaxloss
                            End If
                            GesZero = GesZero + (zeroCnt / 2)
                            GesSatzCnt = GesSatzCnt + SatzCnt

                            saldo = 0
                            SaldoWithOutZ = 0
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
                            Sel1Verlust = 0
                            Sel2Verlust = 0
                            Sel3Verlust = 0
                            Sel4Verlust = 0

                            datum = ""
                            nextDay = False
                            TmpMaxSaldo = 0
                            GesMaxloss = 0

                            arrayGlobal.Add(New ArrayList(arrayPartie))
                            arrayPartie.Clear()

                            RowCount = 0

                            datum = Trim(line.Substring(6))
                        End If
                    End If
                End If

                line = r.ReadLine
            Loop

            arrayGlobal.Add(New ArrayList(arrayPartie))


            If datum = "" Then
                If CheckBox1.Checked = True Then
                    datum = "Endlospartie"
                Else
                    datum = Now.Date
                End If
            End If

            If SatzCnt > 0 Then
                GesSaldo = GesSaldo + saldo
                If CheckBox1.Checked = True Then

                    If saldoverlauf.Length > 100 Then
                        Me.ListBox1.Items.Add(datum & ": Saldo = " & saldo & vbTab & "Zero Verlust = " & (zeroCnt / 2) & vbTab & "Anz. Sätze = " & SatzCnt & vbTab & "Max. Verlustserie = " & GesMaxloss & vbTab & _
                                                      "Saldoverlauf: " & saldoverlauf.Substring(0, 100) & "...")
                    Else
                        Me.ListBox1.Items.Add(datum & ": Saldo = " & saldo & vbTab & "Zero Verlust = " & (zeroCnt / 2) & vbTab & "Anz. Sätze = " & SatzCnt & vbTab & "Max. Verlustserie = " & GesMaxloss & vbTab & _
                                                      "Saldoverlauf: " & saldoverlauf)
                    End If
                Else
                    Me.ListBox1.Items.Add(datum & ": Saldo = " & saldo & vbTab & "Zero Verlust = " & (zeroCnt / 2) & vbTab & "Anz. Sätze = " & SatzCnt & vbTab & "Max. Verlustserie = " & GesMaxloss & vbTab & _
                                                  "Saldoverlauf: " & saldoverlauf)
                End If
            Else
                Me.ListBox1.Items.Add("Es konnten keinen Sätze ermittelt werden")
            End If

            GesZero = GesZero + (zeroCnt / 2)
            GesSatzCnt = GesSatzCnt + SatzCnt
            If GlobalGesMaxloss < GesMaxloss Then
                GlobalGesMaxloss = GesMaxloss
            End If

            ResTxt = " Gesamtergebnis:" & vbNewLine
            ResTxt = ResTxt & vbNewLine & " Saldo (mit Zero): " & GesSaldo & vbNewLine
            ResTxt = ResTxt & " Saldo (ohne Zero): " & GesSaldo + GesZero & vbNewLine
            ResTxt = ResTxt & " Zero Verlust: " & GesZero & vbNewLine
            ResTxt = ResTxt & " Gesamt Anzahl Sätze: " & GesSatzCnt & vbNewLine
            ResTxt = ResTxt & " Größte Verlust Serie: " & GlobalGesMaxloss & vbNewLine
            TextRes.Text = ResTxt
        End Using


        ToolStripProgressBar1.Visible = False
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub InsertNewCoup(ByVal NewCoup As String, Optional ByVal dummySet As Boolean = False)

        Dim lNewCoup As Integer
        Dim Col As Integer


        lNewCoup = NewCoup
        Col = CheckCollong(lNewCoup) 'prüft die Farbe

        cCoupData = New clsCoupData()
        cCoupData.Ldf = RowCount + 1
        cCoupData.Coup = NewCoup

        If lNewCoup = 0 Then
            'nichts tun
        Else
            If Col = 1 Then
                cCoupData.S = "S"
            Else
                cCoupData.R = "R"
            End If

            Call SetTransformator()
            Call fillSelektorPlus()
            Call fillSelektorMinus()
            Call fillSelektorPlusRap()
            Call fillSelektorMinusRap()
            Call fillSelektorSchwarz()
            Call fillSelektorRot()
            Call fillSelektorSchwarzRap()
            Call fillSelektorRotRap()

            If dummySet = True Then Exit Sub
            cCoupData.TPRS = FillSatz(ColNum.eTPR7)
            cCoupData.TmRS = FillSatz(ColNum.eTMR7)
            cCoupData.SRS = FillSatz(ColNum.eSR7)
            cCoupData.RRS = FillSatz(ColNum.eRR7)
        End If

        If dummySet = True Then Exit Sub

        Call CanISetPlus()
        Call CanISetMinus()
        Call CanISetSchwarz()
        Call CanISetRot()
        Call CalkSaldo()

        RowCount = RowCount + 1

        arrayPartie.Add(cCoupData)
    End Sub

    Private Sub SetTransformator()

        Dim Erg1 As Integer = SetTransformatorRegel1()
        If Erg1 = 1 Then
            cCoupData.TP = "+ (R1)"
        ElseIf Erg1 = 2 Then
            cCoupData.TM = "- (R1)"
        End If
        If Erg1 = 0 Then
            Erg1 = SetTransformatorRegel2()
            If Erg1 = 1 Then
                cCoupData.TP = "+ (R2)"
            ElseIf Erg1 = 2 Then
                cCoupData.TM = "- (R2)"
            End If
        End If
        If Erg1 = 0 Then
            Erg1 = SetTransformatorRegel3()
            If Erg1 = 1 Then
                cCoupData.TP = "+ (R3)"
            ElseIf Erg1 = 2 Then
                cCoupData.TM = "- (R3)"
            End If
        End If
        If Erg1 = 0 Then
            Erg1 = SetTransformatorRegel4()
            If Erg1 = 1 Then
                cCoupData.TP = "+ (R4)"
            ElseIf Erg1 = 2 Then
                cCoupData.TM = "- (R4)"
            End If
        End If
        If Erg1 = 0 And RowCount > 3 Then
            Erg1 = SetTransformatorRegel5()
            If Erg1 = 1 Then
                cCoupData.TP = "+ (R5)"
            ElseIf Erg1 = 2 Then
                cCoupData.TM = "- (R5)"
            End If
        End If
        If Erg1 = 0 Then
            Erg1 = SetTransformatorRegel5_1()
            If Erg1 = 1 Then
                cCoupData.TP = "+ (R5)"
            ElseIf Erg1 = 2 Then
                cCoupData.TM = "- (R5)"
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
        Dim stmpges As String = ""
        Dim cnt As Integer = 0
        Dim play As Boolean = True
        Dim played As Boolean = False
        Dim arraylistCnt As Integer = arrayPartie.Count

        If CheckBox2.Checked = False Then
            Sel1Verlust = 0
            Sel2Verlust = 0
            Sel3Verlust = 0
            Sel4Verlust = 0
        End If

        If arrayPartie.Count = 0 Then Exit Sub
        Dim localInstanceofMyClass = CType(arrayPartie(arraylistCnt - 1), clsCoupData)
        tmpsaldo = saldo

        If localInstanceofMyClass.Ldf = 128 Then arraylistCnt = arraylistCnt

        stmp1 = localInstanceofMyClass.TPRS
        stmp2 = localInstanceofMyClass.TmRS
        stmp3 = localInstanceofMyClass.SRS
        stmp4 = localInstanceofMyClass.RRS

        If stmp1 <> "" Then
            If stmp1.Substring(0, 3) = "Sat" Then
                stmp1 = stmp1.Substring(InStr(stmp1, "/") - 2, 1)
                cnt = cnt + 1
            End If
        End If

        If stmp2 <> "" Then
            If stmp2.Substring(0, 3) = "Sat" Then
                stmp2 = stmp2.Substring(InStr(stmp2, "/") - 2, 1)
                cnt = cnt + 1
            End If
        End If

        If stmp3 <> "" Then
            If stmp3.Substring(0, 3) = "Sat" Then
                stmp3 = stmp3.Substring(InStr(stmp3, "/") - 2, 1)
                cnt = cnt + 1
            End If
        End If

        If stmp4 <> "" Then
            If stmp4.Substring(0, 3) = "Sat" Then
                stmp4 = stmp4.Substring(InStr(stmp4, "/") - 2, 1)
                cnt = cnt + 1
            End If
        End If

        If cnt > 1 Then 'überprüfen ob alle auf das gleich sätzen wollen.
            stmpges = stmp1 & stmp2 & stmp3 & stmp4

            For k As Integer = 0 To stmpges.Length - 1
                If k = 0 Then
                    stmp5 = stmpges.Substring(0, 1)
                Else
                    If stmp5 <> stmpges.Substring(k, 1) Then
                        play = False
                        Exit For
                    End If
                End If

            Next k
        End If

        SetTo = localInstanceofMyClass.TPRS
        If SetTo <> "" Then
            If SetTo.Substring(0, 3) = "Sat" Then
                If play = False Then
                    localInstanceofMyClass.TPRS = "Kein Satz(S&R)"
                Else
                    If played = False Then
                        played = True

                        If SetTo.Substring(InStr(SetTo, "/") - 2, 1) = cCoupData.TPR Then
                            SatzCnt = SatzCnt + 1
                            If Sel1Verlust < 3 Then
                                saldo = saldo + 1
                                SaldoWithOutZ = SaldoWithOutZ + 1
                                localInstanceofMyClass.TPRS = SetTo & " +1/" & saldo
                                Maxloss = 0
                            Else
                                localInstanceofMyClass.TPRS = SetTo & " Verl. > 3"
                            End If

                            Sel1Verlust = 0
                        Else
                            If Sel1Verlust < 3 Then
                                SatzCnt = SatzCnt + 1
                                Maxloss = Maxloss + 1
                                If cCoupData.TPR = "" Then
                                    saldo = saldo - 0.5
                                    localInstanceofMyClass.TPRS = SetTo & " -0.5/" & saldo
                                Else
                                    saldo = saldo - 1
                                    SaldoWithOutZ = SaldoWithOutZ - 1
                                    localInstanceofMyClass.TPRS = SetTo & " -1/" & saldo
                                End If
                            Else
                                localInstanceofMyClass.TPRS = SetTo & " Verl. > 3"
                            End If

                            Sel1Verlust = Sel1Verlust + 1
                        End If
                    Else
                        localInstanceofMyClass.TPRS = "bereits gesetzt"
                    End If
                End If
            End If
        End If

        SetTo = localInstanceofMyClass.TmRS
        If SetTo <> "" Then
            If SetTo.Substring(0, 3) = "Sat" Then
                If play = False Then
                    localInstanceofMyClass.TmRS = "Kein Satz(S&R)"
                Else
                    If played = False Then
                        played = True

                        If SetTo.Substring(InStr(SetTo, "/") - 2, 1) = cCoupData.TmR Then
                            SatzCnt = SatzCnt + 1
                            If Sel2Verlust < 3 Then
                                saldo = saldo + 1
                                SaldoWithOutZ = SaldoWithOutZ + 1
                                localInstanceofMyClass.TmRS = SetTo & " +1/" & saldo
                                Maxloss = 0
                            Else
                                localInstanceofMyClass.TmRS = SetTo & " Verl. > 3"
                            End If

                            Sel2Verlust = 0
                        Else
                            If Sel2Verlust < 3 Then
                                SatzCnt = SatzCnt + 1
                                Maxloss = Maxloss + 1
                                If cCoupData.TmR = "" Then
                                    saldo = saldo - 0.5
                                    localInstanceofMyClass.TmRS = SetTo & " -0.5/" & saldo
                                Else
                                    saldo = saldo - 1
                                    SaldoWithOutZ = SaldoWithOutZ - 1
                                    localInstanceofMyClass.TmRS = SetTo & " -1/" & saldo
                                End If
                            Else
                                localInstanceofMyClass.TmRS = SetTo & " Verl. > 3"
                            End If

                            Sel2Verlust = Sel2Verlust + 1
                        End If
                    Else
                        localInstanceofMyClass.TmRS = "bereits gesetzt"
                    End If
                End If
            End If
        End If

        SetTo = localInstanceofMyClass.SRS
        If SetTo <> "" Then
            If SetTo.Substring(0, 3) = "Sat" Then
                If play = False Then
                    localInstanceofMyClass.SRS = "Kein Satz(S&R)"
                Else
                    If played = False Then
                        played = True

                        If SetTo.Substring(InStr(SetTo, "/") - 2, 1) = cCoupData.SR Then
                            SatzCnt = SatzCnt + 1
                            If Sel3Verlust < 3 Then
                                saldo = saldo + 1
                                SaldoWithOutZ = SaldoWithOutZ + 1
                                localInstanceofMyClass.SRS = SetTo & " +1/" & saldo
                                Maxloss = 0
                            Else
                                localInstanceofMyClass.SRS = SetTo & " Verl. > 3"
                            End If

                            Sel3Verlust = 0
                        Else
                            If Sel3Verlust < 3 Then
                                SatzCnt = SatzCnt + 1
                                Maxloss = Maxloss + 1
                                If cCoupData.SR = "" Then
                                    saldo = saldo - 0.5
                                    localInstanceofMyClass.SRS = SetTo & " -0.5/" & saldo
                                Else
                                    saldo = saldo - 1
                                    SaldoWithOutZ = SaldoWithOutZ - 1
                                    localInstanceofMyClass.SRS = SetTo & " -1/" & saldo
                                End If
                            Else
                                localInstanceofMyClass.SRS = SetTo & " Verl. > 3"
                            End If

                            Sel3Verlust = Sel3Verlust + 1
                        End If
                    Else
                        localInstanceofMyClass.SRS = "bereits gesetzt"
                    End If
                End If
            End If
        End If

        SetTo = localInstanceofMyClass.RRS
        If SetTo <> "" Then
            If SetTo.Substring(0, 3) = "Sat" Then
                If play = False Then
                    localInstanceofMyClass.RRS = "Kein Satz(S&R)"
                Else
                    If played = False Then
                        played = True

                        If SetTo.Substring(InStr(SetTo, "/") - 2, 1) = cCoupData.RR Then
                            SatzCnt = SatzCnt + 1
                            If Sel4Verlust < 3 Then
                                saldo = saldo + 1
                                SaldoWithOutZ = SaldoWithOutZ + 1
                                localInstanceofMyClass.RRS = SetTo & " +1/" & saldo
                                Maxloss = 0
                            Else
                                localInstanceofMyClass.RRS = SetTo & " Verl. > 3"
                            End If

                            Sel4Verlust = 0
                        Else
                            If Sel4Verlust < 3 Then
                                SatzCnt = SatzCnt + 1
                                Maxloss = Maxloss + 1
                                If cCoupData.RR = "" Then
                                    saldo = saldo - 0.5
                                    localInstanceofMyClass.RRS = SetTo & " -0.5/" & saldo
                                Else
                                    saldo = saldo - 1
                                    SaldoWithOutZ = SaldoWithOutZ - 1
                                    localInstanceofMyClass.RRS = SetTo & " -1/" & saldo
                                End If
                            Else
                                localInstanceofMyClass.RRS = SetTo & " Verl. > 3"
                            End If

                            Sel4Verlust = Sel4Verlust + 1
                        End If
                    Else
                        localInstanceofMyClass.RRS = "bereits gesetzt"
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
        Dim arraylistCnt As Integer = arrayPartie.Count
        Dim localInstanceofMyClass As clsCoupData
        Dim Strang As String = ""        
        Dim StrangRap As String = ""
        Dim StrangRapCnt As Integer = 0
        Dim firstRap As Boolean = True
        Dim LastIsNotPlayingRange As Boolean = False
        Dim tmpSatz As String = ""
        Dim StrangRap2 As String = ""
        Dim nextRap As Boolean = False

        If SatzPlus <> "" Then
            Rap = cCoupData.TPR
            If Rap <> "" Then Exit Sub

            Ser = cCoupData.TPS
            Int = cCoupData.TPI

            If Ser <> "" Then
                SerInt = "S" + SerInt
                SerIntCnt = SerIntCnt + 1
            End If

            If Int <> "" Then
                SerInt = "I" + SerInt
                SerIntCnt = SerIntCnt + 1
            End If

            For I As Integer = 0 To arrayPartie.Count - 1
                localInstanceofMyClass = CType(arrayPartie(arraylistCnt - I - 1), clsCoupData)
                Ser = localInstanceofMyClass.TPS
                Int = localInstanceofMyClass.TPI

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
                sAkt = cCoupData.TP
                For I As Integer = 1 To arraylistCnt - 1
                    localInstanceofMyClass = CType(arrayPartie(arraylistCnt - I), clsCoupData)
                    coup = localInstanceofMyClass.Coup
                    sLast = localInstanceofMyClass.TP

                    If coup <> "00" And coup <> "0" Then
                        If sAkt <> "" And sLast = "" Then 'And coup <> "0" Then                            
                            Dim tmpCopyStruct As clsCoupData = cCoupData
                            arrayPartie.Add(cCoupData) 'momentaner Datensatz in array
                            InsertNewCoup("1", True)

                            If cCoupData.TPR = SatzPlus Then
                                tmpCopyStruct.TPRS = "Satz auf " & SatzPlus & "/R"
                                tmpSatz = "Satz auf " & SatzPlus & "/R"
                            Else
                                tmpCopyStruct.TPRS = "Satz auf " & SatzPlus & "/S"
                                tmpSatz = "Satz auf " & SatzPlus & "/S"
                            End If

                            Ser = cCoupData.TPS
                            Int = cCoupData.TPI

                            If SerInt = "ISS" Then
                                'Serienstrang
                                Strang = "Serie"
                            ElseIf SerInt = "SII" Then
                                'Intermittenztstrang
                                Strang = "Int"                            
                            End If

                            If CheckBox7.Checked = True Or CheckBox8.Checked = True Or CheckBox9.Checked = True Or CheckBox10.Checked = True Or CheckBox11.Checked = True Then
                                For Z As Integer = 2 To arraylistCnt - 1
                                    localInstanceofMyClass = CType(arrayPartie(arraylistCnt - Z), clsCoupData)
                                    Rap = localInstanceofMyClass.TPR

                                    If Rap <> "" Then
                                        nextRap = False
                                        If Strang = "Serie" Then
                                            If Rap = "X" Then
                                                If localInstanceofMyClass.TPS <> "" Then
                                                    StrangRap = StrangRap & "X"
                                                    StrangRapCnt = StrangRapCnt + 1
                                                    nextRap = True
                                                Else
                                                    If firstRap = True Then
                                                        LastIsNotPlayingRange = True
                                                        firstRap = False
                                                    End If
                                                End If
                                            ElseIf Rap = "O" Then
                                                If localInstanceofMyClass.TPI <> "" Then
                                                    StrangRap = StrangRap & "O"
                                                    StrangRapCnt = StrangRapCnt + 1
                                                    nextRap = True
                                                Else
                                                    If firstRap = True Then
                                                        LastIsNotPlayingRange = True
                                                        firstRap = False
                                                    End If
                                                End If
                                            End If
                                        ElseIf Strang = "Int" Then
                                            If Rap = "X" Then
                                                If localInstanceofMyClass.TPI <> "" Then
                                                    StrangRap = StrangRap & "X"
                                                    StrangRapCnt = StrangRapCnt + 1
                                                    nextRap = True
                                                Else
                                                    If firstRap = True Then
                                                        LastIsNotPlayingRange = True
                                                        firstRap = False
                                                    End If
                                                End If
                                            ElseIf Rap = "O" Then
                                                If localInstanceofMyClass.TPS <> "" Then
                                                    StrangRap = StrangRap & "O"
                                                    StrangRapCnt = StrangRapCnt + 1
                                                    nextRap = True
                                                Else
                                                    If firstRap = True Then
                                                        LastIsNotPlayingRange = True
                                                        firstRap = False
                                                    End If
                                                End If
                                            End If
                                        End If
                                        firstRap = False

                                        If nextRap = True Then
                                            If SatzPlusRegel = "R11" And CheckBox7.Checked = True Then 'Check R14'
                                                If StrangRapCnt < 3 Then
                                                    tmpCopyStruct.TPRS = "KS: R14"

                                                ElseIf StrangRapCnt = 3 Then
                                                    If SatzPlus = "X" Then
                                                        If (StrangRap(0) = "O" And StrangRap(1) = "X") Then
                                                            tmpCopyStruct.TPRS = tmpSatz
                                                            Exit For
                                                        End If
                                                    ElseIf SatzPlus = "O" Then
                                                        If (StrangRap(0) = "X" And StrangRap(1) = "O") Then
                                                            tmpCopyStruct.TPRS = tmpSatz
                                                            Exit For
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If CheckBox8.Checked = True And LastIsNotPlayingRange = True And StrangRap.Length >= 3 Then 'Check R15
                                                    If StrangRap.Substring(0, 2) = "XO" Then 'R15= 1er = jetzt prüfen ob danach 2er
                                                        If StrangRap.Substring(StrangRap.Length - 1, 1) = "X" And StrangRap2 = "" Then
                                                            StrangRap2 = "X"
                                                        ElseIf StrangRap2 <> "" Then
                                                            If (StrangRap2 & StrangRap.Substring(StrangRap.Length - 1, 1)) = "XX" And SatzPlus = "O" Then
                                                                tmpCopyStruct.TPRS = "KS: R15"
                                                                Exit For
                                                            ElseIf (StrangRap2 & StrangRap.Substring(StrangRap.Length - 1, 1)) = "XO" And SatzPlus = "X" Then
                                                                tmpCopyStruct.TPRS = "KS: R15"
                                                                Exit For
                                                            End If
                                                        End If
                                                    ElseIf StrangRap.Substring(0, 2) = "OX" Then
                                                        If StrangRap.Substring(StrangRap.Length - 1, 1) = "O" And StrangRap2 = "" Then
                                                            StrangRap2 = "O"
                                                        ElseIf StrangRap2 <> "" Then
                                                            If (StrangRap2 & StrangRap.Substring(StrangRap.Length - 1, 1)) = "OO" And SatzPlus = "X" Then
                                                                tmpCopyStruct.TPRS = "KS: R15"
                                                                Exit For
                                                            ElseIf (StrangRap2 & StrangRap.Substring(StrangRap.Length - 1, 1)) = "OX" And SatzPlus = "O" Then
                                                                tmpCopyStruct.TPRS = "KS: R15"
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                End If

                                                If CheckBox9.Checked = True Then 'Check R16
                                                    If StrangRap.Length = 3 Then
                                                        If StrangRap = "XXO" Then
                                                            If SatzPlus = "O" Then
                                                                tmpCopyStruct.TPRS = "KS: R16"
                                                            End If
                                                            Exit For
                                                        ElseIf StrangRap = "OOX" Then
                                                            If SatzPlus = "X" Then
                                                                tmpCopyStruct.TPRS = "KS: R16"
                                                            End If
                                                            Exit For
                                                        End If
                                                    ElseIf StrangRap.Length = 5 Then
                                                        If StrangRap = "XXXXO" Then
                                                            If SatzPlus = "O" Then
                                                                tmpCopyStruct.TPRS = "KS: R16"
                                                            End If
                                                            Exit For
                                                        ElseIf StrangRap = "OOOOX" Then
                                                            If SatzPlus = "X" Then
                                                                tmpCopyStruct.TPRS = "KS: R16"
                                                            End If
                                                            Exit For
                                                        End If
                                                    End If
                                                End If

                                                If CheckBox10.Checked = True Then 'Check R17
                                                    If StrangRap.Length = 4 Then
                                                        If StrangRap = "XXXO" Then
                                                            If SatzPlus = "X" Then
                                                                tmpCopyStruct.TPRS = "KS: R17"
                                                            End If
                                                            Exit For
                                                        ElseIf StrangRap = "OOOX" Then
                                                            If SatzPlus = "O" Then
                                                                tmpCopyStruct.TPRS = "KS: R17"
                                                            End If
                                                            Exit For
                                                        End If
                                                    End If
                                                End If

                                                If CheckBox11.Checked = True Then 'Check R18
                                                    If StrangRap.Length = 5 Then
                                                        If StrangRap = "XXXXX" Or StrangRap = "OOOOO" Then
                                                            tmpCopyStruct.TPRS = "KS: R18"
                                                            Exit For
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Next Z

                                If CheckBox7.Checked = True And SatzPlusRegel = "R11" Then
                                    If StrangRap.Length < 3 Then
                                        tmpCopyStruct.TPRS = "Vl. zu klein"
                                    End If
                                End If

                                If CheckBox8.Checked = True And LastIsNotPlayingRange = True Then 'R15 fehlt Vorlauf
                                    If StrangRap.Length = 1 Then
                                        If tmpCopyStruct.TPRS <> "KS: R15" Then
                                            tmpCopyStruct.TPRS = "Vl. zu klein"
                                        End If
                                    ElseIf StrangRap.Length >= 2 Then
                                        If StrangRap.Substring(0, 2) = "OX" Or StrangRap.Substring(0, 2) = "XO" Then
                                            If tmpCopyStruct.TPRS <> "KS: R15" Then
                                                tmpCopyStruct.TPRS = "Vl. zu klein"
                                            End If
                                        End If
                                    Else
                                        tmpCopyStruct.TPRS = "Vl. zu klein"
                                    End If
                                End If

                                If CheckBox9.Checked = True Then
                                    If StrangRap = "X" Or StrangRap = "XX" Or StrangRap = "XXX" Or StrangRap = "XXXX" Then
                                        If SatzPlus = "O" Then
                                            tmpCopyStruct.TPRS = "Vl. zu klein"
                                        End If
                                    ElseIf StrangRap = "O" Or StrangRap = "OO" Or StrangRap = "OOO" Or StrangRap = "OOO" Then
                                        If SatzPlus = "O" Then
                                            tmpCopyStruct.TPRS = "Vl. zu klein"
                                        End If
                                    End If
                                End If

                                If CheckBox10.Checked = True Then
                                    If StrangRap = "X" Or StrangRap = "XX" Or StrangRap = "XXX" Then
                                        If SatzPlus = "X" Then
                                            tmpCopyStruct.TPRS = "Vl. zu klein"
                                        End If
                                    ElseIf StrangRap = "O" Or StrangRap = "OO" Or StrangRap = "OOO" Then
                                        If SatzPlus = "O" Then
                                            tmpCopyStruct.TPRS = "Vl. zu klein"
                                        End If
                                    End If
                                End If

                                If CheckBox11.Checked = True Then
                                    If StrangRap = "X" Or StrangRap = "XX" Or StrangRap = "XXX" Or StrangRap = "XXXX" Or StrangRap = "O" Or StrangRap = "OO" Or StrangRap = "OOO" Or StrangRap = "OOOO" Then
                                        tmpCopyStruct.TPRS = "Vl. zu klein"
                                    End If
                                End If
                            End If


                            DeletelastDummyCoup()
                            ''zurückschreiben
                            cCoupData = tmpCopyStruct
                            SatzPlus = ""
                            SatzPlusRegel = ""
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
        Dim ArraylistCnt As Integer = arrayPartie.Count
        arrayPartie.Remove(arrayPartie(ArraylistCnt - 1))
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
        Dim arraylistCnt As Integer = arrayPartie.Count
        Dim localInstanceofMyClass As clsCoupData

        Dim Strang As String = ""
        Dim StrangRap As String = ""
        Dim StrangRapCnt As Integer = 0
        Dim firstRap As Boolean = True
        Dim LastIsNotPlayingRange As Boolean = False
        Dim tmpSatz As String = ""
        Dim StrangRap2 As String = ""
        Dim nextRap As Boolean = False

        If SatzMinus <> "" Then
            Rap = cCoupData.TmR
            If Rap <> "" Then Exit Sub

            Ser = cCoupData.TmS
            Int = cCoupData.TmI

            If Ser <> "" Then
                SerInt = "S" + SerInt
                SerIntCnt = SerIntCnt + 1
            End If

            If Int <> "" Then
                SerInt = "I" + SerInt
                SerIntCnt = SerIntCnt + 1
            End If

            For I As Integer = 0 To arrayPartie.Count - 1
                localInstanceofMyClass = CType(arrayPartie(arraylistCnt - I - 1), clsCoupData)
                Ser = localInstanceofMyClass.TmS
                Int = localInstanceofMyClass.TmI

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
                sAkt = cCoupData.TM
                For I As Integer = 1 To arraylistCnt - 1
                    localInstanceofMyClass = CType(arrayPartie(arraylistCnt - I), clsCoupData)
                    coup = localInstanceofMyClass.Coup
                    sLast = localInstanceofMyClass.TM

                    If coup <> "00" And coup <> "0" Then
                        If sAkt <> "" And sLast = "" Then 'And coup <> "0" Then                            
                            Dim tmpCopyStruct As clsCoupData = cCoupData
                            arrayPartie.Add(cCoupData) 'momentaner Datensatz in array
                            InsertNewCoup("1", True)
                            If cCoupData.TmR = SatzMinus Then
                                tmpCopyStruct.TmRS = "Satz auf " & SatzMinus & "/R"
                                tmpSatz = "Satz auf " & SatzMinus & "/R"
                            Else
                                tmpCopyStruct.TmRS = "Satz auf " & SatzMinus & "/S"
                                tmpSatz = "Satz auf " & SatzMinus & "/S"
                            End If

                            Ser = cCoupData.TmS
                            Int = cCoupData.TmI

                            If SerInt = "ISS" Then
                                'Serienstrang
                                Strang = "Serie"
                            ElseIf SerInt = "SII" Then
                                'Intermittenztstrang
                                Strang = "Int"
                            End If

                            If CheckBox7.Checked = True Or CheckBox8.Checked = True Or CheckBox9.Checked = True Or CheckBox10.Checked = True Or CheckBox11.Checked = True Then
                                For Z As Integer = 2 To arraylistCnt - 1
                                    localInstanceofMyClass = CType(arrayPartie(arraylistCnt - Z), clsCoupData)
                                    Rap = localInstanceofMyClass.TmR

                                    If Rap <> "" Then
                                        nextRap = False
                                        If Strang = "Serie" Then
                                            If Rap = "X" Then
                                                If localInstanceofMyClass.TmS <> "" Then
                                                    StrangRap = StrangRap & "X"
                                                    StrangRapCnt = StrangRapCnt + 1
                                                    nextRap = True
                                                Else
                                                    If firstRap = True Then
                                                        LastIsNotPlayingRange = True
                                                        firstRap = False
                                                    End If
                                                End If
                                            ElseIf Rap = "O" Then
                                                If localInstanceofMyClass.TmI <> "" Then
                                                    StrangRap = StrangRap & "O"
                                                    StrangRapCnt = StrangRapCnt + 1
                                                    nextRap = True
                                                Else
                                                    If firstRap = True Then
                                                        LastIsNotPlayingRange = True
                                                        firstRap = False
                                                    End If
                                                End If
                                            End If
                                        ElseIf Strang = "Int" Then
                                            If Rap = "X" Then
                                                If localInstanceofMyClass.TmI <> "" Then
                                                    StrangRap = StrangRap & "X"
                                                    StrangRapCnt = StrangRapCnt + 1
                                                    nextRap = True
                                                Else
                                                    If firstRap = True Then
                                                        LastIsNotPlayingRange = True
                                                        firstRap = False
                                                    End If
                                                End If
                                            ElseIf Rap = "O" Then
                                                If localInstanceofMyClass.TmS <> "" Then
                                                    StrangRap = StrangRap & "O"
                                                    StrangRapCnt = StrangRapCnt + 1
                                                    nextRap = True
                                                Else
                                                    If firstRap = True Then
                                                        LastIsNotPlayingRange = True
                                                        firstRap = False
                                                    End If
                                                End If
                                            End If
                                        End If
                                        firstRap = False

                                        If nextRap = True Then
                                            If SatzMinusRegel = "R11" And CheckBox7.Checked = True Then 'Check R14'
                                                If StrangRapCnt < 3 Then
                                                    tmpCopyStruct.TmRS = "KS: R14"

                                                ElseIf StrangRapCnt = 3 Then
                                                    If SatzMinus = "X" Then
                                                        If (StrangRap(0) = "O" And StrangRap(1) = "X") Then
                                                            tmpCopyStruct.TmRS = tmpSatz
                                                            Exit For
                                                        End If
                                                    ElseIf SatzMinus = "O" Then
                                                        If (StrangRap(0) = "X" And StrangRap(1) = "O") Then
                                                            tmpCopyStruct.TmRS = tmpSatz
                                                            Exit For
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If CheckBox8.Checked = True And LastIsNotPlayingRange = True And StrangRap.Length >= 3 Then 'Check R15
                                                    If StrangRap.Substring(0, 2) = "XO" Then 'R15= 1er = jetzt prüfen ob danach 2er
                                                        If StrangRap.Substring(StrangRap.Length - 1, 1) = "X" And StrangRap2 = "" Then
                                                            StrangRap2 = "X"
                                                        ElseIf StrangRap2 <> "" Then
                                                            If (StrangRap2 & StrangRap.Substring(StrangRap.Length - 1, 1)) = "XX" And SatzMinus = "O" Then
                                                                tmpCopyStruct.TmRS = "KS: R15"
                                                                Exit For
                                                            ElseIf (StrangRap2 & StrangRap.Substring(StrangRap.Length - 1, 1)) = "XO" And SatzMinus = "X" Then
                                                                tmpCopyStruct.TmRS = "KS: R15"
                                                                Exit For
                                                            End If
                                                        End If
                                                    ElseIf StrangRap.Substring(0, 2) = "OX" Then
                                                        If StrangRap.Substring(StrangRap.Length - 1, 1) = "O" And StrangRap2 = "" Then
                                                            StrangRap2 = "O"
                                                        ElseIf StrangRap2 <> "" Then
                                                            If (StrangRap2 & StrangRap.Substring(StrangRap.Length - 1, 1)) = "OO" And SatzMinus = "X" Then
                                                                tmpCopyStruct.TmRS = "KS: R15"
                                                                Exit For
                                                            ElseIf (StrangRap2 & StrangRap.Substring(StrangRap.Length - 1, 1)) = "OX" And SatzMinus = "O" Then
                                                                tmpCopyStruct.TmRS = "KS: R15"
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                End If

                                                If CheckBox9.Checked = True Then 'Check R16
                                                    If StrangRap.Length = 3 Then
                                                        If StrangRap = "XXO" Then
                                                            If SatzMinus = "O" Then
                                                                tmpCopyStruct.TmRS = "KS: R16"
                                                            End If
                                                            Exit For
                                                        ElseIf StrangRap = "OOX" Then
                                                            If SatzMinus = "X" Then
                                                                tmpCopyStruct.TmRS = "KS: R16"
                                                            End If
                                                            Exit For
                                                        End If
                                                    ElseIf StrangRap.Length = 5 Then
                                                        If StrangRap = "XXXXO" Then
                                                            If SatzMinus = "O" Then
                                                                tmpCopyStruct.TmRS = "KS: R16"
                                                            End If
                                                            Exit For
                                                        ElseIf StrangRap = "OOOOX" Then
                                                            If SatzMinus = "X" Then
                                                                tmpCopyStruct.TmRS = "KS: R16"
                                                            End If
                                                            Exit For
                                                        End If
                                                    End If
                                                End If

                                                If CheckBox10.Checked = True Then 'Check R17
                                                    If StrangRap.Length = 4 Then
                                                        If StrangRap = "XXXO" Then
                                                            If SatzMinus = "X" Then
                                                                tmpCopyStruct.TmRS = "KS: R17"
                                                            End If
                                                            Exit For
                                                        ElseIf StrangRap = "OOOX" Then
                                                            If SatzMinus = "O" Then
                                                                tmpCopyStruct.TmRS = "KS: R17"
                                                            End If
                                                            Exit For
                                                        End If
                                                    End If
                                                End If

                                                If CheckBox11.Checked = True Then 'Check R18
                                                    If StrangRap.Length = 5 Then
                                                        If StrangRap = "XXXXX" Or StrangRap = "OOOOO" Then
                                                            tmpCopyStruct.TmRS = "KS: R18"
                                                            Exit For
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Next Z

                                If CheckBox7.Checked = True And SatzMinusRegel = "R11" Then
                                    If StrangRap.Length < 3 Then
                                        tmpCopyStruct.TmRS = "Vl. zu klein"
                                    End If
                                End If

                                If CheckBox8.Checked = True And LastIsNotPlayingRange = True Then 'R15 fehlt Vorlauf
                                    If StrangRap.Length = 1 Then
                                        If tmpCopyStruct.TmRS <> "KS: R15" Then
                                            tmpCopyStruct.TmRS = "Vl. zu klein"
                                        End If
                                    ElseIf StrangRap.Length >= 2 Then
                                        If StrangRap.Substring(0, 2) = "OX" Or StrangRap.Substring(0, 2) = "XO" Then
                                            If tmpCopyStruct.TmRS <> "KS: R15" Then
                                                tmpCopyStruct.TmRS = "Vl. zu klein"
                                            End If
                                        End If
                                    Else
                                        tmpCopyStruct.TmRS = "Vl. zu klein"
                                    End If
                                End If

                                If CheckBox9.Checked = True Then
                                    If StrangRap = "X" Or StrangRap = "XX" Or StrangRap = "XXX" Or StrangRap = "XXXX" Then
                                        If SatzMinus = "O" Then
                                            tmpCopyStruct.TmRS = "Vl. zu klein"
                                        End If
                                    ElseIf StrangRap = "O" Or StrangRap = "OO" Or StrangRap = "OOO" Or StrangRap = "OOO" Then
                                        If SatzMinus = "O" Then
                                            tmpCopyStruct.TmRS = "Vl. zu klein"
                                        End If
                                    End If
                                End If

                                If CheckBox10.Checked = True Then
                                    If StrangRap = "X" Or StrangRap = "XX" Or StrangRap = "XXX" Then
                                        If SatzMinus = "X" Then
                                            tmpCopyStruct.TmRS = "Vl. zu klein"
                                        End If
                                    ElseIf StrangRap = "O" Or StrangRap = "OO" Or StrangRap = "OOO" Then
                                        If SatzMinus = "O" Then
                                            tmpCopyStruct.TmRS = "Vl. zu klein"
                                        End If
                                    End If
                                End If

                                If CheckBox11.Checked = True Then
                                    If StrangRap = "X" Or StrangRap = "XX" Or StrangRap = "XXX" Or StrangRap = "XXXX" Or StrangRap = "O" Or StrangRap = "OO" Or StrangRap = "OOO" Or StrangRap = "OOOO" Then
                                        tmpCopyStruct.TmRS = "Vl. zu klein"
                                    End If
                                End If
                            End If

                            DeletelastDummyCoup()
                            ''zurückschreiben
                            cCoupData = tmpCopyStruct
                            SatzMinus = ""
                            SatzMinusRegel = ""
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
        Dim arraylistCnt As Integer = arrayPartie.Count
        Dim localInstanceofMyClass As clsCoupData

        Dim Strang As String = ""
        Dim StrangRap As String = ""
        Dim StrangRapCnt As Integer = 0
        Dim firstRap As Boolean = True
        Dim LastIsNotPlayingRange As Boolean = False
        Dim tmpSatz As String = ""
        Dim StrangRap2 As String = ""
        Dim nextRap As Boolean = False

        If SatzSchwarz <> "" Then
            Rap = cCoupData.SR
            If Rap <> "" Then Exit Sub

            Ser = cCoupData.SS
            Int = cCoupData.SI

            If Ser <> "" Then
                SerInt = "S" + SerInt
                SerIntCnt = SerIntCnt + 1
            End If

            If Int <> "" Then
                SerInt = "I" + SerInt
                SerIntCnt = SerIntCnt + 1
            End If

            For I As Integer = 0 To arrayPartie.Count - 1
                localInstanceofMyClass = CType(arrayPartie(arraylistCnt - I - 1), clsCoupData)
                Ser = localInstanceofMyClass.SS
                Int = localInstanceofMyClass.SI

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
                sAkt = cCoupData.S
                For I As Integer = 1 To arraylistCnt - 1
                    localInstanceofMyClass = CType(arrayPartie(arraylistCnt - I), clsCoupData)
                    coup = localInstanceofMyClass.Coup
                    sLast = localInstanceofMyClass.S

                    If coup <> "00" And coup <> "0" Then
                        If sAkt <> "" And sLast = "" Then 'And coup <> "0" Then                            
                            Dim tmpCopyStruct As clsCoupData = cCoupData
                            arrayPartie.Add(cCoupData) 'momentaner Datensatz in array
                            InsertNewCoup("1", True)
                            If cCoupData.SR = SatzSchwarz Then
                                tmpCopyStruct.SRS = "Satz auf " & SatzSchwarz & "/R"
                                tmpSatz = "Satz auf " & SatzSchwarz & "/R"
                            Else
                                tmpCopyStruct.SRS = "Satz auf " & SatzSchwarz & "/S"
                                tmpSatz = "Satz auf " & SatzSchwarz & "/S"
                            End If

                            Ser = cCoupData.SS
                            Int = cCoupData.SI

                            If SerInt = "ISS" Then
                                'Serienstrang
                                Strang = "Serie"
                            ElseIf SerInt = "SII" Then
                                'Intermittenztstrang
                                Strang = "Int"
                            End If

                            If CheckBox7.Checked = True Or CheckBox8.Checked = True Or CheckBox9.Checked = True Or CheckBox10.Checked = True Or CheckBox11.Checked = True Then
                                For Z As Integer = 2 To arraylistCnt - 1
                                    localInstanceofMyClass = CType(arrayPartie(arraylistCnt - Z), clsCoupData)
                                    Rap = localInstanceofMyClass.SR

                                    If Rap <> "" Then
                                        nextRap = False
                                        If Strang = "Serie" Then
                                            If Rap = "X" Then
                                                If localInstanceofMyClass.SS <> "" Then
                                                    StrangRap = StrangRap & "X"
                                                    StrangRapCnt = StrangRapCnt + 1
                                                    nextRap = True
                                                Else
                                                    If firstRap = True Then
                                                        LastIsNotPlayingRange = True
                                                        firstRap = False
                                                    End If
                                                End If
                                            ElseIf Rap = "O" Then
                                                If localInstanceofMyClass.SI <> "" Then
                                                    StrangRap = StrangRap & "O"
                                                    StrangRapCnt = StrangRapCnt + 1
                                                    nextRap = True
                                                Else
                                                    If firstRap = True Then
                                                        LastIsNotPlayingRange = True
                                                        firstRap = False
                                                    End If
                                                End If
                                            End If
                                        ElseIf Strang = "Int" Then
                                            If Rap = "X" Then
                                                If localInstanceofMyClass.SI <> "" Then
                                                    StrangRap = StrangRap & "X"
                                                    StrangRapCnt = StrangRapCnt + 1
                                                    nextRap = True
                                                Else
                                                    If firstRap = True Then
                                                        LastIsNotPlayingRange = True
                                                        firstRap = False
                                                    End If
                                                End If
                                            ElseIf Rap = "O" Then
                                                If localInstanceofMyClass.SS <> "" Then
                                                    StrangRap = StrangRap & "O"
                                                    StrangRapCnt = StrangRapCnt + 1
                                                    nextRap = True
                                                Else
                                                    If firstRap = True Then
                                                        LastIsNotPlayingRange = True
                                                        firstRap = False
                                                    End If
                                                End If
                                            End If
                                        End If
                                        firstRap = False

                                        If nextRap = True Then
                                            If SatzSchwarzRegel = "R11" And CheckBox7.Checked = True Then 'Check R14'
                                                If StrangRapCnt < 3 Then
                                                    tmpCopyStruct.SRS = "KS: R14"

                                                ElseIf StrangRapCnt = 3 Then
                                                    If SatzSchwarz = "X" Then
                                                        If (StrangRap(0) = "O" And StrangRap(1) = "X") Then
                                                            tmpCopyStruct.SRS = tmpSatz
                                                            Exit For
                                                        End If
                                                    ElseIf SatzSchwarz = "O" Then
                                                        If (StrangRap(0) = "X" And StrangRap(1) = "O") Then
                                                            tmpCopyStruct.SRS = tmpSatz
                                                            Exit For
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If CheckBox8.Checked = True And LastIsNotPlayingRange = True And StrangRap.Length >= 3 Then 'Check R15
                                                    If StrangRap.Substring(0, 2) = "XO" Then 'R15= 1er = jetzt prüfen ob danach 2er
                                                        If StrangRap.Substring(StrangRap.Length - 1, 1) = "X" And StrangRap2 = "" Then
                                                            StrangRap2 = "X"
                                                        ElseIf StrangRap2 <> "" Then
                                                            If (StrangRap2 & StrangRap.Substring(StrangRap.Length - 1, 1)) = "XX" And SatzSchwarz = "O" Then
                                                                tmpCopyStruct.SRS = "KS: R15"
                                                                Exit For
                                                            ElseIf (StrangRap2 & StrangRap.Substring(StrangRap.Length - 1, 1)) = "XO" And SatzSchwarz = "X" Then
                                                                tmpCopyStruct.SRS = "KS: R15"
                                                                Exit For
                                                            End If
                                                        End If
                                                    ElseIf StrangRap.Substring(0, 2) = "OX" Then
                                                        If StrangRap.Substring(StrangRap.Length - 1, 1) = "O" And StrangRap2 = "" Then
                                                            StrangRap2 = "O"
                                                        ElseIf StrangRap2 <> "" Then
                                                            If (StrangRap2 & StrangRap.Substring(StrangRap.Length - 1, 1)) = "OO" And SatzSchwarz = "X" Then
                                                                tmpCopyStruct.SRS = "KS: R15"
                                                                Exit For
                                                            ElseIf (StrangRap2 & StrangRap.Substring(StrangRap.Length - 1, 1)) = "OX" And SatzSchwarz = "O" Then
                                                                tmpCopyStruct.SRS = "KS: R15"
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                End If

                                                If CheckBox9.Checked = True Then 'Check R16
                                                    If StrangRap.Length = 3 Then
                                                        If StrangRap = "XXO" Then
                                                            If SatzSchwarz = "O" Then
                                                                tmpCopyStruct.SRS = "KS: R16"
                                                            End If
                                                            Exit For
                                                        ElseIf StrangRap = "OOX" Then
                                                            If SatzSchwarz = "X" Then
                                                                tmpCopyStruct.SRS = "KS: R16"
                                                            End If
                                                            Exit For
                                                        End If
                                                    ElseIf StrangRap.Length = 5 Then
                                                        If StrangRap = "XXXXO" Then
                                                            If SatzSchwarz = "O" Then
                                                                tmpCopyStruct.SRS = "KS: R16"
                                                            End If
                                                            Exit For
                                                        ElseIf StrangRap = "OOOOX" Then
                                                            If SatzSchwarz = "X" Then
                                                                tmpCopyStruct.SRS = "KS: R16"
                                                            End If
                                                            Exit For
                                                        End If
                                                    End If
                                                End If

                                                If CheckBox10.Checked = True Then 'Check R17
                                                    If StrangRap.Length = 4 Then
                                                        If StrangRap = "XXXO" Then
                                                            If SatzSchwarz = "X" Then
                                                                tmpCopyStruct.SRS = "KS: R17"
                                                            End If
                                                            Exit For
                                                        ElseIf StrangRap = "OOOX" Then
                                                            If SatzSchwarz = "O" Then
                                                                tmpCopyStruct.SRS = "KS: R17"
                                                            End If
                                                            Exit For
                                                        End If
                                                    End If
                                                End If

                                                If CheckBox11.Checked = True Then 'Check R18
                                                    If StrangRap.Length = 5 Then
                                                        If StrangRap = "XXXXX" Or StrangRap = "OOOOO" Then
                                                            tmpCopyStruct.SRS = "KS: R18"
                                                            Exit For
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Next Z

                                If CheckBox7.Checked = True And SatzSchwarzRegel = "R11" Then
                                    If StrangRap.Length < 3 Then
                                        tmpCopyStruct.SRS = "Vl. zu klein"
                                    End If
                                End If

                                If CheckBox8.Checked = True And LastIsNotPlayingRange = True Then 'R15 fehlt Vorlauf
                                    If StrangRap.Length = 1 Then
                                        If tmpCopyStruct.SRS <> "KS: R15" Then
                                            tmpCopyStruct.SRS = "Vl. zu klein"
                                        End If
                                    ElseIf StrangRap.Length >= 2 Then
                                        If StrangRap.Substring(0, 2) = "OX" Or StrangRap.Substring(0, 2) = "XO" Then
                                            If tmpCopyStruct.SRS <> "KS: R15" Then
                                                tmpCopyStruct.SRS = "Vl. zu klein"
                                            End If
                                        End If
                                    Else
                                        tmpCopyStruct.SRS = "Vl. zu klein"
                                    End If
                                End If

                                If CheckBox9.Checked = True Then
                                    If StrangRap = "X" Or StrangRap = "XX" Or StrangRap = "XXX" Or StrangRap = "XXXX" Then
                                        If SatzSchwarz = "O" Then
                                            tmpCopyStruct.SRS = "Vl. zu klein"
                                        End If
                                    ElseIf StrangRap = "O" Or StrangRap = "OO" Or StrangRap = "OOO" Or StrangRap = "OOO" Then
                                        If SatzSchwarz = "O" Then
                                            tmpCopyStruct.SRS = "Vl. zu klein"
                                        End If
                                    End If
                                End If

                                If CheckBox10.Checked = True Then
                                    If StrangRap = "X" Or StrangRap = "XX" Or StrangRap = "XXX" Then
                                        If SatzSchwarz = "X" Then
                                            tmpCopyStruct.SRS = "Vl. zu klein"
                                        End If
                                    ElseIf StrangRap = "O" Or StrangRap = "OO" Or StrangRap = "OOO" Then
                                        If SatzSchwarz = "O" Then
                                            tmpCopyStruct.SRS = "Vl. zu klein"
                                        End If
                                    End If
                                End If

                                If CheckBox11.Checked = True Then
                                    If StrangRap = "X" Or StrangRap = "XX" Or StrangRap = "XXX" Or StrangRap = "XXXX" Or StrangRap = "O" Or StrangRap = "OO" Or StrangRap = "OOO" Or StrangRap = "OOOO" Then
                                        tmpCopyStruct.SRS = "Vl. zu klein"
                                    End If
                                End If
                            End If


                            DeletelastDummyCoup()
                            ''zurückschreiben
                            cCoupData = tmpCopyStruct
                            SatzSchwarz = ""
                            SatzSchwarzRegel = ""
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
        Dim arraylistCnt As Integer = arrayPartie.Count
        Dim localInstanceofMyClass As clsCoupData

        Dim Strang As String = ""
        Dim StrangRap As String = ""
        Dim StrangRapCnt As Integer = 0
        Dim firstRap As Boolean = True
        Dim LastIsNotPlayingRange As Boolean = False
        Dim tmpSatz As String = ""
        Dim StrangRap2 As String = ""
        Dim nextRap As Boolean = False

        If SatzRot <> "" Then
            Rap = cCoupData.RR
            If Rap <> "" Then Exit Sub

            Ser = cCoupData.RS
            Int = cCoupData.RI

            If Ser <> "" Then
                SerInt = "S" + SerInt
                SerIntCnt = SerIntCnt + 1
            End If

            If Int <> "" Then
                SerInt = "I" + SerInt
                SerIntCnt = SerIntCnt + 1
            End If

            For I As Integer = 0 To arrayPartie.Count - 1
                localInstanceofMyClass = CType(arrayPartie(arraylistCnt - I - 1), clsCoupData)
                Ser = localInstanceofMyClass.RS
                Int = localInstanceofMyClass.RI

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
                sAkt = cCoupData.R
                For I As Integer = 1 To arraylistCnt - 1
                    localInstanceofMyClass = CType(arrayPartie(arraylistCnt - I), clsCoupData)
                    coup = localInstanceofMyClass.Coup
                    sLast = localInstanceofMyClass.R

                    If coup <> "00" And coup <> "0" Then
                        If sAkt <> "" And sLast = "" Then 'And coup <> "0" Then                            
                            Dim tmpCopyStruct As clsCoupData = cCoupData
                            arrayPartie.Add(cCoupData) 'momentaner Datensatz in array
                            InsertNewCoup("1", True)
                            If cCoupData.RR = SatzRot Then
                                tmpCopyStruct.RRS = "Satz auf " & SatzRot & "/R"
                                tmpSatz = "Satz auf " & SatzRot & "/R"
                            Else
                                tmpCopyStruct.RRS = "Satz auf " & SatzRot & "/S"
                                tmpSatz = "Satz auf " & SatzRot & "/S"
                            End If

                            Ser = cCoupData.RS
                            Int = cCoupData.RI

                            If SerInt = "ISS" Then
                                'Serienstrang
                                Strang = "Serie"
                            ElseIf SerInt = "SII" Then
                                'Intermittenztstrang
                                Strang = "Int"
                            End If

                            If CheckBox7.Checked = True Or CheckBox8.Checked = True Or CheckBox9.Checked = True Or CheckBox10.Checked = True Or CheckBox11.Checked = True Then
                                For Z As Integer = 2 To arraylistCnt - 1
                                    localInstanceofMyClass = CType(arrayPartie(arraylistCnt - Z), clsCoupData)
                                    Rap = localInstanceofMyClass.RR

                                    If Rap <> "" Then
                                        nextRap = False
                                        If Strang = "Serie" Then
                                            If Rap = "X" Then
                                                If localInstanceofMyClass.RS <> "" Then
                                                    StrangRap = StrangRap & "X"
                                                    StrangRapCnt = StrangRapCnt + 1
                                                    nextRap = True
                                                Else
                                                    If firstRap = True Then
                                                        LastIsNotPlayingRange = True
                                                        firstRap = False
                                                    End If
                                                End If
                                            ElseIf Rap = "O" Then
                                                If localInstanceofMyClass.RI <> "" Then
                                                    StrangRap = StrangRap & "O"
                                                    StrangRapCnt = StrangRapCnt + 1
                                                    nextRap = True
                                                Else
                                                    If firstRap = True Then
                                                        LastIsNotPlayingRange = True
                                                        firstRap = False
                                                    End If
                                                End If
                                            End If
                                        ElseIf Strang = "Int" Then
                                            If Rap = "X" Then
                                                If localInstanceofMyClass.RI <> "" Then
                                                    StrangRap = StrangRap & "X"
                                                    StrangRapCnt = StrangRapCnt + 1
                                                    nextRap = True
                                                Else
                                                    If firstRap = True Then
                                                        LastIsNotPlayingRange = True
                                                        firstRap = False
                                                    End If
                                                End If
                                            ElseIf Rap = "O" Then
                                                If localInstanceofMyClass.RS <> "" Then
                                                    StrangRap = StrangRap & "O"
                                                    StrangRapCnt = StrangRapCnt + 1
                                                    nextRap = True
                                                Else
                                                    If firstRap = True Then
                                                        LastIsNotPlayingRange = True
                                                        firstRap = False
                                                    End If
                                                End If
                                            End If
                                        End If
                                        firstRap = False

                                        If nextRap = True Then
                                            If SatzRotRegel = "R11" And CheckBox7.Checked = True Then 'Check R14'
                                                If StrangRapCnt < 3 Then
                                                    tmpCopyStruct.RRS = "KS: R14"

                                                ElseIf StrangRapCnt = 3 Then
                                                    If SatzRot = "X" Then
                                                        If (StrangRap(0) = "O" And StrangRap(1) = "X") Then
                                                            tmpCopyStruct.RRS = tmpSatz
                                                            Exit For
                                                        End If
                                                    ElseIf SatzRot = "O" Then
                                                        If (StrangRap(0) = "X" And StrangRap(1) = "O") Then
                                                            tmpCopyStruct.RRS = tmpSatz
                                                            Exit For
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If CheckBox8.Checked = True And LastIsNotPlayingRange = True And StrangRap.Length >= 3 Then 'Check R15
                                                    If StrangRap.Substring(0, 2) = "XO" Then 'R15= 1er = jetzt prüfen ob danach 2er
                                                        If StrangRap.Substring(StrangRap.Length - 1, 1) = "X" And StrangRap2 = "" Then
                                                            StrangRap2 = "X"
                                                        ElseIf StrangRap2 <> "" Then
                                                            If (StrangRap2 & StrangRap.Substring(StrangRap.Length - 1, 1)) = "XX" And SatzRot = "O" Then
                                                                tmpCopyStruct.RRS = "KS: R15"
                                                                Exit For
                                                            ElseIf (StrangRap2 & StrangRap.Substring(StrangRap.Length - 1, 1)) = "XO" And SatzRot = "X" Then
                                                                tmpCopyStruct.RRS = "KS: R15"
                                                                Exit For
                                                            End If
                                                        End If
                                                    ElseIf StrangRap.Substring(0, 2) = "OX" Then
                                                        If StrangRap.Substring(StrangRap.Length - 1, 1) = "O" And StrangRap2 = "" Then
                                                            StrangRap2 = "O"
                                                        ElseIf StrangRap2 <> "" Then
                                                            If (StrangRap2 & StrangRap.Substring(StrangRap.Length - 1, 1)) = "OO" And SatzRot = "X" Then
                                                                tmpCopyStruct.RRS = "KS: R15"
                                                                Exit For
                                                            ElseIf (StrangRap2 & StrangRap.Substring(StrangRap.Length - 1, 1)) = "OX" And SatzRot = "O" Then
                                                                tmpCopyStruct.RRS = "KS: R15"
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                End If

                                                If CheckBox9.Checked = True Then 'Check R16
                                                    If StrangRap.Length = 3 Then
                                                        If StrangRap = "XXO" Then
                                                            If SatzRot = "O" Then
                                                                tmpCopyStruct.RRS = "KS: R16"
                                                            End If
                                                            Exit For
                                                        ElseIf StrangRap = "OOX" Then
                                                            If SatzRot = "X" Then
                                                                tmpCopyStruct.RRS = "KS: R16"
                                                            End If
                                                            Exit For
                                                        End If
                                                    ElseIf StrangRap.Length = 5 Then
                                                        If StrangRap = "XXXXO" Then
                                                            If SatzRot = "O" Then
                                                                tmpCopyStruct.RRS = "KS: R16"
                                                            End If
                                                            Exit For
                                                        ElseIf StrangRap = "OOOOX" Then
                                                            If SatzRot = "X" Then
                                                                tmpCopyStruct.RRS = "KS: R16"
                                                            End If
                                                            Exit For
                                                        End If
                                                    End If
                                                End If

                                                If CheckBox10.Checked = True Then 'Check R17
                                                    If StrangRap.Length = 4 Then
                                                        If StrangRap = "XXXO" Then
                                                            If SatzRot = "X" Then
                                                                tmpCopyStruct.RRS = "KS: R17"
                                                            End If
                                                            Exit For
                                                        ElseIf StrangRap = "OOOX" Then
                                                            If SatzRot = "O" Then
                                                                tmpCopyStruct.RRS = "KS: R17"
                                                            End If
                                                            Exit For
                                                        End If
                                                    End If
                                                End If

                                                If CheckBox11.Checked = True Then 'Check R18
                                                    If StrangRap.Length = 5 Then
                                                        If StrangRap = "XXXXX" Or StrangRap = "OOOOO" Then
                                                            tmpCopyStruct.RRS = "KS: R18"
                                                            Exit For
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Next Z

                                If CheckBox7.Checked = True And SatzRotRegel = "R11" Then
                                    If StrangRap.Length < 3 Then
                                        tmpCopyStruct.RRS = "Vl. zu klein"
                                    End If
                                End If

                                If CheckBox8.Checked = True And LastIsNotPlayingRange = True Then 'R15 fehlt Vorlauf
                                    If StrangRap.Length = 1 Then
                                        If tmpCopyStruct.RRS <> "KS: R15" Then
                                            tmpCopyStruct.RRS = "Vl. zu klein"
                                        End If
                                    ElseIf StrangRap.Length >= 2 Then
                                        If StrangRap.Substring(0, 2) = "OX" Or StrangRap.Substring(0, 2) = "XO" Then
                                            If tmpCopyStruct.RRS <> "KS: R15" Then
                                                tmpCopyStruct.RRS = "Vl. zu klein"
                                            End If
                                        End If
                                    Else
                                        tmpCopyStruct.RRS = "Vl. zu klein"
                                    End If
                                End If

                                If CheckBox9.Checked = True Then
                                    If StrangRap = "X" Or StrangRap = "XX" Or StrangRap = "XXX" Or StrangRap = "XXXX" Then
                                        If SatzRot = "O" Then
                                            tmpCopyStruct.RRS = "Vl. zu klein"
                                        End If
                                    ElseIf StrangRap = "O" Or StrangRap = "OO" Or StrangRap = "OOO" Or StrangRap = "OOO" Then
                                        If SatzRot = "O" Then
                                            tmpCopyStruct.RRS = "Vl. zu klein"
                                        End If
                                    End If
                                End If

                                If CheckBox10.Checked = True Then
                                    If StrangRap = "X" Or StrangRap = "XX" Or StrangRap = "XXX" Then
                                        If SatzRot = "X" Then
                                            tmpCopyStruct.RRS = "Vl. zu klein"
                                        End If
                                    ElseIf StrangRap = "O" Or StrangRap = "OO" Or StrangRap = "OOO" Then
                                        If SatzRot = "O" Then
                                            tmpCopyStruct.RRS = "Vl. zu klein"
                                        End If
                                    End If
                                End If

                                If CheckBox11.Checked = True Then
                                    If StrangRap = "X" Or StrangRap = "XX" Or StrangRap = "XXX" Or StrangRap = "XXXX" Or StrangRap = "O" Or StrangRap = "OO" Or StrangRap = "OOO" Or StrangRap = "OOOO" Then
                                        tmpCopyStruct.RRS = "Vl. zu klein"
                                    End If
                                End If
                            End If

                            DeletelastDummyCoup()
                            ''zurückschreiben
                            cCoupData = tmpCopyStruct
                            SatzRot = ""
                            SatzRotRegel = ""
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
        Dim arraylistCnt As Integer = arrayPartie.Count
        Dim localInstanceofMyClass As clsCoupData

        coup = cCoupData.Coup
        If coup <> 0 Then
            Plus1 = cCoupData.TP
        End If

        If arraylistCnt = 0 Then Exit Sub

        For I As Integer = 0 To arrayPartie.Count - 1
            localInstanceofMyClass = CType(arrayPartie(arraylistCnt - I - 1), clsCoupData)
            coup = localInstanceofMyClass.Coup
            If coup <> 0 Then
                PlusTmp = localInstanceofMyClass.TP
                Select Case countPlus
                    Case 1
                        Plus2 = PlusTmp
                        countPlus = 2
                    Case 2
                        Plus3 = PlusTmp
                        countPlus = 3
                End Select
            End If
            If countPlus = 3 Then Exit For
        Next

        If Plus1 <> "" And Plus2 <> "" And Plus3 = "" Then 'Serie            
            cCoupData.TPS = "X"
        End If

        If Plus1 = "" And Plus2 <> "" And Plus3 = "" Then 'Intermittenz            
            cCoupData.TPI = "X"
        End If
    End Sub

    Private Sub fillSelektorMinus(Optional ByVal fiktiv As Boolean = False)

        Dim Minus1 As String = ""
        Dim Minus2 As String = ""
        Dim Minus3 As String = ""
        Dim MinusTmp As String = ""
        Dim countMinus As Integer = 1
        Dim coup As Integer
        Dim arraylistCnt As Integer = arrayPartie.Count
        Dim localInstanceofMyClass As clsCoupData

        coup = cCoupData.Coup
        If coup <> 0 Then
            Minus1 = cCoupData.TM
        End If

        If arraylistCnt = 0 Then Exit Sub

        For I As Integer = 0 To arrayPartie.Count - 1
            localInstanceofMyClass = CType(arrayPartie(arraylistCnt - I - 1), clsCoupData)
            coup = localInstanceofMyClass.Coup
            If coup <> 0 Then
                MinusTmp = localInstanceofMyClass.TM
                Select Case countMinus
                    Case 1
                        Minus2 = MinusTmp
                        countMinus = 2
                    Case 2
                        Minus3 = MinusTmp
                        countMinus = 3
                End Select
            End If
            If countMinus = 3 Then Exit For
        Next

        If Minus1 <> "" And Minus2 <> "" And Minus3 = "" Then 'Serie            
            cCoupData.TmS = "X"
        End If

        If Minus1 = "" And Minus2 <> "" And Minus3 = "" Then 'Intermittenz            
            cCoupData.TmI = "X"
        End If
    End Sub

    Private Sub fillSelektorSchwarz(Optional ByVal fiktiv As Boolean = False)

        Dim Schwarz1 As String = ""
        Dim Schwarz2 As String = ""
        Dim Schwarz3 As String = ""
        Dim SchwarzTmp As String = ""
        Dim countSchwarz As Integer = 1
        Dim coup As Integer
        Dim arraylistCnt As Integer = arrayPartie.Count
        Dim localInstanceofMyClass As clsCoupData

        coup = cCoupData.Coup
        If coup <> 0 Then
            Schwarz1 = cCoupData.S
        End If

        If arraylistCnt = 0 Then Exit Sub

        For I As Integer = 0 To arrayPartie.Count - 1
            localInstanceofMyClass = CType(arrayPartie(arraylistCnt - I - 1), clsCoupData)
            coup = localInstanceofMyClass.Coup
            If coup <> 0 Then
                SchwarzTmp = localInstanceofMyClass.S
                Select Case countSchwarz
                    Case 1
                        Schwarz2 = SchwarzTmp
                        countSchwarz = 2
                    Case 2
                        Schwarz3 = SchwarzTmp
                        countSchwarz = 3
                End Select
            End If
            If countSchwarz = 3 Then Exit For
        Next

        If Schwarz1 <> "" And Schwarz2 <> "" And Schwarz3 = "" Then 'Serie            
            cCoupData.SS = "X"
        End If

        If Schwarz1 = "" And Schwarz2 <> "" And Schwarz3 = "" Then 'Intermittenz            
            cCoupData.SI = "X"
        End If
    End Sub

    Private Sub fillSelektorRot(Optional ByVal fiktiv As Boolean = False)

        Dim Rot1 As String = ""
        Dim Rot2 As String = ""
        Dim Rot3 As String = ""
        Dim RotTmp As String = ""
        Dim countRot As Integer = 1
        Dim coup As Integer
        Dim arraylistCnt As Integer = arrayPartie.Count
        Dim localInstanceofMyClass As clsCoupData

        coup = cCoupData.Coup
        If coup <> 0 Then
            Rot1 = cCoupData.R
        End If

        If arraylistCnt = 0 Then Exit Sub

        For I As Integer = 0 To arrayPartie.Count - 1
            localInstanceofMyClass = CType(arrayPartie(arraylistCnt - I - 1), clsCoupData)
            coup = localInstanceofMyClass.Coup
            If coup <> 0 Then
                RotTmp = localInstanceofMyClass.R
                Select Case countRot
                    Case 1
                        Rot2 = RotTmp
                        countRot = 2
                    Case 2
                        Rot3 = RotTmp
                        countRot = 3
                End Select
            End If
            If countRot = 3 Then Exit For
        Next

        If Rot1 <> "" And Rot2 <> "" And Rot3 = "" Then 'Serie            
            cCoupData.RS = "X"
        End If

        If Rot1 = "" And Rot2 <> "" And Rot3 = "" Then 'Intermittenz            
            cCoupData.RI = "X"
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
        Dim arraylistCnt As Integer = arrayPartie.Count
        Dim localInstanceofMyClass As clsCoupData

        PlusSerie = cCoupData.TPS
        PlusIntermit = cCoupData.TPI
        coup = cCoupData.Coup

        If (PlusSerie = "" And PlusIntermit = "") Or coup = 0 Then Exit Sub
        If PlusSerie <> "" Then
            PlusRap4 = "S"
        ElseIf PlusIntermit <> "" Then
            PlusRap4 = "I"
        End If

        For I As Integer = 0 To arrayPartie.Count - 1
            localInstanceofMyClass = CType(arrayPartie(arraylistCnt - I - 1), clsCoupData)

            coup = localInstanceofMyClass.Coup
            If coup <> 0 Then
                PlusSerie = localInstanceofMyClass.TPS
                PlusIntermit = localInstanceofMyClass.TPI
                If PlusSerie <> "" Then
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
                ElseIf PlusIntermit <> "" Then
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

        cCoupData.TPR = Rap

        If Rap <> "" Then
            cCoupData.TPR7 = GetRapLast7(ColNum.eTPR)
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
        Dim arraylistCnt As Integer = arrayPartie.Count
        Dim localInstanceofMyClass As clsCoupData

        MinusSerie = cCoupData.TmS
        MinusIntermit = cCoupData.TmI
        coup = cCoupData.Coup

        If (MinusSerie = "" And MinusIntermit = "") Or coup = 0 Then Exit Sub
        If MinusSerie <> "" Then
            MinusRap4 = "S"
        ElseIf MinusIntermit <> "" Then
            MinusRap4 = "I"
        End If

        For I As Integer = 0 To arrayPartie.Count - 1
            localInstanceofMyClass = CType(arrayPartie(arraylistCnt - I - 1), clsCoupData)

            coup = localInstanceofMyClass.Coup
            If coup <> 0 Then
                MinusSerie = localInstanceofMyClass.TmS
                MinusIntermit = localInstanceofMyClass.TmI
                If MinusSerie <> "" Then
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
                ElseIf MinusIntermit <> "" Then
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

        cCoupData.TmR = Rap

        If Rap <> "" Then
            cCoupData.TmR7 = GetRapLast7(ColNum.eTMR)
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
        Dim arraylistCnt As Integer = arrayPartie.Count
        Dim localInstanceofMyClass As clsCoupData

        SchwarzSerie = cCoupData.SS
        SchwarzIntermit = cCoupData.SI
        coup = cCoupData.Coup

        If (SchwarzSerie = "" And SchwarzIntermit = "") Or coup = 0 Then Exit Sub
        If SchwarzSerie <> "" Then
            SchwarzRap4 = "S"
        ElseIf SchwarzIntermit <> "" Then
            SchwarzRap4 = "I"
        End If

        For I As Integer = 0 To arrayPartie.Count - 1
            localInstanceofMyClass = CType(arrayPartie(arraylistCnt - I - 1), clsCoupData)

            coup = localInstanceofMyClass.Coup
            If coup <> 0 Then
                SchwarzSerie = localInstanceofMyClass.SS
                SchwarzIntermit = localInstanceofMyClass.SI
                If SchwarzSerie <> "" Then
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
                ElseIf SchwarzIntermit <> "" Then
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

        cCoupData.SR = Rap

        If Rap <> "" Then
            cCoupData.SR7 = GetRapLast7(ColNum.eSR)
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
        Dim arraylistCnt As Integer = arrayPartie.Count
        Dim localInstanceofMyClass As clsCoupData

        RotSerie = cCoupData.RS
        RotIntermit = cCoupData.RI
        coup = cCoupData.Coup

        If (RotSerie = "" And RotIntermit = "") Or coup = 0 Then Exit Sub
        If RotSerie <> "" Then
            RotRap4 = "S"
        ElseIf RotIntermit <> "" Then
            RotRap4 = "I"
        End If

        For I As Integer = 0 To arrayPartie.Count - 1
            localInstanceofMyClass = CType(arrayPartie(arraylistCnt - I - 1), clsCoupData)

            coup = localInstanceofMyClass.Coup
            If coup <> 0 Then
                RotSerie = localInstanceofMyClass.RS
                RotIntermit = localInstanceofMyClass.RI
                If RotSerie <> "" Then
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
                ElseIf RotIntermit <> "" Then
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

        cCoupData.RR = Rap

        If Rap <> "" Then
            cCoupData.RR7 = GetRapLast7(ColNum.eRR)
        End If
    End Sub

    Private Function GetRapLast7(ByVal SelektorCol As Integer) As String
        Dim Rap As String = ""
        Dim RapTmp As String = ""
        Dim count As Integer = 0
        Dim arraylistCnt As Integer = arrayPartie.Count
        Dim localInstanceofMyClass As clsCoupData

        If SelektorCol = ColNum.eTPR Then
            RapTmp = cCoupData.TPR
        ElseIf SelektorCol = ColNum.eTMR Then
            RapTmp = cCoupData.TmR
        ElseIf SelektorCol = ColNum.eSR Then
            RapTmp = cCoupData.SR
        ElseIf SelektorCol = ColNum.eRR Then
            RapTmp = cCoupData.RR
        End If

        If RapTmp <> "" Then
            count = count + 1
            Rap = RapTmp + Rap
        End If

        For I As Integer = 0 To arrayPartie.Count - 1
            localInstanceofMyClass = CType(arrayPartie(arraylistCnt - I - 1), clsCoupData)
            If SelektorCol = ColNum.eTPR Then
                RapTmp = localInstanceofMyClass.TPR
            ElseIf SelektorCol = ColNum.eTMR Then
                RapTmp = localInstanceofMyClass.TmR
            ElseIf SelektorCol = ColNum.eSR Then
                RapTmp = localInstanceofMyClass.SR
            ElseIf SelektorCol = ColNum.eRR Then
                RapTmp = localInstanceofMyClass.RR
            End If

            If RapTmp <> "" Then
                count = count + 1
                Rap = RapTmp + Rap
            End If

            If count = 7 Then Exit For
        Next

        If RadioButton3.Checked = True Or RadioButton6.Checked = True Then
            If count >= 2 Then
                GetRapLast7 = Rap
            Else
                GetRapLast7 = ""
            End If
        Else
            If count >= 3 Then
                GetRapLast7 = Rap
            Else
                GetRapLast7 = ""
            End If
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

        If SelektorCol = ColNum.eTPR7 Then
            RapTmp = cCoupData.TPR7
        ElseIf SelektorCol = ColNum.eTMR7 Then
            RapTmp = cCoupData.TmR7
        ElseIf SelektorCol = ColNum.eSR7 Then
            RapTmp = cCoupData.SR7
        ElseIf SelektorCol = ColNum.eRR7 Then
            RapTmp = cCoupData.RR7
        End If

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
            If RadioButton1.Checked = True Then
                If RapTmp.Length < 6 Then
                    FillSatz = ""
                    Exit Function
                End If

                ''Regel10 XXX(X) | OOO(O)
                If (s7 = s6 And s7 = s5) Then
                    If (s7 <> s4) Then
                        ReturnStr = "R10: " & s7 & s6 & s5 & "(" & s7 & ")"
                        SatzCount = SatzCount + 1
                        Call SetSatz(s7, SelektorCol, "R10")
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
                        Call SetSatz(s6, SelektorCol, "R11")
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
                        Call SetSatz(s7, SelektorCol, "R12")
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
                    Call SetSatz(s7, SelektorCol, "R13")
                    Exit Function
                    'XXOXX(O) | OOXOO(X)
                ElseIf (s7 = s6 And s6 <> s5 And s5 <> s4 And s4 = s3) Then
                    ReturnStr = "R13: " & s3 & s4 & s5 & s6 & s7 & "(" & s5 & ")"
                    SatzCount = SatzCount + 1
                    FillSatz = ReturnStr
                    Call SetSatz(s5, SelektorCol, "R13")
                    Exit Function
                End If
            ElseIf RadioButton2.Checked = True Then 'Biotop Original

                If RapTmp.Length >= 4 Then
                    ''Regel10 XXX(X) | OOO(O)
                    If (s7 = s6 And s7 = s5 And s4 <> s5 And CheckBox3.Checked = True) Then
                        ReturnStr = "R10: " & s5 & s6 & s7 & "(" & s7 & ")"
                        SatzCount = SatzCount + 1
                        Call SetSatz(s7, SelektorCol, "R10")
                        FillSatz = ReturnStr
                        Exit Function
                    End If
                End If

                If RapTmp.Length >= 4 Then
                    ''Regel11 OO XOX(O) | XX OXO(X) 
                    If (s7 <> s6 And s7 = s5 And s5 <> s4 And CheckBox4.Checked = True) Then
                        If RapTmp.Length = 4 Then
                            ReturnStr = "R11: " & s5 & s6 & s7 & "(" & s6 & ")"
                            SatzCount = SatzCount + 1
                            Call SetSatz(s6, SelektorCol, "R11")
                            FillSatz = ReturnStr
                            Exit Function
                        Else
                            If (s6 = s4 And s4 = s3) Then
                                ReturnStr = "R11: " & s5 & s6 & s7 & "(" & s6 & ")"
                                SatzCount = SatzCount + 1
                                Call SetSatz(s6, SelektorCol, "R11")
                                FillSatz = ReturnStr
                                Exit Function
                            End If
                        End If
                    End If
                End If

                If RapTmp.Length >= 4 Then
                    'OOOXOOX
                    ''Regel12 XXO(O) | OOX(X) XXOOX
                    If (s7 <> s6 And s6 = s5 And s5 <> s4 And CheckBox5.Checked = True) Then
                        If RapTmp.Length = 4 Or RapTmp.Length = 5 Then
                            ReturnStr = "R12: " & s5 & s6 & s7 & "(" & s7 & ")"
                            SatzCount = SatzCount + 1
                            Call SetSatz(s7, SelektorCol, "R12")
                            FillSatz = ReturnStr
                            Exit Function
                        End If

                        If RapTmp.Length = 5 Or RapTmp.Length = 6 Then
                            ReturnStr = "R12: " & s5 & s6 & s7 & "(" & s7 & ")"
                            SatzCount = SatzCount + 1
                            Call SetSatz(s7, SelektorCol, "R12")
                            FillSatz = ReturnStr
                            Exit Function
                        ElseIf RapTmp.Length >= 6 Then
                            If Not (s2 <> s3 And s3 = s4) Then
                                ReturnStr = "R12: " & s5 & s6 & s7 & "(" & s7 & ")"
                                SatzCount = SatzCount + 1
                                Call SetSatz(s7, SelektorCol, "R12")
                                FillSatz = ReturnStr
                                Exit Function
                            End If
                        End If
                    End If
                End If

                If RapTmp.Length >= 6 Then
                    ''If RapTmp.Length = 5 Then
                    'Regel13 OXXOX(X) | XOOXO(O)
                    If (s7 <> s6 And s6 <> s5 And s5 = s4 And s4 <> s3 And s2 <> s3 And CheckBox6.Checked = True) Then
                        ReturnStr = "R13: " & s3 & s4 & s5 & s6 & s7 & "(" & s7 & ")"
                        SatzCount = SatzCount + 1
                        Call SetSatz(s7, SelektorCol, "R13")
                        FillSatz = ReturnStr
                        Exit Function
                        'XXOXX(O) | OOXOO(X)
                    ElseIf (s7 = s6 And s6 <> s5 And s5 <> s4 And s4 = s3 And s2 <> s3) Then
                        ReturnStr = "R13: " & s3 & s4 & s5 & s6 & s7 & "(" & s5 & ")"
                        SatzCount = SatzCount + 1
                        Call SetSatz(s5, SelektorCol, "R13")
                        FillSatz = ReturnStr
                        Exit Function
                    End If
                End If

                FillSatz = ""
                Exit Function

            ElseIf RadioButton3.Checked = True Then

                If RapTmp.Length < 2 Then
                    FillSatz = ""
                    Exit Function
                End If

                If s7 <> s6 And RapTmp.Length >= 2 Then 'first try
                    ReturnStr = s6 & s7 & "(" & s6 & ")"
                    SatzCount = SatzCount + 1
                    FillSatz = ReturnStr
                    Call SetSatz(s6, SelektorCol)
                    Exit Function
                End If
                If s7 = s6 And s6 <> s5 And RapTmp.Length >= 3 Then 'first try
                    ReturnStr = s5 & s6 & s7 & "(" & s5 & ")"
                    SatzCount = SatzCount + 1
                    FillSatz = ReturnStr
                    Call SetSatz(s5, SelektorCol)
                    Exit Function
                End If
                If s7 = s6 And s6 = s5 And s5 <> s4 And RapTmp.Length >= 4 Then 'first try
                    ReturnStr = s4 & s5 & s6 & s7 & "(" & s4 & ")"
                    SatzCount = SatzCount + 1
                    FillSatz = ReturnStr
                    Call SetSatz(s4, SelektorCol)
                    Exit Function
                End If

            ElseIf RadioButton6.Checked = True Then

                If RapTmp.Length < 2 Then
                    FillSatz = ""
                    Exit Function
                End If

                If s7 <> s6 And RapTmp.Length >= 2 Then 'first try
                    ReturnStr = s6 & s7 & "(" & s6 & ")"
                    SatzCount = SatzCount + 1
                    FillSatz = ReturnStr
                    Call SetSatz(s6, SelektorCol)
                    Exit Function
                End If
                If s7 = s6 And s6 <> s5 And RapTmp.Length >= 3 Then 'first try
                    ReturnStr = s5 & s6 & s7 & "(" & s5 & ")"
                    SatzCount = SatzCount + 1
                    FillSatz = ReturnStr
                    Call SetSatz(s5, SelektorCol)
                    Exit Function
                End If
            End If
        End If

        FillSatz = ReturnStr

    End Function

    Private Sub SetSatz(ByVal s As String, ByVal col As Integer, Optional ByVal Regel As String = "")
        If s <> "" Then
            Select Case col
                Case ColNum.eTPR7
                    SatzPlus = s
                    SatzPlusRegel = Regel
                Case ColNum.eTMR7
                    SatzMinus = s
                    SatzMinusRegel = Regel
                Case ColNum.eSR7
                    SatzSchwarz = s
                    SatzSchwarzRegel = Regel
                Case ColNum.eRR7
                    SatzRot = s
                    SatzRotRegel = Regel
            End Select
        End If
    End Sub

    Private Function GetNextCol(ByRef pos As Integer, Optional ByVal AddCol As Boolean = True, Optional ByRef coup As Integer = 0) As Integer
        Dim i As Integer
        Dim localInstanceofMyClass As clsCoupData

        GetNextCol = 0

        localInstanceofMyClass = CType(arrayPartie(arrayPartie.Count - 1 - pos), clsCoupData)

        coup = localInstanceofMyClass.Coup
        If coup <> 0 Then
            GetNextCol = CheckCollong(coup)
        End If

        If AddCol = True Then pos = i + 1 'damit gleich der nächst coup ausgewählt ist
    End Function

    Private Function SetTransformatorRegel1() As Integer
        Dim IsCol As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim col1 As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim tmpcol As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer

        Dim IsBreaked As Boolean
        Dim LastcolFound As Boolean = False
        Dim pos As Integer = 0
        Dim coup As Integer
        Dim break As Boolean

        If arrayPartie.Count < 3 Then 'Regel 1 kann erst fühstens ab Coup 4 eintreten!
            SetTransformatorRegel1 = 0
            Exit Function
        End If

        For I As Integer = pos To arrayPartie.Count - 1
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
        For I As Integer = pos To arrayPartie.Count - 1
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
        Dim pos As Integer = 0
        Dim coup As Integer
        Dim break As Boolean

        If arrayPartie.Count < 5 Then 'Regel 2 kann erst fühstens ab Coup 6 eintreten!
            SetTransformatorRegel2 = 0
            Exit Function
        End If

        For I As Integer = pos To arrayPartie.Count - 1
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
        For I As Integer = pos To arrayPartie.Count - 1
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
        Dim pos As Integer = 0
        Dim coup As Integer
        Dim break As Boolean

        If arrayPartie.Count < 5 Then 'Regel 3 kann erst fühstens ab Coup 6 eintreten!
            SetTransformatorRegel3 = 0
            Exit Function
        End If

        For I As Integer = pos To arrayPartie.Count - 1
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
        For I As Integer = pos To arrayPartie.Count - 1
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
        Dim pos As Integer = 0
        Dim coup As Integer
        Dim break As Boolean

        If arrayPartie.Count < 6 Then 'Regel 4 kann erst fühstens ab Coup 7 eintreten!
            SetTransformatorRegel4 = 0
            Exit Function
        End If

        For I As Integer = pos To arrayPartie.Count - 1
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
        For I As Integer = pos To arrayPartie.Count - 1
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
        Dim pos As Integer = 0
        Dim coup As Integer

        If arrayPartie.Count < 3 Then 'Regel 5 kann erst fühstens ab Coup 4 eintreten!
            SetTransformatorRegel5 = 0
            Exit Function
        End If

        For I As Integer = pos To arrayPartie.Count - 1
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
        Dim pos As Integer = 0
        Dim coup As Integer
        Dim break As Boolean = False

        If arrayPartie.Count < 6 Then 'Regel 5_1 kann erst fühstens ab Coup 6 eintreten!
            SetTransformatorRegel5_1 = 0
            Exit Function
        End If

        For I As Integer = pos To arrayPartie.Count - 1
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

        For I As Integer = pos To arrayPartie.Count - 1
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
        If GetNextCol(0) = CheckCollong(cCoupData.Coup) Then
            CheckIsBreaked = False
        Else
            CheckIsBreaked = True
        End If
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
        If e.KeyCode = Keys.Return Then
            Auswerten()
        End If
    End Sub

    Private Sub TextBox1_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "" Then
            Button2.Enabled = False
        Else
            Button2.Enabled = True
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

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            Auswerten()
        End If
    End Sub

    Private Sub Auswerten()
        ToolStripStatusLabel1.Text = "Daten werden eingelesen und verarbeitet, bitte warten...."
        Me.Update()
        Me.Cursor = Cursors.WaitCursor

        clearFullData()
        SatzPlus = ""
        SatzMinus = ""
        SatzSchwarz = ""
        SatzRot = ""
        cCoupData = New clsCoupData()
        Call StartAuswertung()

        Call SetToolTipStatus()

        Me.TextRes.Text = ResTxt

        Me.TextRes.SelectionStart = TextBox1.TextLength
        Me.TextRes.ScrollToCaret()
        CreateGraph()
    End Sub

    Private Sub CreateGraph()

        ClearGraph()

        If zed Is Nothing Then
            zed = New ZedGraph.ZedGraphControl

            zed.Parent = Me.Panel2
            zed.Location = New Point(0, 0)
            zed.Size = New Size(Panel2.Width, Panel2.Height - 10)
            myPane = zed.GraphPane

            myPane.Title.Text = "Gesamt Saldo Verlauf"
            myPane.XAxis.Title.Text = "Satz"
            myPane.YAxis.Title.Text = "Saldo"
        End If

        Dim x As Integer = 0
        Dim y1 As Integer
        Dim list1 As New ZedGraph.PointPairList
        Dim CntS As Integer = 0
        Dim x3 As Integer = 0
        Dim y3 As Integer
        Dim list3 As New ZedGraph.PointPairList
        Dim CntS3 As Integer = 0


        If Not gesSaldoVerlaufhalbZero = Nothing Then
            Dim len As Integer = gesSaldoVerlaufhalbZero.Length
            If gesSaldoVerlaufhalbZero <> "" Then
                Do While CntS < len

                    If gesSaldoVerlaufhalbZero.Substring(CntS, 1) = "+" Then
                        y1 = y1 + 1
                        x = x + 1
                        list1.Add(x, y1)
                    ElseIf gesSaldoVerlaufhalbZero.Substring(CntS, 1) = "-" Then
                        y1 = y1 - 1
                        x = x + 1
                        list1.Add(x, y1)
                    End If
                    CntS = CntS + 1
                Loop
            End If

            Dim len3 As Integer = gesSaldoVerlaufhalbZero.Length
            If gesSaldoVerlaufhalbZero <> "" Then
                Do While CntS3 < len3
                    If gesSaldoVerlaufhalbZero.Substring(CntS3, 1) = "+" Then
                        y3 = y3 + 1
                        x3 = x3 + 1
                        list3.Add(x3, y3)
                    ElseIf gesSaldoVerlaufhalbZero.Substring(CntS3, 1) = "-" Then
                        y3 = y3 - 1
                        x3 = x3 + 1
                        list3.Add(x3, y3)
                    ElseIf gesSaldoVerlaufhalbZero.Substring(CntS3, 1) = "Z" Then
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

            myPane.Chart.Fill = New ZedGraph.Fill(Color.White, Color.LightGray, 90.0F)
            myPane.Fill = New ZedGraph.Fill(Color.AliceBlue)

            zed.AxisChange()
            zed.Size = New Size(Panel2.Width + (Panel2.Width * (HScrollBar1.Value / 10)) - 7, Panel2.Height - 7)
            zed.AxisChange()
            zed.Refresh()
        End If
    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        RowCount = 0

        Me.TextBox1.Focus()

        ToolStripStatusLabel1.Text = "Bitte Permanenz Datei für die System-Auswertung eingeben!"
        Me.Update()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        clearFullData()
    End Sub

    Private Sub clearFullData()
        ClearGraph()
        TextRes.Text = ""
        TextBox2.Text = ""
        ListBox1.Items.Clear()
        Me.ListBox1.Items.Add("Bitte Permanenz Datei für die System-Auswertung eingeben!")
    End Sub

    Private Sub ClearGraph()
        If Not myPane Is Nothing Then
            myPane.CurveList.Clear()
            myPane.GraphObjList.Clear()
            zed.Refresh()
        End If
    End Sub

    Private Sub HScrollBar1_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles HScrollBar1.Scroll
        If Not zed Is Nothing Then
            zed.Size = New Size(Panel2.Width + (Panel2.Width * (HScrollBar1.Value / 10)) - 7, Panel2.Height - 7)
            zed.AxisChange()
            zed.Refresh()
        End If
    End Sub

    Private Sub ListBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.Click

    End Sub

    Private Sub ListBox1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.DoubleClick
        Dim sItem As String = ListBox1.SelectedItem
        Dim SaldoVerlauf As String

        Me.Cursor = Cursors.WaitCursor

        If sItem = "" Then Exit Sub

        If CheckBox1.Checked = True Then
            SaldoVerlauf = gesSaldoVerlaufhalbZero
        Else
            Dim VerlPos As Integer = InStr(sItem, "Verlauf: ", CompareMethod.Text) + 8
            SaldoVerlauf = sItem.Substring(VerlPos, sItem.Length - VerlPos)
        End If

        Dim localArrayList As ArrayList = arrayGlobal(ListBox1.SelectedIndex)

        FormGrid.CoupData = localArrayList
        FormGrid.SaldoVerlaufhalbZero = SaldoVerlauf
        FormGrid.ShowDialog()
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim sItem As String = ListBox1.SelectedItem
        Dim SaldoVerlauf As String
        Dim tmpverlauf As String = ""

        Me.Cursor = Cursors.WaitCursor

        If sItem = "" Then Exit Sub

        If CheckBox1.Checked = True Then
            SaldoVerlauf = gesSaldoVerlaufhalbZero
        Else
            Dim VerlPos As Integer = InStr(sItem, "Verlauf: ", CompareMethod.Text) + 8
            SaldoVerlauf = sItem.Substring(VerlPos, sItem.Length - VerlPos)
        End If

        For I As Integer = 0 To SaldoVerlauf.Length - 1
            If SaldoVerlauf.Substring(I, 1) = "+" Then
                tmpverlauf = tmpverlauf & "+ "
            ElseIf SaldoVerlauf.Substring(I, 1) = "-" Then
                tmpverlauf = tmpverlauf & "- "
            End If
        Next

        Me.TextBox2.Text = tmpverlauf

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If RadioButton2.Checked = True Then
            CheckBox3.Enabled = True
            CheckBox4.Enabled = True
            CheckBox5.Enabled = True
            CheckBox6.Enabled = True
            CheckBox7.Enabled = True
            CheckBox8.Enabled = True
            CheckBox9.Enabled = True
            CheckBox10.Enabled = True
            CheckBox11.Enabled = True
            CheckBox3.Checked = True
            CheckBox4.Checked = True
            CheckBox5.Checked = True
            CheckBox6.Checked = True
            CheckBox7.Checked = True
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If RadioButton1.Checked = True Then
            CheckBox3.Enabled = True
            CheckBox4.Enabled = True
            CheckBox5.Enabled = True
            CheckBox6.Enabled = True
            CheckBox7.Enabled = True
            CheckBox8.Enabled = True
            CheckBox9.Enabled = True
            CheckBox10.Enabled = True
            CheckBox11.Enabled = True
            CheckBox3.Checked = True
            CheckBox4.Checked = True
            CheckBox5.Checked = True
            CheckBox6.Checked = True
            CheckBox7.Checked = True
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If RadioButton3.Checked = True Then
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            CheckBox6.Enabled = False
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
            CheckBox11.Enabled = False
            CheckBox3.Checked = False
            CheckBox4.Checked = False
            CheckBox5.Checked = False
            CheckBox6.Checked = False
            CheckBox7.Checked = False
        End If
    End Sub

    Private Sub RadioButton6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If RadioButton6.Checked = True Then
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            CheckBox6.Enabled = False
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
            CheckBox11.Enabled = False
            CheckBox3.Checked = False
            CheckBox4.Checked = False
            CheckBox5.Checked = False
            CheckBox6.Checked = False
            CheckBox7.Checked = False
        End If
    End Sub
End Class
