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
    Public myarraylist As New ArrayList

    'Private cCoupData As New clsCoupData("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
    Private cCoupData As clsCoupData

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

        



        ToolStripStatusLabel1.Text = "Bitte wählen Sie zum auswerten eine Permanenz Datei aus, oder geben Sie manuell Zahlen ein!"



        ToolStripProgressBar1.Visible = False


        loaded = True
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

        myarraylist.Clear()


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
                            
                            myarraylist.Clear()

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

        
        ToolStripProgressBar1.Visible = False
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub InsertNewCoup(ByVal NewCoup As String, Optional ByVal dummySet As Boolean = False)

        Dim R1 As DataGridViewRow
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

            Call SetTransformator(R1)

            'Call fillSelektorPlus()
            Call fillSelektorPlusX()

            'Call fillSelektorMinus()
            Call fillSelektorMinusX()

            'Call fillSelektorPlusRap()
            Call fillSelektorPlusRapX()

            'Call fillSelektorMinusRap()
            Call fillSelektorMinusRapX()

            'Call fillSelektorSchwarz()

            Call fillSelektorSchwarzX()
            'Call fillSelektorRot()
            Call fillSelektorRotX()

            'Call fillSelektorSchwarzRap()
            Call fillSelektorSchwarzRapX()
            'Call fillSelektorRotRap()
            Call fillSelektorRotRapX()

            If dummySet = True Then Exit Sub

            'R1.Cells(ColNum.eTPRS).Value = FillSatz(ColNum.eTPR7)
            cCoupData.TPRS = FillSatzX(ColNum.eTPR7)
            'R1.Cells(ColNum.eTMRS).Value = FillSatz(ColNum.eTMR7)
            cCoupData.TmRS = FillSatzX(ColNum.eTMR7)
            'R1.Cells(ColNum.eSRS).Value = FillSatz(ColNum.eSR7)
            cCoupData.SRS = FillSatzX(ColNum.eSR7)
            'R1.Cells(ColNum.eRRS).Value = FillSatz(ColNum.eRR7)
            cCoupData.RRS = FillSatzX(ColNum.eRR7)
        End If

        If dummySet = True Then Exit Sub

        'Call CanISetPlus()
        Call CanISetPlusX()
        'Call CanISetMinus()
        Call CanISetMinusX()
        'call CanISetSchwarz()
        Call CanISetSchwarzX()
        'Call CanISetRot()
        Call CanISetRotX()

        'Call CalkSaldo()
        Call CalkSaldoX()

        RowCount = RowCount + 1

        myarraylist.Add(cCoupData)
    End Sub

    Private Sub SetTransformator(ByRef R1 As DataGridViewRow)

        'Dim Erg As Integer = SetTransformatorRegel1()
        'If Erg = 1 Then
        '    R1.Cells(ColNum.eTP).Value = "+ (R1)"
        'ElseIf Erg = 2 Then
        '    R1.Cells(ColNum.eTM).Value = "- (R1)"
        'End If
        'If Erg = 0 Then
        '    Erg = SetTransformatorRegel2()
        '    If Erg = 1 Then
        '        R1.Cells(ColNum.eTP).Value = "+ (R2)"
        '    ElseIf Erg = 2 Then
        '        R1.Cells(ColNum.eTM).Value = "- (R2)"
        '    End If
        'End If
        'If Erg = 0 Then
        '    Erg = SetTransformatorRegel3()
        '    If Erg = 1 Then
        '        R1.Cells(ColNum.eTP).Value = "+ (R3)"
        '    ElseIf Erg = 2 Then
        '        R1.Cells(ColNum.eTM).Value = "- (R3)"
        '    End If
        'End If
        'If Erg = 0 Then
        '    Erg = SetTransformatorRegel4()
        '    If Erg = 1 Then
        '        R1.Cells(ColNum.eTP).Value = "+ (R4)"
        '    ElseIf Erg = 2 Then
        '        R1.Cells(ColNum.eTM).Value = "- (R4)"
        '    End If
        'End If
        'If Erg = 0 And RowCount > 3 Then
        '    Erg = SetTransformatorRegel5()
        '    If Erg = 1 Then
        '        R1.Cells(ColNum.eTP).Value = "+ (R5)"
        '    ElseIf Erg = 2 Then
        '        R1.Cells(ColNum.eTM).Value = "- (R5)"
        '    End If
        'End If
        'If Erg = 0 Then
        '    Erg = SetTransformatorRegel5_1()
        '    If Erg = 1 Then
        '        R1.Cells(ColNum.eTP).Value = "+ (R5)"
        '    ElseIf Erg = 2 Then
        '        R1.Cells(ColNum.eTM).Value = "- (R5)"
        '    End If
        'End If

        '***************************************************

        Dim Erg1 As Integer = SetTransformatorRegel1_2()
        If Erg1 = 1 Then
            cCoupData.TP = "+ (R1)"
        ElseIf Erg1 = 2 Then
            cCoupData.TM = "- (R1)"
        End If
        If Erg1 = 0 Then
            Erg1 = SetTransformatorRegel2_2()
            If Erg1 = 1 Then
                cCoupData.TP = "+ (R2)"
            ElseIf Erg1 = 2 Then
                cCoupData.TM = "- (R2)"
            End If
        End If
        If Erg1 = 0 Then
            Erg1 = SetTransformatorRegel3_2()
            If Erg1 = 1 Then
                cCoupData.TP = "+ (R3)"
            ElseIf Erg1 = 2 Then
                cCoupData.TM = "- (R3)"
            End If
        End If
        If Erg1 = 0 Then
            Erg1 = SetTransformatorRegel4_2()
            If Erg1 = 1 Then
                cCoupData.TP = "+ (R4)"
            ElseIf Erg1 = 2 Then
                cCoupData.TM = "- (R4)"
            End If
        End If
        If Erg1 = 0 And RowCount > 3 Then
            Erg1 = SetTransformatorRegel5_2()
            If Erg1 = 1 Then
                cCoupData.TP = "+ (R5)"
            ElseIf Erg1 = 2 Then
                cCoupData.TM = "- (R5)"
            End If
        End If
        If Erg1 = 0 Then
            Erg1 = SetTransformatorRegel5_3()
            If Erg1 = 1 Then
                cCoupData.TP = "+ (R5)"
            ElseIf Erg1 = 2 Then
                cCoupData.TM = "- (R5)"
            End If
        End If

    End Sub

    
    Private Sub CalkSaldoX()
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
        Dim arraylistCnt As Integer = myarraylist.Count


        If myarraylist.Count = 0 Then Exit Sub
        Dim localInstanceofMyClass = CType(myarraylist(arraylistCnt - 1), clsCoupData)
        tmpsaldo = saldo

        If localInstanceofMyClass.Ldf = 128 Then arraylistCnt = arraylistCnt

        stmp1 = localInstanceofMyClass.TPRS
        stmp2 = localInstanceofMyClass.TmRS
        stmp3 = localInstanceofMyClass.SRS
        stmp4 = localInstanceofMyClass.RRS

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
                        If cCoupData.TPR = "" Then
                            zeroCnt = zeroCnt + 1
                        End If

                        If SetTo.Substring(InStr(SetTo, "("), 1) = cCoupData.TPR Then
                            saldo = saldo + 1
                            localInstanceofMyClass.TPRS = SetTo & " +1/" & saldo
                            Maxloss = 0
                        Else
                            saldo = saldo - 1
                            Maxloss = Maxloss + 1
                            If cCoupData.TPR = "" Then
                                localInstanceofMyClass.TPRS = SetTo & " -0.5/" & saldo
                            Else
                                localInstanceofMyClass.TPRS = SetTo & " -1/" & saldo
                            End If
                        End If
                        SatzCnt = SatzCnt + 1
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
                        If cCoupData.TmR = "" Then
                            zeroCnt = zeroCnt + 1
                        End If

                        If SetTo.Substring(InStr(SetTo, "("), 1) = cCoupData.TmR Then
                            saldo = saldo + 1
                            localInstanceofMyClass.TmRS = SetTo & " +1/" & saldo
                            Maxloss = 0
                        Else
                            saldo = saldo - 1
                            Maxloss = Maxloss + 1
                            If cCoupData.TmR = "" Then
                                localInstanceofMyClass.TmRS = SetTo & " -0.5/" & saldo
                            Else
                                localInstanceofMyClass.TmRS = SetTo & " -1/" & saldo
                            End If
                        End If
                        SatzCnt = SatzCnt + 1
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
                        If cCoupData.SR = "" Then
                            zeroCnt = zeroCnt + 1
                        End If

                        If SetTo.Substring(InStr(SetTo, "("), 1) = cCoupData.SR Then
                            saldo = saldo + 1
                            localInstanceofMyClass.SRS = SetTo & " +1/" & saldo
                            Maxloss = 0
                        Else
                            saldo = saldo - 1
                            Maxloss = Maxloss + 1
                            If cCoupData.SR = "" Then
                                localInstanceofMyClass.SRS = SetTo & " -0.5/" & saldo
                            Else
                                localInstanceofMyClass.SRS = SetTo & " -1/" & saldo
                            End If
                        End If
                        SatzCnt = SatzCnt + 1
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
                        If cCoupData.RR = "" Then
                            zeroCnt = zeroCnt + 1
                        End If

                        If SetTo.Substring(InStr(SetTo, "("), 1) = cCoupData.RR Then
                            saldo = saldo + 1
                            localInstanceofMyClass.RRS = SetTo & " +1/" & saldo
                            Maxloss = 0
                        Else
                            saldo = saldo - 1
                            Maxloss = Maxloss + 1
                            If cCoupData.RR = "" Then
                                localInstanceofMyClass.RRS = SetTo & " -0.5/" & saldo
                            Else
                                localInstanceofMyClass.RRS = SetTo & " -1/" & saldo
                            End If
                        End If
                        SatzCnt = SatzCnt + 1
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

    Private Sub CanISetPlusX()
        Dim Rap As String
        Dim Ser As String
        Dim Int As String
        Dim SerInt As String = ""
        Dim SerIntCnt As Integer = 0
        Dim sAkt As String = ""
        Dim sLast As String = ""
        Dim coup As String = ""
        Dim arraylistCnt As Integer = myarraylist.Count
        Dim localInstanceofMyClass As clsCoupData

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

            For I As Integer = 0 To myarraylist.Count - 1
                localInstanceofMyClass = CType(myarraylist(arraylistCnt - I - 1), clsCoupData)
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
                    localInstanceofMyClass = CType(myarraylist(arraylistCnt - I), clsCoupData)
                    coup = localInstanceofMyClass.Coup
                    sLast = localInstanceofMyClass.TP

                    If coup <> "00" And coup <> "0" Then
                        If sAkt <> "" And sLast = "" Then 'And coup <> "0" Then                            
                            Dim tmpCopyStruct As clsCoupData = cCoupData
                            myarraylist.Add(cCoupData) 'momentaner Datensatz in array
                            InsertNewCoup("1", True)
                            If cCoupData.TPR = SatzPlus Then
                                tmpCopyStruct.TPRS = "Satz auf (" & SatzPlus & "/R)"
                            Else
                                tmpCopyStruct.TPRS = "Satz auf (" & SatzPlus & "/S)"
                            End If

                            DeletelastDummyCoupX()
                            ''zurückschreiben
                            cCoupData = tmpCopyStruct
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

    Private Sub DeletelastDummyCoupX()
        Dim ArraylistCnt As Integer = myarraylist.Count
        myarraylist.Remove(myarraylist(ArraylistCnt - 1))
    End Sub


    Private Sub CanISetMinusX()
        Dim Rap As String
        Dim Ser As String
        Dim Int As String
        Dim SerInt As String = ""
        Dim SerIntCnt As Integer = 0
        Dim sAkt As String = ""
        Dim sLast As String = ""
        Dim coup As String = ""
        Dim arraylistCnt As Integer = myarraylist.Count
        Dim localInstanceofMyClass As clsCoupData

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

            For I As Integer = 0 To myarraylist.Count - 1
                localInstanceofMyClass = CType(myarraylist(arraylistCnt - I - 1), clsCoupData)
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
                    localInstanceofMyClass = CType(myarraylist(arraylistCnt - I), clsCoupData)
                    coup = localInstanceofMyClass.Coup
                    sLast = localInstanceofMyClass.TM

                    If coup <> "00" And coup <> "0" Then
                        If sAkt <> "" And sLast = "" Then 'And coup <> "0" Then                            
                            Dim tmpCopyStruct As clsCoupData = cCoupData
                            myarraylist.Add(cCoupData) 'momentaner Datensatz in array
                            InsertNewCoup("1", True)
                            If cCoupData.TmR = SatzMinus Then
                                tmpCopyStruct.TmRS = "Satz auf (" & SatzMinus & "/R)"
                            Else
                                tmpCopyStruct.TmRS = "Satz auf (" & SatzMinus & "/S)"
                            End If

                            DeletelastDummyCoupX()
                            ''zurückschreiben
                            cCoupData = tmpCopyStruct
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

    Private Sub CanISetSchwarzX()
        Dim Rap As String
        Dim Ser As String
        Dim Int As String
        Dim SerInt As String = ""
        Dim SerIntCnt As Integer = 0
        Dim sAkt As String = ""
        Dim sLast As String = ""
        Dim coup As String = ""
        Dim arraylistCnt As Integer = myarraylist.Count
        Dim localInstanceofMyClass As clsCoupData

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

            For I As Integer = 0 To myarraylist.Count - 1
                localInstanceofMyClass = CType(myarraylist(arraylistCnt - I - 1), clsCoupData)
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
                    localInstanceofMyClass = CType(myarraylist(arraylistCnt - I), clsCoupData)
                    coup = localInstanceofMyClass.Coup
                    sLast = localInstanceofMyClass.S

                    If coup <> "00" And coup <> "0" Then
                        If sAkt <> "" And sLast = "" Then 'And coup <> "0" Then                            
                            Dim tmpCopyStruct As clsCoupData = cCoupData
                            myarraylist.Add(cCoupData) 'momentaner Datensatz in array
                            InsertNewCoup("1", True)
                            If cCoupData.SR = SatzSchwarz Then
                                tmpCopyStruct.SRS = "Satz auf (" & SatzSchwarz & "/R)"
                            Else
                                tmpCopyStruct.SRS = "Satz auf (" & SatzSchwarz & "/S)"
                            End If

                            DeletelastDummyCoupX()
                            ''zurückschreiben
                            cCoupData = tmpCopyStruct
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

    Private Sub CanISetRotX()
        Dim Rap As String
        Dim Ser As String
        Dim Int As String
        Dim SerInt As String = ""
        Dim SerIntCnt As Integer = 0
        Dim sAkt As String = ""
        Dim sLast As String = ""
        Dim coup As String = ""
        Dim arraylistCnt As Integer = myarraylist.Count
        Dim localInstanceofMyClass As clsCoupData

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

            For I As Integer = 0 To myarraylist.Count - 1
                localInstanceofMyClass = CType(myarraylist(arraylistCnt - I - 1), clsCoupData)
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
                    localInstanceofMyClass = CType(myarraylist(arraylistCnt - I), clsCoupData)
                    coup = localInstanceofMyClass.Coup
                    sLast = localInstanceofMyClass.R

                    If coup <> "00" And coup <> "0" Then
                        If sAkt <> "" And sLast = "" Then 'And coup <> "0" Then                            
                            Dim tmpCopyStruct As clsCoupData = cCoupData
                            myarraylist.Add(cCoupData) 'momentaner Datensatz in array
                            InsertNewCoup("1", True)
                            If cCoupData.RR = SatzRot Then
                                tmpCopyStruct.RRS = "Satz auf (" & SatzRot & "/R)"
                            Else
                                tmpCopyStruct.RRS = "Satz auf (" & SatzRot & "/S)"
                            End If

                            DeletelastDummyCoupX()
                            ''zurückschreiben
                            cCoupData = tmpCopyStruct
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

    Private Sub fillSelektorPlusX(Optional ByVal fiktiv As Boolean = False)

        Dim Plus1 As String = ""
        Dim Plus2 As String = ""
        Dim Plus3 As String = ""
        Dim PlusTmp As String = ""
        Dim countPlus As Integer = 1
        Dim coup As Integer
        Dim arraylistCnt As Integer = myarraylist.Count
        Dim localInstanceofMyClass As clsCoupData

        coup = cCoupData.Coup
        If coup <> 0 Then
            Plus1 = cCoupData.TP
        End If

        If arraylistCnt = 0 Then Exit Sub

        For I As Integer = 0 To myarraylist.Count - 1
            localInstanceofMyClass = CType(myarraylist(arraylistCnt - I - 1), clsCoupData)
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

    Private Sub fillSelektorMinusX(Optional ByVal fiktiv As Boolean = False)

        Dim Minus1 As String = ""
        Dim Minus2 As String = ""
        Dim Minus3 As String = ""
        Dim MinusTmp As String = ""
        Dim countMinus As Integer = 1
        Dim coup As Integer
        Dim arraylistCnt As Integer = myarraylist.Count
        Dim localInstanceofMyClass As clsCoupData

        coup = cCoupData.Coup
        If coup <> 0 Then
            Minus1 = cCoupData.TM
        End If

        If arraylistCnt = 0 Then Exit Sub

        For I As Integer = 0 To myarraylist.Count - 1
            localInstanceofMyClass = CType(myarraylist(arraylistCnt - I - 1), clsCoupData)
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

    Private Sub fillSelektorSchwarzX(Optional ByVal fiktiv As Boolean = False)

        Dim Schwarz1 As String = ""
        Dim Schwarz2 As String = ""
        Dim Schwarz3 As String = ""
        Dim SchwarzTmp As String = ""
        Dim countSchwarz As Integer = 1
        Dim coup As Integer
        Dim arraylistCnt As Integer = myarraylist.Count
        Dim localInstanceofMyClass As clsCoupData

        coup = cCoupData.Coup
        If coup <> 0 Then
            Schwarz1 = cCoupData.S
        End If

        If arraylistCnt = 0 Then Exit Sub

        For I As Integer = 0 To myarraylist.Count - 1
            localInstanceofMyClass = CType(myarraylist(arraylistCnt - I - 1), clsCoupData)
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

    Private Sub fillSelektorRotX(Optional ByVal fiktiv As Boolean = False)

        Dim Rot1 As String = ""
        Dim Rot2 As String = ""
        Dim Rot3 As String = ""
        Dim RotTmp As String = ""
        Dim countRot As Integer = 1
        Dim coup As Integer
        Dim arraylistCnt As Integer = myarraylist.Count
        Dim localInstanceofMyClass As clsCoupData

        coup = cCoupData.Coup
        If coup <> 0 Then
            Rot1 = cCoupData.R
        End If

        If arraylistCnt = 0 Then Exit Sub

        For I As Integer = 0 To myarraylist.Count - 1
            localInstanceofMyClass = CType(myarraylist(arraylistCnt - I - 1), clsCoupData)
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

    Private Sub fillSelektorPlusRapX()

        Dim PlusSerie As String = ""
        Dim PlusIntermit As String = ""
        Dim PlusRap1 As String = ""
        Dim PlusRap2 As String = ""
        Dim PlusRap3 As String = ""
        Dim PlusRap4 As String = ""
        Dim PlusSum As String = ""
        Dim coup As Integer
        Dim Rap As String = ""
        Dim arraylistCnt As Integer = myarraylist.Count
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

        For I As Integer = 0 To myarraylist.Count - 1
            localInstanceofMyClass = CType(myarraylist(arraylistCnt - I - 1), clsCoupData)

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
            cCoupData.TPR7 = GetRapLast7X(ColNum.eTPR)
        End If
    End Sub

    Private Sub fillSelektorMinusRapX()

        Dim MinusSerie As String = ""
        Dim MinusIntermit As String = ""
        Dim MinusRap1 As String = ""
        Dim MinusRap2 As String = ""
        Dim MinusRap3 As String = ""
        Dim MinusRap4 As String = ""
        Dim MinusSum As String = ""
        Dim coup As Integer
        Dim Rap As String = ""
        Dim arraylistCnt As Integer = myarraylist.Count
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

        For I As Integer = 0 To myarraylist.Count - 1
            localInstanceofMyClass = CType(myarraylist(arraylistCnt - I - 1), clsCoupData)

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
            cCoupData.TmR7 = GetRapLast7X(ColNum.eTMR)
        End If
    End Sub

    Private Sub fillSelektorSchwarzRapX()

        Dim SchwarzSerie As String = ""
        Dim SchwarzIntermit As String = ""
        Dim SchwarzRap1 As String = ""
        Dim SchwarzRap2 As String = ""
        Dim SchwarzRap3 As String = ""
        Dim SchwarzRap4 As String = ""
        Dim SchwarzSum As String = ""
        Dim coup As Integer
        Dim Rap As String = ""
        Dim arraylistCnt As Integer = myarraylist.Count
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

        For I As Integer = 0 To myarraylist.Count - 1
            localInstanceofMyClass = CType(myarraylist(arraylistCnt - I - 1), clsCoupData)

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
            cCoupData.SR7 = GetRapLast7X(ColNum.eSR)
        End If
    End Sub

    Private Sub fillSelektorRotRapX()

        Dim RotSerie As String = ""
        Dim RotIntermit As String = ""
        Dim RotRap1 As String = ""
        Dim RotRap2 As String = ""
        Dim RotRap3 As String = ""
        Dim RotRap4 As String = ""
        Dim RotSum As String = ""
        Dim coup As Integer
        Dim Rap As String = ""
        Dim arraylistCnt As Integer = myarraylist.Count
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

        For I As Integer = 0 To myarraylist.Count - 1
            localInstanceofMyClass = CType(myarraylist(arraylistCnt - I - 1), clsCoupData)

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
            cCoupData.RR7 = GetRapLast7X(ColNum.eRR)
        End If
    End Sub

    Private Function GetRapLast7X(ByVal SelektorCol As Integer) As String
        Dim Rap As String = ""
        Dim RapTmp As String = ""
        Dim count As Integer = 0
        Dim arraylistCnt As Integer = myarraylist.Count
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

        For I As Integer = 0 To myarraylist.Count - 1
            localInstanceofMyClass = CType(myarraylist(arraylistCnt - I - 1), clsCoupData)
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

        If count >= 3 Then
            GetRapLast7X = Rap
        Else
            GetRapLast7X = ""
        End If
    End Function

    Private Function FillSatzX(ByVal SelektorCol As Integer) As String
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

            If RapTmp.Length < 6 Then
                FillSatzX = ""
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
                FillSatzX = ReturnStr
                Exit Function
            End If

            ''Is R10 aborted?
            If s3 <> "" And s4 <> "" And s5 <> "" Then
                If s3 = s4 And s4 = s5 Then
                    FillSatzX = ""
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

                FillSatzX = ReturnStr
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

                FillSatzX = ReturnStr
                Exit Function
            End If

            ''If RapTmp.Length = 5 Then
            'Regel13 OXXOX(X) | XOOXO(O)
            If (s7 <> s6 And s6 <> s5 And s5 = s4 And s4 <> s3) Then
                ReturnStr = "R13: " & s3 & s4 & s5 & s6 & s7 & "(" & s7 & ")"
                SatzCount = SatzCount + 1
                FillSatzX = ReturnStr
                Call SetSatz(s7, SelektorCol)
                Exit Function
                'XXOXX(O) | OOXOO(X)
            ElseIf (s7 = s6 And s6 <> s5 And s5 <> s4 And s4 = s3) Then
                ReturnStr = "R13: " & s3 & s4 & s5 & s6 & s7 & "(" & s5 & ")"
                SatzCount = SatzCount + 1
                FillSatzX = ReturnStr
                Call SetSatz(s5, SelektorCol)
                Exit Function
            End If
        End If

        FillSatzX = ReturnStr

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

    Private Function GetNextColX(ByRef pos As Integer, Optional ByVal AddCol As Boolean = True, Optional ByRef coup As Integer = 0) As Integer
        Dim i As Integer
        Dim localInstanceofMyClass As clsCoupData

        GetNextColX = 0

        localInstanceofMyClass = CType(myarraylist(myarraylist.Count - 1 - pos), clsCoupData)

        coup = localInstanceofMyClass.Coup
        If coup <> 0 Then
            GetNextColX = CheckCollong(coup)
        End If

        If AddCol = True Then pos = i + 1 'damit gleich der nächst coup ausgewählt ist
    End Function

    Private Function SetTransformatorRegel1_2() As Integer
        Dim IsCol As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim col1 As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim tmpcol As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer

        Dim IsBreaked As Boolean
        Dim LastcolFound As Boolean = False
        Dim pos As Integer = 0
        Dim coup As Integer
        Dim break As Boolean

        If myarraylist.Count < 3 Then 'Regel 1 kann erst fühstens ab Coup 4 eintreten!
            SetTransformatorRegel1_2 = 0
            Exit Function
        End If

        For I As Integer = pos To myarraylist.Count - 1
            tmpcol = GetNextColX(I, False, coup)
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
            SetTransformatorRegel1_2 = 0
            Exit Function
        End If

        IsBreaked = CheckIsBreakedX()
        For I As Integer = pos To myarraylist.Count - 1
            tmpcol = GetNextColX(I, False, coup)
            If LastcolFound = True Then
                If tmpcol <> 0 Then
                    If tmpcol <> IsCol Then
                        'Regel 1: Vorherige Farbe war Intermittenz: 
                        'aktuelle Farbe abgebrochen -> +
                        'aktuelle Farbe weitergelaufen -> -
                        If IsBreaked = True Then
                            SetTransformatorRegel1_2 = 1 '+
                        Else
                            SetTransformatorRegel1_2 = 2 '-
                        End If
                        Exit Function
                    Else
                        'Regel 1: nicht in Kraft getreten, Vorherige Farbe war keine Intermittenz
                        SetTransformatorRegel1_2 = 0
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

    Private Function SetTransformatorRegel2_2() As Integer
        Dim IsCol As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim col1 As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim col2 As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim tmpcol As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer

        Dim IsBreaked As Boolean
        Dim LastcolFound As Boolean = False
        Dim pos As Integer = 0
        Dim coup As Integer
        Dim break As Boolean

        If myarraylist.Count < 5 Then 'Regel 2 kann erst fühstens ab Coup 6 eintreten!
            SetTransformatorRegel2_2 = 0
            Exit Function
        End If

        For I As Integer = pos To myarraylist.Count - 1
            tmpcol = GetNextColX(I, False, coup)
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
            SetTransformatorRegel2_2 = 0
            Exit Function
        End If

        IsBreaked = CheckIsBreakedX()
        For I As Integer = pos To myarraylist.Count - 1
            tmpcol = GetNextColX(I, False, coup)
            If LastcolFound = True Then
                If tmpcol <> 0 Then
                    If tmpcol <> IsCol Then
                        'Regel 1: Vorherige Farbe war Intermittenz: 
                        'aktuelle Farbe abgebrochen -> +
                        'aktuelle Farbe weitergelaufen -> -
                        If IsBreaked = True Then
                            SetTransformatorRegel2_2 = 1 '+
                        Else
                            SetTransformatorRegel2_2 = 2 '-
                        End If
                        Exit Function
                    Else
                        'Regel 1: nicht in Kraft getreten, Vorherige Farbe war keine Intermittenz
                        SetTransformatorRegel2_2 = 0
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

    Private Function SetTransformatorRegel3_2() As Integer
        Dim IsCol As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim col1 As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim col2 As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim tmpcol As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer

        Dim IsBreaked As Boolean
        Dim LastcolFound As Boolean = False
        Dim pos As Integer = 0
        Dim coup As Integer
        Dim break As Boolean

        If myarraylist.Count < 5 Then 'Regel 3 kann erst fühstens ab Coup 6 eintreten!
            SetTransformatorRegel3_2 = 0
            Exit Function
        End If

        For I As Integer = pos To myarraylist.Count - 1
            tmpcol = GetNextColX(I, False, coup)
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
            SetTransformatorRegel3_2 = 0
            Exit Function
        End If

        col1 = 0

        IsBreaked = CheckIsBreakedX()
        For I As Integer = pos To myarraylist.Count - 1
            tmpcol = GetNextColX(I, False, coup)
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
                                SetTransformatorRegel3_2 = 2 '-
                            Else
                                SetTransformatorRegel3_2 = 1 '+
                            End If
                            Exit Function
                        Else
                            'Regel 1: nicht in Kraft getreten, Vorherige Farbe war keine Intermittenz
                            SetTransformatorRegel3_2 = 0
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

    Private Function SetTransformatorRegel4_2() As Integer
        Dim IsCol As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim col1 As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim col2 As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim tmpcol As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer

        Dim IsBreaked As Boolean
        Dim LastcolFound As Boolean = False
        Dim pos As Integer = 0
        Dim coup As Integer
        Dim break As Boolean

        If myarraylist.Count < 6 Then 'Regel 4 kann erst fühstens ab Coup 7 eintreten!
            SetTransformatorRegel4_2 = 0
            Exit Function
        End If

        For I As Integer = pos To myarraylist.Count - 1
            tmpcol = GetNextColX(I, False, coup)
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
            SetTransformatorRegel4_2 = 0
            Exit Function
        End If

        col1 = 0

        IsBreaked = CheckIsBreakedX()
        For I As Integer = pos To myarraylist.Count - 1
            tmpcol = GetNextColX(I, False, coup)
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
                                SetTransformatorRegel4_2 = 1 '+
                            Else
                                SetTransformatorRegel4_2 = 2 '-
                            End If
                            Exit Function
                        Else
                            'Regel 1: nicht in Kraft getreten, Vorherige Farbe war keine Intermittenz
                            SetTransformatorRegel4_2 = 0
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

    Private Function SetTransformatorRegel5_2() As Integer
        Dim IsCol As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim col1 As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim col2 As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim tmpcol As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer

        Dim IsBreaked As Boolean
        Dim pos As Integer = 0
        Dim coup As Integer

        If myarraylist.Count < 3 Then 'Regel 5 kann erst fühstens ab Coup 4 eintreten!
            SetTransformatorRegel5_2 = 0
            Exit Function
        End If

        For I As Integer = pos To myarraylist.Count - 1
            tmpcol = GetNextColX(I, False, coup)
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
            SetTransformatorRegel5_2 = 0
            Exit Function
        End If

        col1 = 0

        IsBreaked = CheckIsBreakedX()
        If IsBreaked = True Then
            SetTransformatorRegel5_2 = 2 '-
        Else
            SetTransformatorRegel5_2 = 1 '+
        End If

    End Function

    Private Function SetTransformatorRegel5_3() As Integer
        Dim IsCol As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim col1 As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim col2 As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer
        Dim tmpcol As Integer = 0 '1 = schwarz, 2= rot, 0 = zero/noch leer

        Dim IsBreaked As Boolean
        Dim LastcolFound As Boolean = False
        Dim pos As Integer = 0
        Dim coup As Integer
        Dim break As Boolean = False

        If myarraylist.Count < 6 Then 'Regel 5_1 kann erst fühstens ab Coup 6 eintreten!
            SetTransformatorRegel5_3 = 0
            Exit Function
        End If

        For I As Integer = pos To myarraylist.Count - 1
            tmpcol = GetNextColX(I, False, coup)
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
        IsBreaked = CheckIsBreakedX()

        For I As Integer = pos To myarraylist.Count - 1
            tmpcol = GetNextColX(I, False, coup)
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
                                SetTransformatorRegel5_3 = 2 '-
                            Else
                                SetTransformatorRegel5_3 = 1 '+
                            End If
                            Exit Function
                        Else
                            'Regel 1: nicht in Kraft getreten, Vorherige Farbe war keine Intermittenz
                            SetTransformatorRegel5_3 = 0
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

    Private Function CheckIsBreakedX()
        If GetNextColX(0) = CheckCollong(cCoupData.Coup) Then
            CheckIsBreakedX = False
        Else
            CheckIsBreakedX = True
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

    'Private Sub DataGridView1_ColumnWidthChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs)
    '    If loaded = True Then Call SetSizeLabelColPos()
    'End Sub


    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Call InsertManuell()
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
        cCoupData = New clsCoupData()
        Call StartAuswertung()

        Call SetToolTipStatus()

        FormGrid.CoupData = myarraylist
        FormGrid.ShowDialog()
        FormRes.prop1() = ResTxt
        FormRes.SaldoVerlaufhalbZero = gesSaldoVerlaufhalbZero
        Call FormRes.ShowDialog()
        Me.Update()
    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
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
