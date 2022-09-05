
Imports System.Configuration
Imports System.Text

Public Class frmMain

#Region " Declarations "
    Private myCommSvr As CommSvr
    Private myColMessages As Collection
    Private myMiniColMessages As Collection
    Private myIntensity As String
    Private myInsideIntensity As String
    Private myOutsideIntensity As String
    Private myToteCompany As Integer
    Private myOfficialTimeStart As Date
    Private myOfficialTimeSpan As Integer
    Private p_intMTPToClear As Integer = 10
    Private m_Ctr As String = "F5"
    Private myComPort As ComPort.Port
    'Private myMiniComPort As ComPort.Port
    Private myTrackCondition As String = "" 'TrackCondition
    Private myCurrentExotic As CurrentExotic
    Private p_blnUpdateManually As Boolean = True
    Private m_intCtr As Integer = 0
    Private m_FlagExacta As Boolean = True
    Private m_intCycleExotic As Integer 'used to cycle teh exotics on teh ASCII strip.
    Private m_FlagChangeEvent As Boolean = True

    Private p_intMaxNumbOfExotics As Integer = 10
    Private results_secondpage As Boolean
    Private results_thirdpage As Boolean
    Private results_forthpage As Boolean
    Private results_fivepage As Boolean
    Private results_sixpage As Boolean
    Private results_sevenpage As Boolean
    Private results_eightpage As Boolean
    Private results_ninepage As Boolean
    Private results_tenpage As Boolean
    Private results_num(p_intMaxNumbOfExotics + 1) As String
    Private results_amt(p_intMaxNumbOfExotics + 1) As String
    Private results_type(p_intMaxNumbOfExotics + 1) As String
    Private results_num_ctr As Integer
    Private p_intExoCtr As Integer = 1 'One exotic at the time

#End Region

#Region " Enums "
    Private Enum TrackCondition
        Fast = 1
        Slow = 2
        Sloppy = 3
    End Enum

    Private Enum CurrentExotic
        None = 0
        Exacta = 1
        Trifecta = 2
    End Enum

    Private Enum PayloadType
        Odds = 11
        RunningOrder = 12
        WPS = 15
        Timing = 16
        allPayLoads = 20
    End Enum

#End Region

#Region " Properties "

    Private ReadOnly Property HasInsideBoard() As Boolean
        Get
            Return My.Settings.HasInsideBoard
        End Get
    End Property

    Private ReadOnly Property HasMiniBoard() As Boolean
        Get
            Return My.Settings.HasMiniBoard
        End Get
    End Property

#End Region

#Region " Digital Data Arrays "
    Private strCollectionOn As Collection
    Private strCollectionOff As Collection
    Private strMiniBoardCollectionOn As Collection
    Private strMiniBoardCollectionOff As Collection

    Private strToteDisplayCollection As Collection
    Private intSegmentHexPositionCollection As Collection

    '
#End Region

#Region " Constructor "

    Public Sub New()
        Try
            ' This call is required by the Windows Form Designer.
            InitializeComponent()
            'kill RSIPort from task manager
            Me.KillComSvr()
            Me.KillDataSvr()
            'Me.myCommSvr = New CommSvr

            'Comm Port to use will always be the second parameter
            Dim args As String()
            args = Environment.GetCommandLineArgs

            Dim comPort As Integer
            Dim comPortParam = args.FirstOrDefault(Function(a) a.ToLower().IndexOf("toteboardportnumber=") >= 0)

            If String.IsNullOrWhiteSpace(comPortParam) Then
                comPort = 2 'Default to comPort2 for tote board
            Else
                If Not Integer.TryParse(comPortParam.Substring(20), comPort) Then
                    Throw New InvalidOperationException("The ToteBoardPortNumber has an invalid value")
                End If
            End If


            Me.InitializePorts(comPort)
            Me.myColMessages = New Collection
            Me.myMiniColMessages = New Collection
            'Me.RaceDisplayDataset = Me.myCommSvr.Dataset
            Me.myOfficialTimeStart = Date.MinValue
            'Me.myTrackCondition = TrackCondition.Fast
            '
            'AddHandler Me.myCommSvr.DisplayNewRace, AddressOf myCommSvr_DisplayNewRace
            'AddHandler Me.myCommSvr.DisplayNewOdds, AddressOf myCommSvr_DisplayNewOdds
            'AddHandler Me.myCommSvr.DisplayMTP, AddressOf myCommSvr_DisplayMTP
            'AddHandler Me.myCommSvr.DisplayTrackCondition, AddressOf myCommSvr_DisplayTrackCondition
            'AddHandler Me.myCommSvr.DisplayRaceStatus, AddressOf myCommSvr_DisplayRaceStatus
            'AddHandler Me.myCommSvr.DisplayRaceStatusJM, AddressOf myCommSvr_DisplayRaceStatusJM
            'AddHandler Me.myCommSvr.DisplayPostTime, AddressOf myCommSvr_DisplayPostTime
            'AddHandler Me.myCommSvr.DisplayTOD, AddressOf myCommSvr_DisplayTOD
            'AddHandler Me.myCommSvr.DisplayRunningOrder, AddressOf myCommSvr_DisplayRunningOrder
            'AddHandler Me.myCommSvr.DisplayRunningOrderAmtote, AddressOf myCommSvr_DisplayRunningOrderAmtote
            'AddHandler Me.myCommSvr.DisplayNewResults, AddressOf myCommSvr_DisplayNewResults
            ''AddHandler Me.myCommSvr.DisplayExotics, AddressOf myCommSvr_DisplayExotics
            'AddHandler Me.myCommSvr.DisplayTiming, AddressOf myCommSvr_DisplayTiming
            'AddHandler Me.myCommSvr.DisplayTeletimer, AddressOf myCommSvr_DisplayTeletimer
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        ''
    End Sub

#End Region

#Region " Form Events "

    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'myCommSvr.m_ApplicationBussy = True
        frmSplash.ShowSplash()
        For i As Integer = 0 To 3500000
            Application.DoEvents()
        Next
        '
        frmSplash.CloseForm()
        'check if there's an instace of the appl open
        Me.CheckForExistingInstance()
        '
        Me.SetTransparent()
        'default radio buttons unchecked
        Me.UncheckRadioButtons()
        '
        Me.DisableAllTextBoxes()
        'prepare digital collections and arrays
        Me.PrepareDigitalData()
        Me.PrepareMiniBoardDigitalData()
        Me.PrepareToteDisplayControler()
        Me.PrepareSegmentHexPosition()
        'populate intensity, set it to 22 default
        Me.SetAndPopulateIntensity()
        '        Me.GetToteCompany()
        '
        Me.SetAndPopulateOfficialTime()
        '
        Me.cboMTPToClearRS.Text = Me.p_intMTPToClear
        '
        Me.myCurrentExotic = CurrentExotic.None
        tmrTest.Enabled = False
        'myCommSvr.m_ApplicationBussy = False
        '
        ''''
        'TODO: *** Make skipNewOfficialStatusEvent CommSvr parameter configurable ***

        Dim skipOfficialEevent As Boolean = True
        Dim skipNewOfficialStatusEvent As String = ConfigurationManager.AppSettings("skipNewOfficialStatusEvent")
        If Not String.IsNullOrWhiteSpace(skipNewOfficialStatusEvent) Then
            Dim configSkip As Boolean
            If Boolean.TryParse(skipNewOfficialStatusEvent, configSkip) Then
                skipOfficialEevent = configSkip
            End If
        End If
        Me.myCommSvr = New CommSvr(skipOfficialEevent)

        Me.RaceDisplayDataset = Me.myCommSvr.Dataset
        AddHandler Me.myCommSvr.DisplayNewRace, AddressOf myCommSvr_DisplayNewRace
        AddHandler Me.myCommSvr.DisplayNewOdds, AddressOf myCommSvr_DisplayNewOdds
        AddHandler Me.myCommSvr.DisplayMTP, AddressOf myCommSvr_DisplayMTP
        AddHandler Me.myCommSvr.DisplayTrackCondition, AddressOf myCommSvr_DisplayTrackCondition
        AddHandler Me.myCommSvr.DisplayRaceStatus, AddressOf myCommSvr_DisplayRaceStatus
        AddHandler Me.myCommSvr.DisplayRaceStatusJM, AddressOf myCommSvr_DisplayRaceStatusJM
        AddHandler Me.myCommSvr.DisplayPostTime, AddressOf myCommSvr_DisplayPostTime
        AddHandler Me.myCommSvr.DisplayTOD, AddressOf myCommSvr_DisplayTOD
        AddHandler Me.myCommSvr.DisplayRunningOrder, AddressOf myCommSvr_DisplayRunningOrder
        AddHandler Me.myCommSvr.DisplayRunningOrderAmtote, AddressOf myCommSvr_DisplayRunningOrderAmtote
        AddHandler Me.myCommSvr.DisplayNewResults, AddressOf myCommSvr_DisplayNewResults
        'AddHandler Me.myCommSvr.DisplayExotics, AddressOf myCommSvr_DisplayExotics
        AddHandler Me.myCommSvr.DisplayTiming, AddressOf myCommSvr_DisplayTiming
        AddHandler Me.myCommSvr.DisplayTeletimer, AddressOf myCommSvr_DisplayTeletimer
        AddHandler Me.myCommSvr.DisplayJudgesMessage, AddressOf DisplayJudgesMessage

        myCommSvr.m_ApplicationBussy = False

        Me.myTrackCondition = ""
        rbTote.Checked = True
        '
        Me.tmrWPSPools.Enabled = True


    End Sub

    'Private Sub rbSloppy_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbSloppy.CheckedChanged
    '    'If Me.rbSloppy.Checked Then Me.myTrackCondition = TrackCondition.Sloppy
    '    If sender.Checked Then Me.myTrackCondition = TrackCondition.Sloppy
    'End Sub

    'Private Sub rbSlow_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbSlow.CheckedChanged
    '    'If Me.rbSloppy.Checked Then Me.myTrackCondition = TrackCondition.Slow
    '    If sender.Checked Then Me.myTrackCondition = TrackCondition.Slow
    'End Sub

    'Private Sub rbFast_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbFast.CheckedChanged
    '    'If Me.rbFast.Checked Then Me.myTrackCondition = TrackCondition.Fast
    '    If sender.Checked Then Me.myTrackCondition = TrackCondition.Fast
    'End Sub

    Private Sub rbTote_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbTote.CheckedChanged
        If (Me.Visible) Then
            If sender.Checked Then
                If myCommSvr.m_ApplicationBussy Then
                    Application.DoEvents()
                End If
                Me.myTrackCondition = Me.myCommSvr.p_strToteTrackCondition 'Me.myCommSvr.oCommServerNet.CurrentTrackCondition 'Tote
                txtTrackCondition.Text = myTrackCondition
                Me.PrepareTrackCondition(False)
            End If
        End If
    End Sub

    Private Sub rbText_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbText.CheckedChanged
        If (Me.Visible) Then
            If sender.Checked Then
                If myCommSvr.m_ApplicationBussy Then
                    Application.DoEvents()
                End If
                Me.myTrackCondition = Me.txtTrackCondition.Text
                Me.PrepareTrackCondition(False)
            End If
        End If
    End Sub

    Private Sub txtTrackCondition_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTrackCondition.TextChanged
        If (Me.Visible) Then
            If (rbText.Checked) Then
                If myCommSvr.m_ApplicationBussy Then
                    Application.DoEvents()
                End If
                Me.myTrackCondition = Me.txtTrackCondition.Text
                Me.PrepareTrackCondition(False)
            End If
        End If
    End Sub

    'Private Sub PrepareTrackCondition(ByVal blnFlagClear As Boolean)
    '    'start of header
    '    Dim SOH As String = "01"
    '    'address, lets get it fom the settings, tania will later decide how to approach this
    '    Dim BoardAddress As String = "00" 'My.Settings.MiniBoardAddress
    '    'control, will send it always "on" from here
    '    Dim BoardControl As String = "00"
    '    'dimming, it should be done through the settings as well
    '    Dim BoardDimming As String = "00" '"0" & Hex(My.Settings.MiniBoardDimming)
    '    'pay load type, this value will be always 16 from here (hex(10))
    '    Dim PayloadType As String = "17"
    '    Dim NumbOfString As String = "03"
    '    Dim NumbOfBytes As String = "24" '30 Bytes (6 Columns * 6 Characters)
    '    Dim StartColumn As String = "00" '"5A" '15(0-14)(15-20)
    '    'end of transmission
    '    Dim EOT As String = "04"
    '    '
    '    Dim strTemp As String = " "
    '    If (blnFlagClear = False) Then
    '        strTemp = UCase(myTrackCondition)
    '    End If
    '    strTemp = strTemp.ToString.PadRight(6, " ")

    '    Dim displaySegments() As Emac.DisplayMatrixUtils.DisplaySegment
    '    Dim strCol As String
    '    Dim strColTemp As String
    '    Dim col As String = ""
    '    Dim col1 As String = ""
    '    Dim col2 As String = ""
    '    Dim col3 As String = ""
    '    Dim col4 As String = ""
    '    Dim col5 As String = ""
    '    Dim col6 As String = ""

    '    For intCtr As Integer = 1 To Len(strTemp)
    '        Try
    '            col = ""
    '            strColTemp = ""
    '            strCol = Mid(strTemp, intCtr, 1)
    '            displaySegments = Emac.DisplayMatrixUtils.DisplaySegmentDictionary.GetDisplaySegments(strCol)
    '            For colIdx As Integer = 0 To 5
    '                strColTemp = displaySegments(colIdx).HexRowValue
    '                col = col & strColTemp.PadLeft(2, "0")
    '            Next colIdx
    '        Catch ex As Exception
    '            col = "000000000000"
    '        End Try
    '        If intCtr = 1 Then
    '            col1 = col
    '        ElseIf intCtr = 2 Then
    '            col2 = col
    '        ElseIf intCtr = 3 Then
    '            col3 = col
    '        ElseIf intCtr = 4 Then
    '            col4 = col
    '        ElseIf intCtr = 5 Then
    '            col5 = col
    '        ElseIf intCtr = 6 Then
    '            col6 = col
    '        End If
    '    Next intCtr
    '    '
    '    'build the string to send
    '    Dim MessageToSend As String = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}", _
    '                                             SOH, BoardAddress, BoardControl, BoardDimming, PayloadType, _
    '                                             NumbOfString, StartColumn, NumbOfBytes, _
    '                                             col1, col2, col3, col4, col5, col6, _
    '                                             EOT, Me.CalculateCRC(1), Me.CalculateCRC(2))

    '    'Me.myMiniComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend("010080001700001E008080FE8080008080FE8080008080FE8080008080FE8080008080FE8080040000"))

    '    Me.myMiniComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend(MessageToSend))
    '    If Not Me.myMiniComPort.ErrorMessage = "" Then Me.ShowError(True)

    '    ''
    'End Sub

    Private Sub PrepareTrackCondition(ByVal blnFlagClear As Boolean)
        Dim myDataToSend As String = ""
        Dim MessageToSend As String = ""
        'start of header
        Dim SOH As String = "01"
        'address, lets get it fom the settings, tania will later decide how to approach this
        Dim BoardAddress As String = "02"
        'control, will send it always "on" from here
        Dim BoardControl As String = "00"
        'dimming, it should be done through the settings as well
        Dim BoardDimming As String = "00"
        'pay load type, this value will be always 24 from here (hex(18))
        Dim PayloadType As String = "18"
        'end of transmission
        Dim EOT As String = "04"
        '
        Dim StringNumber As String = "03" 'Ver que string number es TrackconDition
        '
        Dim strTemp As String = " "
        If (blnFlagClear = False) Then
            strTemp = UCase(myTrackCondition)
        End If
        strTemp = strTemp.ToString.PadRight(6, " ")

        Dim strDatatosend = PrepareDataThreeColorSign(strTemp, Color.Yellow, 24)
        strDatatosend = strDatatosend.Replace(" ", "")
        'Me.myComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend(strDatatosend))

        Try
            If (strDatatosend <> "-1") Then
                '(between 1 and 240 bytes of Data
                Dim intNumbDataBytes As Integer
                intNumbDataBytes = (strDatatosend.Length / 2)
                Dim NumbDataBytes As String
                NumbDataBytes = ("0" & Hex(intNumbDataBytes))
                NumbDataBytes = NumbDataBytes.Substring(NumbDataBytes.Length - 2)

                myDataToSend = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}0000",
                                                     SOH, BoardAddress, BoardControl, BoardDimming, PayloadType,
                                                     StringNumber, NumbDataBytes, strDatatosend,
                                                     EOT)

                'convert it to char
                MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
                '
                Me.myComPort.Output(MessageToSend)

                If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()
            End If

        Catch ex As Exception

        End Try
        ''
    End Sub

    Private Sub PrepareExoticsToSend(ByVal intRace As Integer, ByVal blnFlag As Boolean, ByVal blnClearFlag As Boolean)
        Dim myDataToSend As String = ""
        Dim MessageToSend As String = ""
        'start of header
        Dim SOH As String = "01"
        'address, lets get it fom the settings, tania will later decide how to approach this
        Dim BoardAddress As String = "02"
        'control, will send it always "on" from here
        Dim BoardControl As String = "00"
        'dimming, it should be done through the settings as well
        Dim BoardDimming As String = "00"
        'pay load type, this value will be always 24 from here (hex(18))
        Dim PayloadType As String = "18"
        'end of transmission
        Dim EOT As String = "04"
        '
        Dim StringNumber As String = "02" 'string number for Exotics
        '
        Dim strTemp As String = ""
        Dim strTempRunners As String = ""
        '
        Dim intTagPost As Integer = 1
        '
        If (blnFlag) Then
            strTemp = ""
            Try
                'pos#type
                strTemp = "" + WriteExotics(intTagPost, "type") 'pos#

                strTempRunners = WriteExotics(intTagPost, "#")

                Dim tmp() As String
                tmp = Split(strTempRunners, "/")
                Dim count As Integer = tmp.Length - 1
                If (count < 3) Then
                    strTemp = strTemp + "  " + strTempRunners
                    strTemp = strTemp.Replace("[]", "All")
                End If
                'pos#amt
                strTemp = strTemp + "  " + WriteExotics(intTagPost, "amt")
            Catch ex As Exception
                '
            End Try
        Else
            If (blnClearFlag) Then
                strTemp = " "
            Else
                strTemp = UCase(txtRaceStatus.Text)
            End If
        End If

        strTemp = strTemp.ToString.PadRight(25, " ")

        Dim strDatatosend = PrepareDataThreeColorSignExotics(strTemp, Color.Yellow, 24)
        strDatatosend = strDatatosend.Replace(" ", "")
        'strDatatosend = "3132333435"
        'Me.myComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend(strDatatosend))

        Try
            If (strDatatosend <> "-1") Then
                '(between 1 and 240 bytes of Data
                Dim intNumbDataBytes As Integer
                intNumbDataBytes = (strDatatosend.Length / 2)
                Dim NumbDataBytes As String
                NumbDataBytes = ("0" & Hex(intNumbDataBytes))
                NumbDataBytes = NumbDataBytes.Substring(NumbDataBytes.Length - 2)

                'NumbDataBytes = "05"

                myDataToSend = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}0000",
                                                     SOH, BoardAddress, BoardControl, BoardDimming, PayloadType,
                                                     StringNumber, NumbDataBytes, strDatatosend,
                                                     EOT)

                'convert it to char
                MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
                '
                Me.myComPort.Output(MessageToSend)

                If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()
            End If

        Catch ex As Exception
            '
        End Try
        ''
        txtRaceStatus.Text = strTemp

    End Sub

    Private Function WriteExotics(intTag As Integer, strTag As String) As String
        Dim n As Integer = 0

        If (results_tenpage) Then
            n = intTag + (p_intExoCtr * 9)
        ElseIf (results_ninepage) Then
            n = intTag + (p_intExoCtr * 8)
        ElseIf (results_eightpage) Then
            n = intTag + (p_intExoCtr * 7)
        ElseIf (results_sevenpage) Then
            n = intTag + (p_intExoCtr * 6)
        ElseIf (results_sixpage) Then
            n = intTag + (p_intExoCtr * 5)
        ElseIf (results_fivepage) Then
            n = intTag + (p_intExoCtr * 4)
        ElseIf (results_forthpage) Then
            n = intTag + (p_intExoCtr * 3)
        ElseIf (results_thirdpage) Then
            n = intTag + (p_intExoCtr * 2)
        ElseIf (results_secondpage) Then
            n = intTag + p_intExoCtr
        Else
            n = intTag
        End If
        '
        Select Case (strTag)
            Case "#"
                Return results_num(n)
            Case "amt"
                Return results_amt(n)
            Case "type"
                Return results_type(n)
        End Select
        Return ""
    End Function

    'Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    '    Application.DoEvents()
    '    Me.lblTime.Text = Format(Now, "hh:mm:ss tt")
    'End Sub

    Private Sub TimerExTri_Tick(ByVal sender As Object, ByVal e As System.EventArgs)
        'if we're here is bc we have exacta and trifecta
    End Sub

    Private Sub timerOfficial_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles timerOfficial.Tick
        Try
            DisplayAllResults()
        Catch ex As Exception
            ex.ToString()
        End Try
    End Sub

    Private Sub DisplayAllResults()

        Try
            myCommSvr.p_blnTmrResultsBusy = True

            If (myCommSvr.p_blnResultsOut) Then

                Dim intRace As Integer = myCommSvr.p_intRaceResultsOut

                myCommSvr.UpdateResultsWPS(intRace)
                DisplayWPS()

                LoadExotics(intRace, "${0:#,0.00}", True)
                DisplayExotics(intRace)

                If (results_tenpage) Then
                    results_secondpage = False
                    results_thirdpage = False
                    results_forthpage = False
                    results_fivepage = False
                    results_sixpage = False
                    results_sevenpage = False
                    results_eightpage = False
                    results_ninepage = False
                    results_tenpage = False
                ElseIf (results_ninepage) Then
                    results_secondpage = False
                    results_thirdpage = False
                    results_forthpage = False
                    results_fivepage = False
                    results_sixpage = False
                    results_sevenpage = False
                    results_eightpage = False
                    results_ninepage = False
                    results_tenpage = True
                ElseIf (results_eightpage) Then
                    results_secondpage = False
                    results_thirdpage = False
                    results_forthpage = False
                    results_fivepage = False
                    results_sixpage = False
                    results_sevenpage = False
                    results_eightpage = False
                    results_ninepage = True
                    results_tenpage = False
                ElseIf (results_sevenpage) Then
                    results_secondpage = False
                    results_thirdpage = False
                    results_forthpage = False
                    results_fivepage = False
                    results_sixpage = False
                    results_sevenpage = False
                    results_eightpage = True
                    results_ninepage = False
                    results_tenpage = False
                ElseIf (results_sixpage) Then
                    results_secondpage = False
                    results_thirdpage = False
                    results_forthpage = False
                    results_fivepage = False
                    results_sixpage = False
                    results_sevenpage = True
                    results_eightpage = False
                    results_ninepage = False
                    results_tenpage = False
                ElseIf (results_fivepage) Then
                    results_secondpage = False
                    results_thirdpage = False
                    results_forthpage = False
                    results_fivepage = False
                    results_sixpage = True
                    results_sevenpage = False
                    results_eightpage = False
                    results_ninepage = False
                    results_tenpage = False
                ElseIf (results_forthpage) Then
                    results_secondpage = False
                    results_thirdpage = False
                    results_forthpage = False
                    results_fivepage = True
                    results_sixpage = False
                    results_sevenpage = False
                    results_eightpage = False
                    results_ninepage = False
                    results_tenpage = False
                ElseIf (results_thirdpage) Then
                    results_secondpage = False
                    results_thirdpage = False
                    results_forthpage = True
                    results_fivepage = False
                    results_sixpage = False
                    results_sevenpage = False
                    results_eightpage = False
                    results_ninepage = False
                    results_tenpage = False
                ElseIf (results_secondpage) Then
                    results_secondpage = False
                    results_thirdpage = True
                    results_forthpage = False
                    results_fivepage = False
                    results_sixpage = False
                    results_sevenpage = False
                    results_eightpage = False
                    results_ninepage = False
                    results_tenpage = False
                Else
                    results_secondpage = True
                    results_thirdpage = False
                    results_forthpage = False
                    results_fivepage = False
                    results_sixpage = False
                    results_sevenpage = False
                    results_eightpage = False
                    results_ninepage = False
                    results_tenpage = False
                End If

                'If (GlobalClass.dicResultExotic.ContainsKey(GlobalClass.p_intRaceResultsOut)) Then
                '    DisplayFullExoticsResults(GlobalClass.dicResultExotic[GlobalClass.p_intRaceResultsOut], GlobalClass.p_intRaceResultsOut);DisplayFullExoticsResults(GlobalClass.dicResultExotic[GlobalClass.p_intRaceResultsOut], GlobalClass.p_intRaceResultsOut);
                'End If

                myCommSvr.p_intClearResults = myCommSvr.p_intClearResults + 1
                ''If (myCommSvr.p_intClearResults > 24) Then
                ''    ClearResults(True, True)
                ''    myCommSvr.p_intClearResults = 1
                ''End If
                'If (myCommSvr.p_intCurrentRace > intRace) Then
                '    If (Val(myCommSvr.p_strCurrentMTP) <= p_intMTPToClear) Then
                '        ClearResults(True, True)
                '        myCommSvr.p_intClearResults = 1
                '    End If
                'Else
                '    If (myCommSvr.p_intClearResults > 40) Then
                '        ClearResults(True, True)
                '        myCommSvr.p_intClearResults = 1
                '    End If
                'End If
            Else
                timerOfficial.Enabled = False
            End If

            myCommSvr.p_blnTmrResultsBusy = False
        Catch ex As Exception
            ex.ToString()
            myCommSvr.p_blnTmrResultsBusy = False
        End Try
    End Sub

    Private Sub LoadExotics(intRace As Integer, strFormat As String, blnRequire As Boolean)
        Dim intCtrFor As Integer
        Dim intExoticsCount As Integer

        For intCtrFor = 1 To p_intMaxNumbOfExotics
            results_num(intCtrFor) = ""
            results_amt(intCtrFor) = ""
            results_type(intCtrFor) = ""
        Next

        intExoticsCount = 0

        Dim shrRace As Short = CShort(intRace)

        LoadExoticsByType(intRace, strFormat, "EX ", "EXA", "2", False, intExoticsCount)
        LoadExoticsByType(intRace, strFormat, "TRI", "TRI", "50", False, intExoticsCount)
        LoadExoticsByType(intRace, strFormat, "SPR", "SUPER", "10", False, intExoticsCount)
        LoadExoticsByType(intRace, strFormat, "DD ", "D/D", "1", False, intExoticsCount)
        LoadExoticsByType(intRace, strFormat, "P03", "Pick3", "50", False, intExoticsCount)
        LoadExoticsByType(intRace, strFormat, "P04", "Pick4", "50", blnRequire, intExoticsCount)

        results_num_ctr = intExoticsCount

        If (results_num_ctr <= p_intExoCtr) Then
            results_secondpage = False
            results_thirdpage = False
            results_forthpage = False
            results_fivepage = False
            results_sixpage = False
            results_sevenpage = False
            results_eightpage = False
            results_ninepage = False
            results_tenpage = False
        ElseIf (results_num_ctr <= (p_intExoCtr * 2)) Then
            results_thirdpage = False
            results_forthpage = False
            results_fivepage = False
            results_sixpage = False
            results_sevenpage = False
            results_eightpage = False
            results_ninepage = False
            results_tenpage = False
        ElseIf (results_num_ctr <= (p_intExoCtr * 3)) Then
            results_forthpage = False
            results_fivepage = False
            results_sixpage = False
            results_sevenpage = False
            results_eightpage = False
            results_ninepage = False
            results_tenpage = False
        ElseIf (results_num_ctr <= (p_intExoCtr * 4)) Then
            results_fivepage = False
            results_sixpage = False
            results_sevenpage = False
            results_eightpage = False
            results_ninepage = False
            results_tenpage = False
        ElseIf (results_num_ctr <= (p_intExoCtr * 5)) Then
            results_sixpage = False
            results_sevenpage = False
            results_eightpage = False
            results_ninepage = False
            results_tenpage = False
        ElseIf (results_num_ctr <= (p_intExoCtr * 6)) Then
            results_sevenpage = False
            results_eightpage = False
            results_ninepage = False
            results_tenpage = False
        ElseIf (results_num_ctr <= (p_intExoCtr * 7)) Then
            results_eightpage = False
            results_ninepage = False
            results_tenpage = False
        ElseIf (results_num_ctr <= (p_intExoCtr * 8)) Then
            results_ninepage = False
            results_tenpage = False
        ElseIf (results_num_ctr <= (p_intExoCtr * 9)) Then
            results_tenpage = False
        End If

    End Sub

    Private Sub LoadExoticsByType(intRace As Integer, strFormat As String, strType As String, strBetName As String, strBetAmt As String, blnRequired As Boolean, ByRef intExoticsCount As Integer)
        Dim objRsExo As RSIData.clsResultExotic
        Dim strTempWinningList As String = ""
        Dim shrExoCtr As Short
        Dim intNumOfPrices As Integer
        Dim strBetText As String = ""
        Dim strTemp As String = ""

        objRsExo = New RSIData.clsResultExotic()

        Dim shrRace As Short = CShort(intRace)

        Try
            Dim Key As String
            Dim strKey As String
            Key = intRace.ToString() + "_" + strType
            strKey = intRace.ToString() + "_" + strType
            objRsExo = myCommSvr.oCommServerNet.ObjectRsExoAutototeV6(shrRace, strKey).Item(Key)
            intNumOfPrices = objRsExo.NumberOfPrices
            For shrExoCtr = 1 To intNumOfPrices
                If (Char.IsUpper(objRsExo.Status(shrExoCtr), 0)) Then
                    strTemp = "" + objRsExo.BetAmount(shrExoCtr)
                    If String.IsNullOrEmpty(strTemp) Then
                        strBetText = strBetName
                    Else
                        strBetText = "$" + strTemp.Trim() + " " + strBetName
                    End If
                    intExoticsCount = intExoticsCount + 1
                    If (FormatWinList(intRace, objRsExo.WinnerList(shrExoCtr), False).Trim() = "0") Then
                        results_num(intExoticsCount) = objRsExo.WinnerList(shrExoCtr)
                    Else
                        results_num(intExoticsCount) = FormatWinList(intRace, objRsExo.WinnerList(shrExoCtr), False)
                    End If
                    If (Not blnRequired) Then
                        results_amt(intExoticsCount) = TagVal(objRsExo.Amount(shrExoCtr), strFormat + " includecents")
                    Else
                        results_amt(intExoticsCount) = objRsExo.Required(shrExoCtr).Trim() + " " + TagVal(objRsExo.Amount(shrExoCtr), strFormat + " includecents")
                    End If
                    results_type(intExoticsCount) = strBetText
                End If
            Next
            '
            objRsExo = Nothing
            '
        Catch ex As Exception
            ex.ToString()
        End Try

    End Sub

    Private Function TagVal(vntIncome As Object, strNumFormat As String) As String
        Dim n As Integer
        Dim strRes As String = ""
        Dim intpntr As Integer
        Dim strOrig As String
        Dim dblTemp As Double

        Try
            If IsDBNull(vntIncome) Then
                Return " "
            Else
                strOrig = Convert.ToString(vntIncome)
            End If

            strOrig = strOrig.Trim()

            If (IsDBNull(strNumFormat)) Then
                strNumFormat = ""
            End If

            If ((strNumFormat <> "") And (strOrig.Length > 0)) Then
                intpntr = strNumFormat.IndexOf(" ")
                If (intpntr <> -1) Then
                    'includecents
                    strNumFormat = Strings.Left(strNumFormat, intpntr)

                    Try
                        dblTemp = (Double.Parse(strOrig) / 100)
                        strOrig = String.Format(strNumFormat, dblTemp)
                    Catch ex As Exception
                        ex.ToString()
                    End Try
                Else
                    'numeric formatting only
                    strOrig = String.Format(strNumFormat, strOrig) 'VB6.Format(strOrig, strNumFormat)
                End If
            End If

            strOrig = strOrig.Trim()
            Dim bytTemp() As Byte = System.Text.Encoding.ASCII.GetBytes(strOrig)
            For n = 1 To bytTemp.Length
                If (bytTemp(n - 1) < 32) Then
                    'not allowed
                ElseIf ((bytTemp(n - 1) > 127)) Then
                    'not allowed
                Else
                    strRes = strRes + Strings.Mid(strOrig, n, 1)
                End If
            Next

            If (strRes.IndexOf("......") <> -1) Then
                strRes = " "
            End If

            If (strRes.IndexOf("--") <> -1) Then
                strRes = "-"
            End If

            If (Strings.Left(strRes, 5) = "00000") Then
                If (Microsoft.VisualBasic.Conversion.Val(strRes) = 0) Then
                    strRes = "0"
                End If
            End If
            If (strRes.IndexOf("9999") <> -1) Then
                strRes = "999"
            End If

            If strRes.Length = 0 Then
                Return " "
            Else
                Return strRes
            End If

        Catch ex As Exception
            Return " "
        End Try

    End Function

    Public Function FormatWinList(intRace As Integer, strWinList As String, blnNoSpaces As Boolean) As String
        'The variable strWinList will stored the winners for an specific exotic.
        'The actual winners are separeted by '/' or '-'; while the scratches, consolations
        'Or dh in pick-n are separeted by ','.
        Dim strTemp As String = ""
        Dim intRaceCounter As Integer
        Dim intpntr As Integer
        Dim intLenWinList As Integer
        Dim strFormatWinList As String = ""
        Dim strDelimeter As String = "-"

        intRaceCounter = 1
        intLenWinList = strWinList.Length

        For intpntr = 1 To intLenWinList
            strDelimeter = Strings.Mid(strWinList, intpntr, 1)

            Select Case (Strings.Mid(strWinList, intpntr, 1))
                Case ","
                    If (blnNoSpaces) Then
                        If ((strTemp = "[]") Or (strTemp.ToUpper() = "ALL")) Then
                            strFormatWinList = strFormatWinList + strTemp + ","
                        Else
                            strFormatWinList = strFormatWinList + Conversion.Val(strTemp).ToString() + ","
                        End If
                    Else
                        If ((strTemp = "[]") Or (strTemp.ToUpper() = "ALL")) Then
                            strFormatWinList = strFormatWinList + strTemp + " , "
                        Else
                            strFormatWinList = strFormatWinList + Conversion.Val(strTemp).ToString() + " , "
                        End If
                    End If
                    strTemp = ""
                Case "/", "-"
                    If (blnNoSpaces) Then
                        If ((strTemp = "[]") Or (strTemp.ToUpper() = "ALL")) Then
                            strFormatWinList = strFormatWinList + strTemp + strDelimeter
                        Else
                            strFormatWinList = strFormatWinList + Microsoft.VisualBasic.Conversion.Val(strTemp).ToString() + strDelimeter
                        End If
                    Else
                        If ((strTemp = "[]") Or (strTemp.ToUpper() = "ALL")) Then
                            strFormatWinList = strFormatWinList + strTemp + " " + strDelimeter + " "
                        Else
                            strFormatWinList = strFormatWinList + Conversion.Val(strTemp).ToString() + " " + strDelimeter + " "
                        End If

                    End If
                    strTemp = ""
                    intRaceCounter = intRaceCounter + 1
                Case Else
                    strTemp = strTemp + Strings.Mid(strWinList, intpntr, 1)
            End Select
        Next

        If ((strTemp = "[]") Or (strTemp.ToUpper() = "ALL")) Then
            '
        Else
            strTemp = "" + Conversion.Val(strTemp).ToString()
        End If

        Return strFormatWinList + strTemp

    End Function

    Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Application.DoEvents()
        Me.myCommSvr.m_ApplicationBussy = True
        p_blnUpdateManually = False

        'Clear Board
        myCommSvr.p_blnResultsOut = False
        myCommSvr.p_intRaceResultsOut = 0
        Me.timerOfficial.Enabled = False
        Me.tmrWPSPools.Enabled = False
        Me.tmrTest.Enabled = False
        '
        Me.ClearItems(26) 'ClearOdds out
        'Me.ClearItems(28) 'ClearOdds in
        Me.ClearItems(27) 'ClearOdds out
        'Me.ClearItems(29) 'ClearOdds in
        'Me.ClearItems(10) 'ClearOdds out 15-16
        'Me.ClearItems(11) 'ClearOdds in 15-16

        Me.ClearItems(50) 'Time Out

        Me.ClearItems(14) 'running order out
        'Me.ClearItems(30) 'running order in

        Me.ClearItems(15) 'win out
        'Me.ClearItems(31) 'win in

        Me.ClearItems(16) 'place out
        'Me.ClearItems(32) 'place in

        Me.ClearItems(17) 'show out
        'Me.ClearItems(18) 'show out
        'Me.ClearItems(32) 'show out
        Me.ClearItems(34) 'show out
        'Me.ClearItems(33) 'show in

        Me.ClearItems(48)
        Me.ClearItems(36)
        Me.ClearItems(38)
        Me.ClearItems(40)
        Me.ClearItems(42)

        '
        txtPoolType.Text = ""
        Me.PrepareWPSPoolHeaderToSend(True)

        txtRaceStatus.Text = ""
        Me.PrepareExoticsToSend(1, False, True)

        txtOfficial.Text = " "
        Me.PrepareStatusDataToSend()

        'txtTrackCondition.Text = " "
        'Me.myTrackCondition = Me.txtTrackCondition.Text
        Me.myTrackCondition = " "
        Me.PrepareTrackCondition(True)

        'lblFIN.Text = " "
        'lbl78th.Text = " "
        'lbl58th.Text = " "
        'lblHalf.Text = " "
        'lblQtr.Text = " "
        'Me.PrepareFinalsToSend(True)

        'Me.ClearItems(18) 'DD out
        'Me.ClearItems(34) 'DD in

        'Me.rbExacta.Checked = False
        'Me.rbTrifecta.Checked = False
        'Me.ClearItems(19) 'exacta out

        'Me.ClearItems(35) 'exacta in

        'Me.ClearItems(20) 'pick 3 out
        'Me.ClearItems(36) 'pick 3 in

        'Me.ClearItems(21) 'trifecta out
        'Me.ClearItems(37) 'trifecta in
        ''

        myCommSvr.CancelProcess()
        Application.DoEvents()
        myCommSvr = Nothing
        Me.myComPort = Nothing
        Me.KillComSvr()
        Me.KillDataSvr()
    End Sub

#End Region

#Region " Toolbar Events "

    'Private Sub ChangeSerialPortToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChangeSerialPortToolStripMenuItem.Click
    '    Me.gbPorts.Visible = True
    '    Me.lblPort.Visible = False
    'End Sub

    Private Sub cbOfficialTime_DropDownClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbOfficialTime.DropDownClosed
        If Me.cbOfficialTime.Text.Trim = "" Then Return
        '
        Me.myOfficialTimeSpan = Me.cbOfficialTime.Text.Trim
        '
        Me.TimersToolStripMenuItem.HideDropDown()
    End Sub

    'Private Sub tscbIntensity_DropDownClosed(ByVal sender As Object, ByVal e As System.EventArgs)
    '    If Me.tscbIntensity.Text.Trim = "" Then Return
    '    '
    '    myInsideIntensity = Me.tscbIntensity.Text.Trim
    '    '
    '    Me.IntensityToolStripMenuItem.HideDropDown()
    'End Sub

    Private Sub CheckIfReceivingDataToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckIfReceivingDataToolStripMenuItem.Click
        Dim f As New frmCheckMessage(Me.myCommSvr)
        f.ShowDialog()
        f.Dispose()
        ''
    End Sub

    Private Sub CloseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub ClearOddsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearOddsToolStripMenuItem.Click
        If (Not myCommSvr.m_ApplicationBussy) Then
            Try
                myCommSvr.m_ApplicationBussy = True

                For Each ctl As Control In Me.GroupBox1.Controls
                    If TypeOf ctl Is Windows.Forms.TextBox Then
                        ctl.Text = ""
                    End If
                Next

                For Each ctl As Control In Me.gbMTP.Controls
                    If TypeOf ctl Is Windows.Forms.TextBox Then
                        ctl.Text = ""
                    End If
                Next

                For Each ctl As Control In Me.gbTime.Controls
                    If TypeOf ctl Is Windows.Forms.TextBox Then
                        ctl.Text = ""
                    End If
                Next

                For Each ctl As Control In Me.gbPostTime.Controls
                    If TypeOf ctl Is Windows.Forms.TextBox Then
                        ctl.Text = ""
                    End If
                Next

                For Each ctl As Control In Me.gbOdds.Controls
                    If TypeOf ctl Is Windows.Forms.TextBox Then
                        ctl.Text = ""
                    End If
                    If TypeOf ctl Is Windows.Forms.Label Then
                        If ctl.Text.Trim() = "/" Then
                            ctl.Visible = False
                        End If
                    End If
                Next

                Me.bsRace.EndEdit()
                Me.bsMTP.EndEdit()
                Me.bsTOD.EndEdit()
                Me.bsPostTime.EndEdit()
                Me.bsODDS.EndEdit()

                Me.ClearItems(26) 'ClearOdds out
                'Me.ClearItems(28) 'ClearOdds in
                Me.ClearItems(27) 'ClearOdds out
                'Me.ClearItems(29) 'ClearOdds in
                'Me.ClearItems(10) 'ClearOdds out 15-16
                'Me.ClearItems(11) 'ClearOdds in  15-16
                'If Me.HasMiniBoard Then Me.ClearMiniBoard(PayloadType.Odds)

                myCommSvr.m_ApplicationBussy = False
            Catch
                myCommSvr.m_ApplicationBussy = False
            End Try
        End If
    End Sub

    Private Sub ClearRunningOrderToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearRunningOrderToolStripMenuItem.Click
        If (Not myCommSvr.m_ApplicationBussy) Then
            Try
                myCommSvr.m_ApplicationBussy = True
                p_blnUpdateManually = False

                For Each ctl As Control In Me.gbOrder.Controls
                    If TypeOf ctl Is Windows.Forms.TextBox Then
                        ctl.Text = ""
                    End If
                    If TypeOf ctl Is Windows.Forms.CheckBox Then
                        CType(ctl, Windows.Forms.CheckBox).Checked = False
                    End If
                Next

                For Each ctl As Control In Me.gbOptions.Controls
                    If TypeOf ctl Is Windows.Forms.CheckBox Then
                        CType(ctl, Windows.Forms.CheckBox).Checked = False
                    End If
                Next

                Me.bsRunningOrder.EndEdit()
                Me.bsStatus.EndEdit()

                Me.ClearItems(14) 'running order out
                Me.ClearItems(30) 'running order in
                'Me.ClearMiniBoard(PayloadType.RunningOrder)
                '
                p_blnUpdateManually = True
                myCommSvr.m_ApplicationBussy = False
            Catch
                p_blnUpdateManually = True
                myCommSvr.m_ApplicationBussy = False
            End Try
        End If
    End Sub

    Private Sub ClearWinToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearWinToolStripMenuItem.Click
        If (Not myCommSvr.m_ApplicationBussy) Then
            Try
                myCommSvr.m_ApplicationBussy = True

                For Each ctl As Control In Me.gbWin.Controls
                    If TypeOf ctl Is Windows.Forms.TextBox Then
                        ctl.Text = ""
                    End If
                Next

                Me.bsWin.EndEdit()

                Me.ClearItems(15) 'win out
                'Me.ClearItems(31) 'win in

                myCommSvr.m_ApplicationBussy = False
            Catch
                myCommSvr.m_ApplicationBussy = False
            End Try
        End If
    End Sub

    Private Sub ClearPlaceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearPlaceToolStripMenuItem.Click
        If (Not myCommSvr.m_ApplicationBussy) Then
            Try
                myCommSvr.m_ApplicationBussy = True

                For Each ctl As Control In Me.gbPlace.Controls
                    If TypeOf ctl Is Windows.Forms.TextBox Then
                        ctl.Text = ""
                    End If
                Next

                Me.bsPlace.EndEdit()

                Me.ClearItems(16) 'place out
                'Me.ClearItems(32) 'place in

                myCommSvr.m_ApplicationBussy = False
            Catch
                myCommSvr.m_ApplicationBussy = False
            End Try
        End If

    End Sub

    Private Sub ClearShowToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearShowToolStripMenuItem.Click
        If (Not myCommSvr.m_ApplicationBussy) Then
            Try
                myCommSvr.m_ApplicationBussy = True

                For Each ctl As Control In Me.gbShow.Controls
                    If TypeOf ctl Is Windows.Forms.TextBox Then
                        ctl.Text = ""
                    End If
                Next

                Me.bsShow.EndEdit()

                Me.ClearItems(17) 'show out
                Me.ClearItems(18) 'show out
                'Me.ClearItems(33) 'show in

                myCommSvr.m_ApplicationBussy = False
            Catch
                myCommSvr.m_ApplicationBussy = False
            End Try
        End If
    End Sub

    Private Sub ClearTimingToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearTimingToolStripMenuItem.Click
        If (Not myCommSvr.m_ApplicationBussy) Then
            Try
                myCommSvr.m_ApplicationBussy = True
                p_blnUpdateManually = False

                For Each ctl As Control In Me.Label80.Controls
                    If TypeOf ctl Is Windows.Forms.TextBox Then
                        ctl.Text = ""
                    End If
                Next

                Me.bsTimingFinish.EndEdit()
                Me.bsTimingMile.EndEdit()
                Me.bsTiming14.EndEdit()
                Me.bsTiming12.EndEdit()
                Me.bsTiming34.EndEdit()

                'Me.ClearMiniBoard(PayloadType.Timing)
                '
                p_blnUpdateManually = True
                myCommSvr.m_ApplicationBussy = False
            Catch
                p_blnUpdateManually = True
                myCommSvr.m_ApplicationBussy = False
            End Try
        End If
    End Sub


    'Private Sub ClearDailyDoubleToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    For Each ctl As Control In Me.gbDD.Controls
    '        If TypeOf ctl Is Windows.Forms.TextBox Then ctl.Text = ""
    '    Next
    '    '
    '    Me.bsDD.EndEdit()
    '    Me.ClearItems(18) 'DD
    '    ''
    'End Sub

    Private Sub ClearExoticsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearExoticsToolStripMenuItem.Click
        'If (Not myCommSvr.m_ApplicationBussy) Then
        '    Try
        '        myCommSvr.m_ApplicationBussy = True

        '        Me.timerOfficial.Enabled = False

        '        'For Each ctl As Control In Me.gbDD.Controls
        '        '   If TypeOf ctl Is Windows.Forms.TextBox Then
        '        '       ctl.Text = ""
        '        '   End If
        '        'Next
        '        For Each ctl As Control In Me.gbPerfecta.Controls
        '            If TypeOf ctl Is Windows.Forms.TextBox Then
        '                ctl.Text = ""
        '            End If
        '        Next
        '        'For Each ctl As Control In Me.gbBet3.Controls
        '        '   If TypeOf ctl Is Windows.Forms.TextBox Then
        '        '       ctl.Text = ""
        '        '   End If
        '        'Next
        '        For Each ctl As Control In Me.gbTrifecta.Controls
        '            If TypeOf ctl Is Windows.Forms.TextBox Then
        '                ctl.Text = ""
        '            End If
        '        Next

        '        Me.rbExacta.Checked = False
        '        Me.rbTrifecta.Checked = False

        '        'Me.bsDD.EndEdit()
        '        Me.bsPerfecta.EndEdit()
        '        'Me.bsBet3.EndEdit()
        '        Me.bsTrifecta.EndEdit()

        '        'Me.ClearItems(18) 'DD out
        '        'Me.ClearItems(34) 'DD in

        '        Me.ClearItems(19) 'exacta out
        '        'Me.ClearItems(35) 'exacta in

        '        'Me.ClearItems(20) 'pick 3 out
        '        'Me.ClearItems(36) 'pick 3 in

        '        'Me.ClearItems(21) 'trifecta out
        '        'Me.ClearItems(37) 'trifecta in

        '        myCommSvr.m_ApplicationBussy = False
        '    Catch
        '        myCommSvr.m_ApplicationBussy = False
        '    End Try
        'End If
    End Sub

    'Private Sub ClearBET3ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    For Each ctl As Control In Me.gbBet3.Controls
    '        If TypeOf ctl Is Windows.Forms.TextBox Then ctl.Text = ""
    '    Next
    '    '
    '    Me.bsBet3.EndEdit()
    '    Me.ClearItems(20) 'pick 3
    'End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        'Dim f As New frmAbout
        'f.ShowDialog()
        ''
    End Sub

    Private Sub ClearAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearAllToolStripMenuItem.Click
        If (Not myCommSvr.m_ApplicationBussy) Then
            Try
                myCommSvr.m_ApplicationBussy = True
                ClearAll()
                myCommSvr.m_ApplicationBussy = False
            Catch
                myCommSvr.m_ApplicationBussy = False
            End Try
        End If
    End Sub

    Private Sub ClearAll()
        Try
            myCommSvr.p_blnResultsOut = False
            myCommSvr.p_intRaceResultsOut = 0
            Me.timerOfficial.Enabled = False
            Me.tmrWPSPools.Enabled = False
            p_blnUpdateManually = False

            For Each ctl As Control In Me.GroupBox1.Controls
                If TypeOf ctl Is Windows.Forms.TextBox Then
                    ctl.Text = ""
                End If
            Next

            For Each ctl As Control In Me.gbWPSPools.Controls
                If TypeOf ctl Is Windows.Forms.TextBox Then
                    ctl.Text = ""
                End If
            Next

            For Each ctl As Control In Me.gbMTP.Controls
                If TypeOf ctl Is Windows.Forms.TextBox Then
                    ctl.Text = ""
                End If
            Next

            For Each ctl As Control In Me.gbTime.Controls
                If TypeOf ctl Is Windows.Forms.TextBox Then
                    ctl.Text = ""
                End If
            Next

            For Each ctl As Control In Me.gbPostTime.Controls
                If TypeOf ctl Is Windows.Forms.TextBox Then
                    ctl.Text = ""
                End If
            Next

            For Each ctl As Control In Me.gbOdds.Controls
                If TypeOf ctl Is Windows.Forms.TextBox Then
                    ctl.Text = ""
                End If
                If TypeOf ctl Is Windows.Forms.Label Then
                    If ctl.Text.Trim() = "/" Then
                        ctl.Visible = False
                    End If
                End If
            Next

            For Each ctl As Control In Me.gbOrder.Controls
                If TypeOf ctl Is Windows.Forms.TextBox Then
                    ctl.Text = ""
                End If
                If TypeOf ctl Is Windows.Forms.CheckBox Then
                    CType(ctl, Windows.Forms.CheckBox).Checked = False
                End If
            Next

            For Each ctl As Control In Me.gbWin.Controls
                If TypeOf ctl Is Windows.Forms.TextBox Then
                    ctl.Text = ""
                End If
            Next

            For Each ctl As Control In Me.gbPlace.Controls
                If TypeOf ctl Is Windows.Forms.TextBox Then
                    ctl.Text = ""
                End If
            Next

            For Each ctl As Control In Me.gbShow.Controls
                If TypeOf ctl Is Windows.Forms.TextBox Then
                    ctl.Text = ""
                End If
            Next

            For Each ctl As Control In Me.gbOptions.Controls
                If TypeOf ctl Is Windows.Forms.CheckBox Then
                    CType(ctl, Windows.Forms.CheckBox).Checked = False
                End If
            Next

            For Each ctl As Control In Me.gbTiming.Controls
                If TypeOf ctl Is Windows.Forms.TextBox Then
                    ctl.Text = ""
                End If
            Next

            'For Each ctl As Control In Me.gbDD.Controls
            '   If TypeOf ctl Is Windows.Forms.TextBox Then
            '       ctl.Text = ""
            '   End If
            'Next
            'For Each ctl As Control In Me.gbPerfecta.Controls
            '    If TypeOf ctl Is Windows.Forms.TextBox Then
            '        ctl.Text = ""
            '    End If
            'Next
            'For Each ctl As Control In Me.gbBet3.Controls
            '   If TypeOf ctl Is Windows.Forms.TextBox Then
            '       ctl.Text = ""
            '   End If
            'Next
            'For Each ctl As Control In Me.gbTrifecta.Controls
            '    If TypeOf ctl Is Windows.Forms.TextBox Then
            '        ctl.Text = ""
            '    End If
            'Next

            'Me.rbExacta.Checked = False
            'Me.rbTrifecta.Checked = False

            'For Each ctl As Object In Me.components.Components
            '    If TypeOf ctl Is BindingSource Then
            '        CType(ctl, BindingSource).EndEdit()
            '    End If
            'Next

            Me.bsRace.EndEdit()
            Me.bsMTP.EndEdit()
            Me.bsTOD.EndEdit()
            Me.bsPostTime.EndEdit()
            Me.bsODDS.EndEdit()
            Me.bsRunningOrder.EndEdit()
            Me.bsStatus.EndEdit()
            Me.bsWin.EndEdit()
            Me.bsPlace.EndEdit()
            Me.bsShow.EndEdit()
            Me.bsWPSPools.EndEdit()
            Me.bsTimingFinish.EndEdit()
            Me.bsTimingMile.EndEdit()
            Me.bsTiming34.EndEdit()
            Me.bsTiming12.EndEdit()
            Me.bsTiming14.EndEdit()
            'Me.bsDD.EndEdit()
            'Me.bsPerfecta.EndEdit()
            'Me.bsBet3.EndEdit()
            'Me.bsTrifecta.EndEdit()

            Me.ClearItems(12)
            Me.ClearItems(13)

            Me.ClearItems(26) 'ClearOdds out
            'Me.ClearItems(28) 'ClearOdds in
            Me.ClearItems(27) 'ClearOdds out
            'Me.ClearItems(29) 'ClearOdds in
            'Me.ClearItems(10) 'ClearOdds out 15-16
            'Me.ClearItems(11) 'ClearOdds in 15-16

            Me.ClearItems(50) 'Time Out

            Me.ClearItems(14) 'running order out
            'Me.ClearItems(30) 'running order in

            Me.ClearItems(15) 'win out
            'Me.ClearItems(31) 'win in

            Me.ClearItems(16) 'place out
            'Me.ClearItems(32) 'place in

            Me.ClearItems(17) 'show out
            'Me.ClearItems(18) 'show out
            Me.ClearItems(34) 'show out
            'Me.ClearItems(33) 'show in

            Me.ClearItems(48)
            Me.ClearItems(36)
            Me.ClearItems(38)
            Me.ClearItems(40)
            Me.ClearItems(42)

            txtPoolType.Text = ""
            Me.PrepareWPSPoolHeaderToSend(True)

            txtRaceStatus.Text = ""
            Me.PrepareExoticsToSend(1, False, True)

            txtOfficial.Text = ""
            Me.PrepareStatusDataToSend()

            If (rbText.Checked) Then
                txtTrackCondition.Text = " "
            Else
                txtTrackCondition.Text = " "
                rbText.Checked = True
            End If

            'rbText.Checked = True
            'txtTrackCondition.Text = " "
            'Me.myTrackCondition = Me.txtTrackCondition.Text
            'Me.PrepareTrackCondition(True)

            'lblFIN.Text = " "
            'lbl78th.Text = " "
            'lbl58th.Text = " "
            'lblHalf.Text = " "
            'lblQtr.Text = " "
            'Me.PrepareFinalsToSend(True)

            'Me.ClearItems(18) 'DD out
            'Me.ClearItems(34) 'DD in

            'Me.ClearItems(19) 'exacta out
            'Me.ClearItems(35) 'exacta in

            'Me.ClearItems(20) 'pick 3 out
            'Me.ClearItems(36) 'pick 3 in

            'Me.ClearItems(21) 'trifecta out
            'Me.ClearItems(37) 'trifecta in

            'Me.ClearMiniBoard(PayloadType.allPayLoads)

            p_blnUpdateManually = True
        Catch ex As Exception
            p_blnUpdateManually = True
        End Try
    End Sub

    Private Sub ClearResultsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearResultsToolStripMenuItem.Click
        If (Not myCommSvr.m_ApplicationBussy) Then
            Try
                myCommSvr.m_ApplicationBussy = True
                ClearResults(True, True)
                myCommSvr.m_ApplicationBussy = False
            Catch
                myCommSvr.m_ApplicationBussy = False
            End Try
        End If
    End Sub

    Private Sub ClearResults(ByVal blnClearRO As Boolean, ByVal blnClearTT As Boolean)
        Try
            myCommSvr.p_blnResultsOut = False
            myCommSvr.p_intRaceResultsOut = 0
            Me.timerOfficial.Enabled = False
            p_blnUpdateManually = False

            If blnClearRO Then
                For Each ctl As Control In Me.gbOrder.Controls
                    If TypeOf ctl Is Windows.Forms.TextBox Then
                        ctl.Text = ""
                    End If
                    If TypeOf ctl Is Windows.Forms.CheckBox Then
                        CType(ctl, Windows.Forms.CheckBox).Checked = False
                    End If
                Next
            End If

            For Each ctl As Control In Me.gbWin.Controls
                If TypeOf ctl Is Windows.Forms.TextBox Then
                    ctl.Text = ""
                End If
            Next

            For Each ctl As Control In Me.gbPlace.Controls
                If TypeOf ctl Is Windows.Forms.TextBox Then
                    ctl.Text = ""
                End If
            Next

            For Each ctl As Control In Me.gbShow.Controls
                If TypeOf ctl Is Windows.Forms.TextBox Then
                    ctl.Text = ""
                End If
            Next

            For Each ctl As Control In Me.gbOptions.Controls
                If TypeOf ctl Is Windows.Forms.CheckBox Then
                    CType(ctl, Windows.Forms.CheckBox).Checked = False
                End If
            Next

            If blnClearTT Then
                For Each ctl As Control In Me.gbTiming.Controls
                    If TypeOf ctl Is Windows.Forms.TextBox Then
                        ctl.Text = ""
                    End If
                Next
            End If

            If blnClearRO Then
                Me.bsRunningOrder.EndEdit()
            End If

            Me.bsStatus.EndEdit()
            Me.bsWin.EndEdit()
            Me.bsPlace.EndEdit()
            Me.bsShow.EndEdit()

            If blnClearTT Then
                Me.bsTimingFinish.EndEdit()
                Me.bsTimingMile.EndEdit()
                Me.bsTiming34.EndEdit()
                Me.bsTiming12.EndEdit()
                Me.bsTiming14.EndEdit()
                Me.ClearItems(12)
                Me.ClearItems(13)
            End If


            If blnClearRO Then
                Me.ClearItems(14) 'running order out
                'Me.ClearItems(30) 'running order in
            End If

            Me.ClearItems(15) 'win out
            'Me.ClearItems(31) 'win in

            Me.ClearItems(16) 'place out
            'Me.ClearItems(32) 'place in

            Me.ClearItems(17) 'show out
            Me.ClearItems(34) 'show out
            'Me.ClearItems(18) 'show out
            'Me.ClearItems(33) 'show in

            txtRaceStatus.Text = ""
            PrepareExoticsToSend(1, False, True)

            results_secondpage = False
            results_thirdpage = False
            results_forthpage = False
            results_fivepage = False
            results_sixpage = False
            results_sevenpage = False
            results_eightpage = False
            results_ninepage = False
            results_tenpage = False

            txtOfficial.Text = ""
            PrepareStatusDataToSend()

            'Me.ClearItems(18) 'DD out
            'Me.ClearItems(34) 'DD in

            'Me.ClearItems(19) 'exacta out
            'Me.ClearItems(35) 'exacta in

            'Me.ClearItems(20) 'pick 3 out
            'Me.ClearItems(36) 'pick 3 in

            'Me.ClearItems(21) 'trifecta out
            'Me.ClearItems(37) 'trifecta in

            'Me.ClearMiniBoard(PayloadType.WPS)

            p_blnUpdateManually = True
        Catch
            p_blnUpdateManually = True
        End Try
        'End If    
    End Sub

    Private Sub IntensityToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IntensityToolStripMenuItem.Click
        Dim f As New frmIntensity
        f.ShowDialog()
        '
        Me.myInsideIntensity = My.Settings.InsideIntensity
        Me.myOutsideIntensity = My.Settings.OutsideIntensity
    End Sub

    Private Sub SettingsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SettingsToolStripMenuItem.Click
        Dim f As New frmSettings
        f.ShowDialog()
        '
        Me.myOutsideIntensity = My.Settings.OutsideIntensity
        Me.myInsideIntensity = My.Settings.InsideIntensity
        '

        ''
    End Sub

#End Region

#Region " Clear Board Methods "

    'Private Sub ClearOdds()
    '    Me.myColMessages.Clear()
    '    'race
    '    Dim RaceText1 As String = IIf(Me.txtRaceA.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtRaceA.Text.Trim = "", "S", Me.txtRaceA.Text.Trim)))
    '    Dim RaceText2 As String = IIf(Me.txtRaceB.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtRaceB.Text.Trim = "", "S", Me.txtRaceB.Text.Trim)))
    '    'mtp
    '    Dim mtpText1 As String = IIf(Me.txtMTPA.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtMTPA.Text.Trim = "", "S", Me.txtMTPA.Text.Trim)))
    '    Dim mtpText2 As String = IIf(Me.txtMTPB.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtMTPB.Text.Trim = "", "S", Me.txtMTPB.Text.Trim)))

    '    Dim strCheckSum1 As String = String.Format("{0}{1}{2}{3}000000000000000000000000",
    '                                              RaceText1, RaceText2, mtpText1, mtpText2)

    '    Dim strCheckSum2 As String = String.Format("00000000000000000000000000000000")

    '    Try
    '        Dim MessageToSend As String = ""
    '        Dim ConstantDataToSend As String = ""
    '        Dim myDataToSend As String = ""
    '        Dim strPort As String

    '        'constant data to send
    '        ConstantDataToSend = String.Format("55AA{0}11", Hex(26)) 'port 26
    '        'complete String with odds 1-6

    '        Dim strMsgCount As String = Me.Getm_Ctr

    '        myDataToSend = String.Format("55AA{0}11{1}{2}{3}{4}{5}{6}{7}000000000000000000000000{8}" _
    '                          , Hex(26), strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myOutsideIntensity _
    '                          , RaceText1, RaceText2, mtpText1, mtpText2, Me.CalcChecksumToAsciiHexString(Me.myOutsideIntensity & strCheckSum1))

    '        'work it out
    '        MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
    '        Me.myColMessages.Add(MessageToSend)

    '        ''clear outside board
    '        'ConstantDataToSend = String.Format("55AA{0}11", Hex(28)) 'port 26
    '        'strMsgCount = Me.Getm_Ctr
    '        ''
    '        'myDataToSend = String.Format("55AA{0}11{1}{2}{3}{4}{5}{6}{7}000000000000000000000000{8}" _
    '        '                  , Hex(28), strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myInsideIntensity _
    '        '                  , RaceText1, RaceText2, mtpText1, mtpText2, Me.CalcChecksumToAsciiHexString(Me.myInsideIntensity & strCheckSum1))
    '        ''work it out
    '        'MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
    '        'Me.myColMessages.Add(MessageToSend)

    '        'complete string with odds 7-12
    '        MessageToSend = ""
    '        myDataToSend = ""
    '        'ConstantDataToSend = ""
    '        '
    '        ConstantDataToSend = String.Format("55AA{0}11", Hex(27))
    '        '
    '        strMsgCount = Me.Getm_Ctr

    '        myDataToSend = String.Format("55AA{0}11{1}{2}{3}00000000000000000000000000000000{4}" _
    '                          , Hex(27), strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myOutsideIntensity _
    '                          , Me.CalcChecksumToAsciiHexString(Me.myOutsideIntensity & strCheckSum2))

    '        'work it out
    '        MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
    '        '
    '        Me.myColMessages.Add(MessageToSend)

    '        'clear the rest of the outside board
    '        ConstantDataToSend = String.Format("55AA{0}11", Hex(29))
    '        '
    '        strMsgCount = Me.Getm_Ctr
    '        '
    '        myDataToSend = String.Format("55AA{0}11{1}{2}{3}00000000000000000000000000000000{4}" _
    '                          , Hex(29), strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myInsideIntensity _
    '                          , Me.CalcChecksumToAsciiHexString(Me.myInsideIntensity & strCheckSum2))
    '        'work it out
    '        MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
    '        '
    '        Me.myColMessages.Add(MessageToSend)

    '        'send data
    '        For i As Integer = 1 To Me.myColMessages.Count
    '            Me.myComPort.Output(Me.myColMessages(i))
    '        Next
    '        If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError(False)
    '        '
    '    Catch

    '    End Try

    'End Sub

    Private Sub ClearOdds()

        Me.myColMessages.Clear()
        '
        Try
            Dim MessageToSend As String = ""
            'Dim ConstantDataToSend As String = ""
            Dim myDataToSend As String = ""
            'start of header
            Dim SOH As String = "01"
            Dim strPort As String
            'control, will send it always "on" from here
            Dim BoardControl As String
            'dimming, it should be done through the settings as well
            Dim BoardDimming As String = "0" & Hex(My.Settings.MiniBoardDimming)
            'pay load type, this value will be always 16 from here (hex(10))
            Dim PayloadType As String = "01"
            'end of transmission
            Dim EOT As String = "04"
            Dim BytesInPayload As String

            BytesInPayload = ("0" & Hex(21))
            BytesInPayload = BytesInPayload.Substring(BytesInPayload.Length - 2)

            strPort = ("0" & Hex(26))
            strPort = strPort.Substring(strPort.Length - 2)

            BoardControl = "01"

            myDataToSend = String.Format("{0}{1}{2}{3}{4}{5}0000000000000000000000000000000000000000{6}0000",
                                                 SOH, strPort, BoardControl, BoardDimming, PayloadType, BytesInPayload,
                                                 EOT)

            'work it out
            MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
            Me.myColMessages.Add(MessageToSend)

            'complete string with odds 7-12
            MessageToSend = ""
            myDataToSend = ""

            BytesInPayload = ("0" & Hex(17))
            BytesInPayload = BytesInPayload.Substring(BytesInPayload.Length - 2)

            strPort = ("0" & Hex(27))
            strPort = strPort.Substring(strPort.Length - 2)

            BoardControl = "01"

            myDataToSend = String.Format("{0}{1}{2}{3}{4}{5}00000000000000000000000000000000{6}0000",
                                                 SOH, strPort, BoardControl, BoardDimming, PayloadType, BytesInPayload,
                                                 EOT)

            'work it out
            MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
            '
            Me.myColMessages.Add(MessageToSend)

            'send data
            For i As Integer = 1 To Me.myColMessages.Count
                Me.myComPort.Output(Me.myColMessages(i))
            Next
            If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError(False)
            '
        Catch

        End Try

    End Sub

    'Private Sub ClearMiniBoard(ByVal payLoad As PayloadType)
    '    '40 is empty space in mini board collection off
    '    Select Case payLoad
    '        Case PayloadType.Odds
    '            Me.ClearMiniBoardOdds()
    '            '
    '        Case PayloadType.RunningOrder
    '            Me.ClearMiniBoardRO()
    '        Case PayloadType.WPS
    '            Me.ClearMiniBoardWPS()
    '            '
    '        Case PayloadType.allPayLoads
    '            Me.ClearMiniBoardOdds()
    '            Me.ClearMiniBoardWPS()
    '            Me.ClearMiniBoardTiming()
    '            '
    '        Case PayloadType.Timing
    '            Me.ClearMiniBoardTiming()
    '            '
    '    End Select

    'End Sub

    'Private Sub ClearMiniBoardWPS()
    '    'RO
    '    Me.myMiniComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend("01000100124040404040404040040000"))
    '    'W
    '    Me.myMiniComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend("010001001340404040404040404040404040040000"))
    '    'P
    '    Me.myMiniComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend("0100010014404040404040404040404040404040040000"))
    '    'S
    '    Me.myMiniComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend("01000100154040404040404040404040404040404040404040040000"))
    'End Sub

    'Private Sub ClearMiniBoardRO()
    '    Me.myMiniComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend("01000100124040404040404040040000"))
    'End Sub

    'Private Sub ClearMiniBoardOdds()
    '    Me.myMiniComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend("0100010010404040404040404040404040040000"))
    '    Me.myMiniComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend("01000100114040404040404040404040404040404040404040404040404040404040404040040000"))
    'End Sub

    'Private Sub ClearMiniBoardTiming()
    '    Me.myMiniComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend("010001001640404040404040404040404040404040404040404040404040040000"))
    'End Sub

    'Private Sub ClearItems(ByVal intPort As Integer)
    '    Dim ConstantDataToSend As String = ""
    '    Dim myDataToSend As String = ""
    '    Dim MessageToSend As String = ""

    '    Dim strPort As String
    '    strPort = ("0" & Hex(intPort))
    '    strPort = strPort.Substring(strPort.Length - 2)

    '    'constant data to send
    '    ConstantDataToSend = String.Format("55AA{0}11", strPort)
    '    '
    '    Dim strCheckSum As String = myInsideIntensity & "00000000000000000000000000000000"
    '    '
    '    Dim strMsgCount As String = Me.Getm_Ctr
    '    '
    '    myDataToSend = String.Format("55AA{0}11{1}{2}{3}00000000000000000000000000000000{4}" _
    '                          , strPort, strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myInsideIntensity _
    '                          , Me.CalcChecksumToAsciiHexString(strCheckSum))
    '    '
    '    MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
    '    '
    '    Me.myComPort.Output(MessageToSend)
    '    If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()
    '    ''
    'End Sub

    Private Sub ClearItems(ByVal intPort As Integer)
        Dim myDataToSend As String = ""
        Dim MessageToSend As String = ""

        'start of header
        Dim SOH As String = "01"
        Dim strPort As String
        'control, will send it always "on" from here
        Dim BoardControl As String = "01"
        'dimming, it should be done through the settings as well
        Dim BoardDimming As String = "0" & Hex(My.Settings.MiniBoardDimming)
        'pay load type, this value will be always 16 from here (hex(10))
        Dim PayloadType As String = "01"
        'end of transmission
        Dim EOT As String = "04"
        Dim BytesInPayload As String

        BytesInPayload = ("0" & Hex(21))
        BytesInPayload = BytesInPayload.Substring(BytesInPayload.Length - 2)

        strPort = ("0" & Hex(intPort))
        strPort = strPort.Substring(strPort.Length - 2)
        myDataToSend = String.Format("{0}{1}{2}{3}{4}{5}0000000000000000000000000000000000000000{6}0000",
                                                 SOH, strPort, BoardControl, BoardDimming, PayloadType,
                                                 BytesInPayload, EOT)

        MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
        '
        Me.myComPort.Output(MessageToSend)
        If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()
        ''
    End Sub

#End Region

#Region " Helper Methods "

    Private Sub InitializePorts(comPort As Integer)
        'Este programa usa el CommPort 1 para Tote
        'No tiene Teletimer, pero si lo tuviera fuera el CommPort 4

        Me.myComPort = New ComPort.Port
        'Me.myMiniComPort = New ComPort.Port
        '
        Me.myComPort.SetPortSettings(comPort, "19200,N,8,1") '2
        'Me.myMiniComPort.SetPortSettings(3, "19200,N,8,1") '3 9600
        ''
    End Sub

    Private Sub SetTransparent()
        Me.GroupBox1.BackColor = Color.Transparent
        Me.gbMTP.BackColor = Color.Transparent
        Me.GroupBox9.BackColor = Color.Transparent
        Me.gbTime.BackColor = Color.Transparent
        Me.gbTrack.BackColor = Color.Transparent
        Me.gbPostTime.BackColor = Color.Transparent
    End Sub

    Private Sub ShowError()
        'Me.ShowError(False)
        'InitializeCommPort()
    End Sub

    Private Sub ShowError(ByVal isMiniPort As Boolean)
        'If Not isMiniPort Then MessageBox.Show("Error sending message" & vbCrLf &
        '                     Me.myComPort.ErrorMessage, "Error Sending Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '
        'If isMiniPort Then MessageBox.Show("Error sending message" & vbCrLf &
        '                     Me.myMiniComPort.ErrorMessage, "Error Sending Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    Private Sub InitializeCommPort()
        Try

            Me.myComPort = Nothing
            'Me.myMiniComPort = Nothing

            Me.InitializePorts(2)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SetAndPopulateIntensity()
        Me.myInsideIntensity = My.Settings.InsideIntensity
        Me.myOutsideIntensity = My.Settings.OutsideIntensity
        ''
    End Sub

    Private Sub GetToteCompany()
        Me.myToteCompany = My.Settings.ToteCompany
    End Sub

    Private Sub SetAndPopulateOfficialTime()
        Me.myOfficialTimeSpan = 3
        For i As Integer = 1 To 10
            Me.cbOfficialTime.Items.Add(i)
        Next
        ''
    End Sub

    Private Sub PrepareMiniBoardDigitalData()
        Me.strMiniBoardCollectionOn = New Collection
        Me.strMiniBoardCollectionOff = New Collection
        'add details as previously calculated by Tania
        Me.strMiniBoardCollectionOn.Add("60", "KS") 'for empty space in mini board
        Dim valuesON() As String = {"20", "21", "22", "23", "24", "25", "26", "27", "28", "29"}
        '
        For i As Integer = 0 To valuesON.Count - 1
            Me.strMiniBoardCollectionOn.Add(valuesON(i), "K" & i)
        Next
        '
        Me.strMiniBoardCollectionOff.Add("40", "KS") 'for empty space in mini board
        Dim valuesOFF() As String = {"00", "01", "02", "03", "04", "05", "06", "07", "08", "09"}
        '
        For i As Integer = 0 To valuesOFF.Count - 1
            Me.strMiniBoardCollectionOff.Add(valuesOFF(i), "K" & i)
        Next
        ''
    End Sub

    Private Sub PrepareDigitalData()
        Me.strCollectionOff = New Collection
        Me.strCollectionOn = New Collection
        'add details based on provided trp documents
        Me.strCollectionOn.Add("80", "KS") 'for empty space in board
        Me.strCollectionOn.Add("B9", "K[")
        Me.strCollectionOn.Add("8F", "K]")
        Me.strCollectionOn.Add("F7", "KA")
        Me.strCollectionOn.Add("FC", "KB")
        Me.strCollectionOn.Add("D8", "KC")

        Dim valuesON() As String = {"BF", "86", "DB", "CF", "E6", "ED", "FD", "87", "FF", "E7"}
        '
        For i As Integer = 0 To valuesON.Count - 1
            Me.strCollectionOn.Add(valuesON(i), "K" & i)
        Next
        '
        Me.strCollectionOff.Add("00", "KS") 'for empty space in board
        Me.strCollectionOff.Add("39", "K[")
        Me.strCollectionOff.Add("0F", "K]")
        Me.strCollectionOff.Add("77", "KA")
        Me.strCollectionOff.Add("7C", "KB")
        Me.strCollectionOff.Add("58", "KC")

        Dim valuesOFF() As String = {"3F", "06", "5B", "4F", "66", "6D", "7D", "07", "7F", "67"}
        '
        For i As Integer = 0 To valuesOFF.Count - 1
            Me.strCollectionOff.Add(valuesOFF(i), "K" & i)
        Next
        ''
    End Sub

    Private Sub PrepareToteDisplayControler()
        Me.strToteDisplayCollection = New Collection

        Me.strToteDisplayCollection.Add("20", "KS") 'for empty space in mini board
        Me.strToteDisplayCollection.Add("5B", "K[")
        Me.strToteDisplayCollection.Add("5D", "K]")
        Me.strToteDisplayCollection.Add("41", "KA")
        Me.strToteDisplayCollection.Add("42", "KB")
        Me.strToteDisplayCollection.Add("43", "KC")

        Dim valuesOFF() As String = {"30", "31", "32", "33", "34", "35", "36", "37", "38", "39"}
        '
        For i As Integer = 0 To valuesOFF.Count - 1
            Me.strToteDisplayCollection.Add(valuesOFF(i), "K" & i)
        Next
        ''
    End Sub

    Private Sub PrepareSegmentHexPosition()
        Me.intSegmentHexPositionCollection = New Collection

        Dim valuesbyPosit() As String = {1, 2, 4, 8, 16, 32, 64, 128}
        '
        For i As Integer = 0 To valuesbyPosit.Count - 1
            Me.intSegmentHexPositionCollection.Add(valuesbyPosit(i), "K" & i + 1)
        Next
        ''
    End Sub
    Private Sub CheckForExistingInstance()
        If Process.GetProcessesByName(Process.GetCurrentProcess.ProcessName).Length > 1 Then
            MessageBox.Show("Another instance of this application is already running." & vbCrLf &
                             "Process cannot continue.", "Multiple Instances Forbidden", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Application.Exit()
        End If
        ''
        'KillCommSvr()
        ''
        'bool onlyInstance = false;
        '    System.Threading.Mutex mutex = new System.Threading.Mutex(true, @"Global\" + "RSI_ToteBoard", out onlyInstance);
        'If (!onlyInstance) Then
        '    {
        '        MessageBox.Show("Another instance is already running.");
        '  return;
        ' }

        '    Application.EnableVisualStyles();
        '    Application.SetCompatibleTextRenderingDefault(false);
        '    Application.Run(new RaceFx());
        '    GC.KeepAlive(mutex);

    End Sub

    Private Sub KillCommSvr()
        Dim intCtr As Integer = 0

        For Each p As Process In Process.GetProcesses()
            If (p.ProcessName.ToLower() = "rsiport") Then
                p.Kill()
                intCtr = intCtr + 1
            ElseIf (p.ProcessName.ToLower() = "rsiservice") Then
                p.Kill()
                intCtr = intCtr + 1
            ElseIf (intCtr = 2) Then
                Exit For
            End If
        Next
    End Sub

    Private Sub DisableAllTextBoxes()
        For Each ctlGB As Control In Me.Controls
            If TypeOf ctlGB Is Windows.Forms.GroupBox Then
                For Each ctltxt As Control In ctlGB.Controls
                    If TypeOf ctltxt Is Windows.Forms.TextBox Then
                        CType(ctltxt, Windows.Forms.TextBox).ReadOnly = True
                        CType(ctltxt, Windows.Forms.TextBox).TabStop = False
                    End If
                Next
            End If
        Next
    End Sub

    'Private Sub PrepareOddsDataToSend(ByVal ArrayOnOff As Array)
    '    Dim strPort As String
    '    'clear all messages
    '    Me.myColMessages.Clear()
    '    '
    '    'convert values to hex
    '    Dim RaceText1 As String = IIf(Me.txtRaceA.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtRaceA.Text.Trim = "", "S", Me.txtRaceA.Text.Trim)))
    '    Dim RaceText2 As String = IIf(Me.txtRaceB.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtRaceB.Text.Trim = "", "S", Me.txtRaceB.Text.Trim)))
    '    'MTP
    '    Dim mtpText1 As String = IIf(Me.txtMTPA.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtMTPA.Text.Trim = "", "S", Me.txtMTPA.Text.Trim)))
    '    Dim mtpText2 As String = IIf(Me.txtMTPB.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtMTPB.Text.Trim = "", "S", Me.txtMTPB.Text.Trim)))
    '    'odds 1 - 16
    '    Dim OddsText1 As String = IIf(Me.txtOdds1.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds1.Text.Trim = "", "S", Me.txtOdds1.Text.Trim)))
    '    Dim OddsText1a As String
    '    If ArrayOnOff(0) = 1 Then
    '        'use on collection
    '        OddsText1a = IIf(Me.txtOdds1A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtOdds1A.Text.Trim = "", "S", Me.txtOdds1A.Text.Trim)))
    '    Else
    '        'use off collection
    '        OddsText1a = IIf(Me.txtOdds1A.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds1A.Text.Trim = "", "S", Me.txtOdds1A.Text.Trim)))
    '    End If
    '    '
    '    Dim OddsText2 As String = IIf(Me.txtOdds2.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds2.Text.Trim = "", "S", Me.txtOdds2.Text.Trim)))
    '    Dim OddsText2a As String
    '    If ArrayOnOff(1) = 1 Then
    '        'use on collection
    '        OddsText2a = IIf(Me.txtOdds2A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtOdds2A.Text.Trim = "", "S", Me.txtOdds2A.Text.Trim)))
    '    Else
    '        'use off collection
    '        OddsText2a = IIf(Me.txtOdds2A.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds2A.Text.Trim = "", "S", Me.txtOdds2A.Text.Trim)))
    '    End If
    '    '
    '    Dim OddsText3 As String = IIf(Me.txtOdds3.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds3.Text.Trim = "", "S", Me.txtOdds3.Text.Trim)))
    '    Dim OddsText3a As String
    '    If ArrayOnOff(2) = 1 Then
    '        'use on collection
    '        OddsText3a = IIf(Me.txtOdds3A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtOdds3A.Text.Trim = "", "S", Me.txtOdds3A.Text.Trim)))
    '    Else
    '        'use off collection
    '        OddsText3a = IIf(Me.txtOdds3A.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds3A.Text.Trim = "", "S", Me.txtOdds3A.Text.Trim)))
    '    End If
    '    '
    '    Dim OddsText4 As String = IIf(Me.txtOdds4.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds4.Text.Trim = "", "S", Me.txtOdds4.Text.Trim)))
    '    Dim OddsText4a As String
    '    If ArrayOnOff(3) = 1 Then
    '        'use on collection
    '        OddsText4a = IIf(Me.txtOdds4A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtOdds4A.Text.Trim = "", "S", Me.txtOdds4A.Text.Trim)))
    '    Else
    '        'use off collection
    '        OddsText4a = IIf(Me.txtOdds4A.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds4A.Text.Trim = "", "S", Me.txtOdds4A.Text.Trim)))
    '    End If
    '    '
    '    Dim OddsText5 As String = IIf(Me.txtOdds5.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds5.Text.Trim = "", "S", Me.txtOdds5.Text.Trim)))
    '    Dim OddsText5a As String
    '    If ArrayOnOff(4) = 1 Then
    '        'use on collection
    '        OddsText5a = IIf(Me.txtOdds5A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtOdds5A.Text.Trim = "", "S", Me.txtOdds5A.Text.Trim)))
    '    Else
    '        'use off collection
    '        OddsText5a = IIf(Me.txtOdds5A.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds5A.Text.Trim = "", "S", Me.txtOdds5A.Text.Trim)))
    '    End If
    '    '
    '    Dim OddsText6 As String = IIf(Me.txtOdds6.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds6.Text.Trim = "", "S", Me.txtOdds6.Text.Trim)))
    '    Dim OddsText6a As String
    '    If ArrayOnOff(5) = 1 Then
    '        'use on collection
    '        OddsText6a = IIf(Me.txtOdds6A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtOdds6A.Text.Trim = "", "S", Me.txtOdds6A.Text.Trim)))
    '    Else
    '        'use off collection
    '        OddsText6a = IIf(Me.txtOdds6A.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds6A.Text.Trim = "", "S", Me.txtOdds6A.Text.Trim)))
    '    End If
    '    '
    '    Dim OddsText7 As String = IIf(Me.txtOdds7.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds7.Text.Trim = "", "S", Me.txtOdds7.Text.Trim)))
    '    Dim OddsText7a As String
    '    If ArrayOnOff(6) = 1 Then
    '        'use on collection
    '        OddsText7a = IIf(Me.txtOdds7A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtOdds7A.Text.Trim = "", "S", Me.txtOdds7A.Text.Trim)))
    '    Else
    '        'use off collection
    '        OddsText7a = IIf(Me.txtOdds7A.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds7A.Text.Trim = "", "S", Me.txtOdds7A.Text.Trim)))
    '    End If
    '    '
    '    Dim OddsText8 As String = IIf(Me.txtOdds8.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds8.Text.Trim = "", "S", Me.txtOdds8.Text.Trim)))
    '    Dim OddsText8a As String
    '    If ArrayOnOff(7) = 1 Then
    '        'use on collection
    '        OddsText8a = IIf(Me.txtOdds8A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtOdds8A.Text.Trim = "", "S", Me.txtOdds8A.Text.Trim)))
    '    Else
    '        'use off collection
    '        OddsText8a = IIf(Me.txtOdds8A.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds8A.Text.Trim = "", "S", Me.txtOdds8A.Text.Trim)))
    '    End If
    '    '
    '    Dim OddsText9 As String = IIf(Me.txtOdds9.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds9.Text.Trim = "", "S", Me.txtOdds9.Text.Trim)))
    '    Dim OddsText9a As String
    '    If ArrayOnOff(8) = 1 Then
    '        'use on collection
    '        OddsText9a = IIf(Me.txtOdds9A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtOdds9A.Text.Trim = "", "S", Me.txtOdds9A.Text.Trim)))
    '    Else
    '        'use off collection
    '        OddsText9a = IIf(Me.txtOdds9A.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds9A.Text.Trim = "", "S", Me.txtOdds9A.Text.Trim)))
    '    End If
    '    '
    '    Dim OddsText10 As String = IIf(Me.txtOdds10.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds10.Text.Trim = "", "S", Me.txtOdds10.Text.Trim)))
    '    Dim OddsText10a As String
    '    If ArrayOnOff(9) = 1 Then
    '        'use on collection
    '        OddsText10a = IIf(Me.txtOdds10A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtOdds10A.Text.Trim = "", "S", Me.txtOdds10A.Text.Trim)))
    '    Else
    '        'use off collection
    '        OddsText10a = IIf(Me.txtOdds10A.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds10A.Text.Trim = "", "S", Me.txtOdds10A.Text.Trim)))
    '    End If
    '    '
    '    Dim OddsText11 As String = IIf(Me.txtOdds11.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds11.Text.Trim = "", "S", Me.txtOdds11.Text.Trim)))
    '    Dim OddsText11a As String
    '    If ArrayOnOff(10) = 1 Then
    '        'use on collection
    '        OddsText11a = IIf(Me.txtOdds11A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtOdds11A.Text.Trim = "", "S", Me.txtOdds11A.Text.Trim)))
    '    Else
    '        'use off collection
    '        OddsText11a = IIf(Me.txtOdds11A.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds11A.Text.Trim = "", "S", Me.txtOdds11A.Text.Trim)))
    '    End If
    '    '
    '    Dim OddsText12 As String = IIf(Me.txtOdds12.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds12.Text.Trim = "", "S", Me.txtOdds12.Text.Trim)))
    '    Dim OddsText12a As String
    '    If ArrayOnOff(11) = 1 Then
    '        'use on collection
    '        OddsText12a = IIf(Me.txtOdds12A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtOdds12A.Text.Trim = "", "S", Me.txtOdds12A.Text.Trim)))
    '    Else
    '        'use off collection
    '        OddsText12a = IIf(Me.txtOdds12A.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds12A.Text.Trim = "", "S", Me.txtOdds12A.Text.Trim)))
    '    End If
    '    ''13
    '    'Dim OddsText13 As String = IIf(Me.txtOdds13.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds13.Text.Trim = "", "S", Me.txtOdds13.Text.Trim)))
    '    'Dim OddsText13a As String
    '    'If ArrayOnOff(12) = 1 Then
    '    '    'use on collection
    '    '    OddsText13a = IIf(Me.txtOdds13A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtOdds13A.Text.Trim = "", "S", Me.txtOdds13A.Text.Trim)))
    '    'Else
    '    '    'use off collection
    '    '    OddsText13a = IIf(Me.txtOdds13A.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds13A.Text.Trim = "", "S", Me.txtOdds13A.Text.Trim)))
    '    'End If
    '    ''14
    '    'Dim OddsText14 As String = IIf(Me.txtOdds14.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds14.Text.Trim = "", "S", Me.txtOdds14.Text.Trim)))
    '    'Dim OddsText14a As String
    '    'If ArrayOnOff(13) = 1 Then
    '    '    'use on collection
    '    '    OddsText14a = IIf(Me.txtOdds14A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtOdds14A.Text.Trim = "", "S", Me.txtOdds14A.Text.Trim)))
    '    'Else
    '    '    'use off collection
    '    '    OddsText14a = IIf(Me.txtOdds14A.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds14A.Text.Trim = "", "S", Me.txtOdds14A.Text.Trim)))
    '    'End If
    '    ''15
    '    'Dim OddsText15 As String = IIf(Me.txtOdds15.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds15.Text.Trim = "", "S", Me.txtOdds15.Text.Trim)))
    '    'Dim OddsText15a As String
    '    'If ArrayOnOff(14) = 1 Then
    '    '    'use on collection
    '    '    OddsText15a = IIf(Me.txtOdds15A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtOdds15A.Text.Trim = "", "S", Me.txtOdds15A.Text.Trim)))
    '    'Else
    '    '    'use off collection
    '    '    OddsText15a = IIf(Me.txtOdds15A.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds15A.Text.Trim = "", "S", Me.txtOdds15A.Text.Trim)))
    '    'End If
    '    ''16
    '    'Dim OddsText16 As String = IIf(Me.txtOdds16.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds16.Text.Trim = "", "S", Me.txtOdds16.Text.Trim)))
    '    'Dim OddsText16a As String
    '    'If ArrayOnOff(15) = 1 Then
    '    '    'use on collection
    '    '    OddsText16a = IIf(Me.txtOdds16A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtOdds16A.Text.Trim = "", "S", Me.txtOdds16A.Text.Trim)))
    '    'Else
    '    '    'use off collection
    '    '    OddsText16a = IIf(Me.txtOdds16A.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtOdds16A.Text.Trim = "", "S", Me.txtOdds16A.Text.Trim)))
    '    'End If
    '    '
    '    Dim strCheckSum1 As String = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}",
    '                                              RaceText1, RaceText2, mtpText1, mtpText2, OddsText1, OddsText1a,
    '                                              OddsText2, OddsText2a, OddsText3, OddsText3a, OddsText4, OddsText4a,
    '                                              OddsText5, OddsText5a, OddsText6, OddsText6a)
    '    '
    '    'Dim strCheckSum2 As String = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}", _
    '    '                                          OddsText7, OddsText7a, _
    '    '                                          OddsText8, OddsText8a, OddsText9, OddsText9a, OddsText10, OddsText10a, _
    '    '                                          OddsText11, OddsText11a, OddsText12, OddsText12a, _
    '    '                                          OddsText13, OddsText13a, OddsText14, OddsText14a)

    '    Dim strCheckSum2 As String = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}",
    '                                              OddsText7, OddsText7a,
    '                                              OddsText8, OddsText8a, OddsText9, OddsText9a, OddsText10, OddsText10a,
    '                                              OddsText11, OddsText11a, OddsText12, OddsText12a)
    '    '
    '    'Dim strCheckSum3 As String = String.Format("{0}{1}{2}{3}000000000000000000000000", _
    '    '                                          OddsText15, OddsText15a, _
    '    '                                          OddsText16, OddsText16a)
    '    Try
    '        Dim MessageToSend As String = ""
    '        Dim ConstantDataToSend As String = ""
    '        Dim strMsgCount As String
    '        Dim myDataToSend As String = ""
    '        'set constant for inside board
    '        ConstantDataToSend = String.Format("55AA{0}11", Hex(26))
    '        strMsgCount = Me.Getm_Ctr
    '        'complete string with odds 1-6
    '        myDataToSend = String.Format("55AA{0}11{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}{19}{20}" _
    '                          , Hex(26), strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myOutsideIntensity _
    '                          , RaceText1, RaceText2, mtpText1, mtpText2 _
    '                          , OddsText1, OddsText1a, OddsText2, OddsText2a, OddsText3, OddsText3a _
    '                          , OddsText4, OddsText4a, OddsText5, OddsText5a, OddsText6, OddsText6a _
    '                          , Me.CalcChecksumToAsciiHexString(Me.myOutsideIntensity & strCheckSum1))
    '        'work it out
    '        MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
    '        '
    '        Me.myColMessages.Add(MessageToSend)

    '        'do same thing for inside board, port and intensity change
    '        If (My.Settings.HasInsideBoard) Then
    '            MessageToSend = ""
    '            myDataToSend = ""
    '            ConstantDataToSend = ""
    '            '
    '            ConstantDataToSend = String.Format("55AA{0}11", Hex(28))
    '            '
    '            strMsgCount = Me.Getm_Ctr
    '            'complete string with odds 1-6
    '            myDataToSend = String.Format("55AA{0}11{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}{19}{20}" _
    '                              , Hex(28), strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myInsideIntensity _
    '                              , RaceText1, RaceText2, mtpText1, mtpText2 _
    '                              , OddsText1, OddsText1a, OddsText2, OddsText2a, OddsText3, OddsText3a _
    '                              , OddsText4, OddsText4a, OddsText5, OddsText5a, OddsText6, OddsText6a _
    '                              , Me.CalcChecksumToAsciiHexString(Me.myInsideIntensity & strCheckSum1))
    '            'work it out
    '            MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
    '            '
    '            Me.myColMessages.Add(MessageToSend)
    '        End If

    '        'now complete string with odds 7-14 for inside board
    '        MessageToSend = ""
    '        myDataToSend = ""
    '        ConstantDataToSend = ""
    '        '
    '        ConstantDataToSend = String.Format("55AA{0}11", Hex(27))
    '        '
    '        strMsgCount = Me.Getm_Ctr
    '        '
    '        'myDataToSend = String.Format("55AA{0}11{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}{19}{20}" _
    '        '                  , Hex(27), strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myOutsideIntensity _
    '        '                  , OddsText7, OddsText7a, OddsText8, OddsText8a, OddsText9, OddsText9a _
    '        '                  , OddsText10, OddsText10a, OddsText11, OddsText11a, OddsText12, OddsText12a _
    '        '                  , OddsText13, OddsText13a, OddsText14, OddsText14a _
    '        '                  , Me.CalcChecksumToAsciiHexString(Me.myOutsideIntensity & strCheckSum2))

    '        myDataToSend = String.Format("55AA{0}11{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}00000000{16}" _
    '                          , Hex(27), strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myOutsideIntensity _
    '                          , OddsText7, OddsText7a, OddsText8, OddsText8a, OddsText9, OddsText9a _
    '                          , OddsText10, OddsText10a, OddsText11, OddsText11a, OddsText12, OddsText12a _
    '                          , Me.CalcChecksumToAsciiHexString(Me.myOutsideIntensity & strCheckSum2))
    '        'work it out
    '        MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
    '        '
    '        Me.myColMessages.Add(MessageToSend)

    '        'inside board
    '        If (My.Settings.HasInsideBoard) Then
    '            MessageToSend = ""
    '            myDataToSend = ""
    '            ConstantDataToSend = ""
    '            '
    '            ConstantDataToSend = String.Format("55AA{0}11", Hex(29))
    '            '
    '            strMsgCount = Me.Getm_Ctr
    '            '
    '            'myDataToSend = String.Format("55AA{0}11{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}{19}{20}" _
    '            '                  , Hex(29), strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myInsideIntensity _
    '            '                  , OddsText7, OddsText7a, OddsText8, OddsText8a, OddsText9, OddsText9a _
    '            '                  , OddsText10, OddsText10a, OddsText11, OddsText11a, OddsText12, OddsText12a _
    '            '                  , OddsText13, OddsText13a, OddsText14, OddsText14a _
    '            '                  , Me.CalcChecksumToAsciiHexString(Me.myInsideIntensity & strCheckSum2))

    '            myDataToSend = String.Format("55AA{0}11{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}00000000{16}" _
    '                              , Hex(29), strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myInsideIntensity _
    '                              , OddsText7, OddsText7a, OddsText8, OddsText8a, OddsText9, OddsText9a _
    '                              , OddsText10, OddsText10a, OddsText11, OddsText11a, OddsText12, OddsText12a _
    '                              , Me.CalcChecksumToAsciiHexString(Me.myInsideIntensity & strCheckSum2))
    '            'work it out
    '            MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
    '            '
    '            Me.myColMessages.Add(MessageToSend)
    '        End If

    '        ''outside board odds 15 - 16
    '        'MessageToSend = ""
    '        'myDataToSend = ""
    '        'ConstantDataToSend = ""
    '        ''
    '        'strPort = ("0" & Hex(10))
    '        'strPort = strPort.Substring(strPort.Length - 2)

    '        'ConstantDataToSend = String.Format("55AA{0}11", strPort)
    '        ''
    '        'strMsgCount = Me.Getm_Ctr
    '        ''
    '        'myDataToSend = String.Format("55AA{0}11{1}{2}{3}{4}{5}{6}{7}000000000000000000000000{8}" _
    '        '                  , strPort, strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myOutsideIntensity _
    '        '                  , OddsText15, OddsText15a, OddsText16, OddsText16a _
    '        '                  , Me.CalcChecksumToAsciiHexString(Me.myOutsideIntensity & strCheckSum3))
    '        ''work it out
    '        'MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
    '        ''
    '        'Me.myColMessages.Add(MessageToSend)

    '        ''inside board
    '        'If (My.Settings.HasInsideBoard) Then
    '        '    MessageToSend = ""
    '        '    myDataToSend = ""
    '        '    ConstantDataToSend = ""
    '        '    '
    '        '    strPort = ("0" & Hex(11))
    '        '    strPort = strPort.Substring(strPort.Length - 2)

    '        '    ConstantDataToSend = String.Format("55AA{0}11", strPort)
    '        '    '
    '        '    strMsgCount = Me.Getm_Ctr
    '        '    '
    '        '    myDataToSend = String.Format("55AA{0}11{1}{2}{3}{4}{5}{6}{7}000000000000000000000000{8}" _
    '        '                      , strPort, strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myInsideIntensity _
    '        '                      , OddsText15, OddsText15a, OddsText16, OddsText16a _
    '        '                      , Me.CalcChecksumToAsciiHexString(Me.myInsideIntensity & strCheckSum3))
    '        '    'work it out
    '        '    MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
    '        '    '
    '        '    Me.myColMessages.Add(MessageToSend)
    '        '    '
    '        'End If

    '    Catch ex As Exception
    '        'nothing
    '    End Try
    '    ''
    'End Sub
    Private Sub PrepareOddsDataToSend(ByVal ArrayOnOff As Array)
        'clear all messages
        Me.myColMessages.Clear()

        'start of header
        Dim SOH As String = "01"
        Dim strPort As String
        'control, will send it always "on" from here
        Dim BoardControl As String
        'dimming, it should be done through the settings as well
        Dim BoardDimming As String = "0" & Hex(My.Settings.MiniBoardDimming)
        'pay load type, this value will be always 16 from here (hex(10))
        Dim PayloadType As String = "01"
        'end of transmission
        Dim EOT As String = "04"
        Dim BytesInPayload As String

        Dim Flash9to16 As String = "00"
        Dim Flash1to8 As String = "00"

        'H segment 16-9 hhhhhhhh
        'H segment  8-1 hhhhhhhh

        Dim intHsegment9to16_1 As Integer = 0
        Dim intHsegment1to8_1 As Integer = 0
        Dim Hsegment9to16_1 As String = "00"
        Dim Hsegment1to8_1 As String = "00"

        Dim intHsegment9to16_2 As Integer = 0
        Dim intHsegment1to8_2 As Integer = 0
        Dim Hsegment9to16_2 As String = "00"
        Dim Hsegment1to8_2 As String = "00"

        ''Race
        'Dim RaceText1 As String = PrepareHexData(Me.txtRaceA.Text.Trim)
        'Dim RaceText2 As String = PrepareHexData(Me.txtRaceB.Text.Trim)
        ''MTP
        'Dim mtpText1 As String = PrepareHexData(Me.txtMTPA.Text.Trim)
        'Dim mtpText2 As String = PrepareHexData(Me.txtMTPB.Text.Trim)

        'odds 1 - 16
        Dim OddsText1 As String = PrepareHexData(Me.txtOdds1.Text.Trim)
        Dim OddsText1a As String = PrepareHexData(Me.txtOdds1A.Text.Trim)
        If ArrayOnOff(0) = 1 Then
            'OddsText1a is in segment 2, if "/" on then 00000010 or 2d
            intHsegment1to8_1 = intHsegment1to8_1 + Me.intSegmentHexPositionCollection("K" & 2)
        End If
        '
        Dim OddsText2 As String = PrepareHexData(Me.txtOdds2.Text.Trim)
        Dim OddsText2a As String = PrepareHexData(Me.txtOdds2A.Text.Trim)
        If ArrayOnOff(1) = 1 Then
            'OddsText2a is in segment 4, if "/" on then 00001000 or 8d
            intHsegment1to8_1 = intHsegment1to8_1 + Me.intSegmentHexPositionCollection("K" & 4)
        End If
        '
        Dim OddsText3 As String = PrepareHexData(Me.txtOdds3.Text.Trim)
        Dim OddsText3a As String = PrepareHexData(Me.txtOdds3A.Text.Trim)
        If ArrayOnOff(2) = 1 Then
            'OddsText3a is in segment 6, if "/" on then 00100000 or 32d
            intHsegment1to8_1 = intHsegment1to8_1 + Me.intSegmentHexPositionCollection("K" & 6)
        End If
        '
        Dim OddsText4 As String = PrepareHexData(Me.txtOdds4.Text.Trim)
        Dim OddsText4a As String = PrepareHexData(Me.txtOdds4A.Text.Trim)
        If ArrayOnOff(3) = 1 Then
            'OddsText4a is in segment 8, if "/" on then 10000000 or 128d
            intHsegment1to8_1 = intHsegment1to8_1 + Me.intSegmentHexPositionCollection("K" & 8)
        End If
        '
        Dim OddsText5 As String = PrepareHexData(Me.txtOdds5.Text.Trim)
        Dim OddsText5a As String = PrepareHexData(Me.txtOdds5A.Text.Trim)
        If ArrayOnOff(4) = 1 Then
            'OddsText5a is in segment 2, if "/" on then 00000010 or 2d
            intHsegment9to16_1 = intHsegment9to16_1 + +Me.intSegmentHexPositionCollection("K" & 2)
        End If
        '
        Dim OddsText6 As String = PrepareHexData(Me.txtOdds6.Text.Trim)
        Dim OddsText6a As String = PrepareHexData(Me.txtOdds6A.Text.Trim)
        If ArrayOnOff(5) = 1 Then
            'OddsText6a is in segment 4, if "/" on then 00001000 or 8d
            intHsegment9to16_1 = intHsegment9to16_1 + Me.intSegmentHexPositionCollection("K" & 4)
        End If
        '
        Dim OddsText7 As String = PrepareHexData(Me.txtOdds7.Text.Trim)
        Dim OddsText7a As String = PrepareHexData(Me.txtOdds7A.Text.Trim)
        If ArrayOnOff(6) = 1 Then
            'OddsText7a is in segment 2, if "/" on then 00000010 or 2d
            intHsegment1to8_2 = intHsegment1to8_2 + Me.intSegmentHexPositionCollection("K" & 2)
        End If
        '
        Dim OddsText8 As String = PrepareHexData(Me.txtOdds8.Text.Trim)
        Dim OddsText8a As String = PrepareHexData(Me.txtOdds8A.Text.Trim)
        If ArrayOnOff(7) = 1 Then
            'OddsText8a is in segment 4, if "/" on then 00001000 or 8d
            intHsegment1to8_2 = intHsegment1to8_2 + Me.intSegmentHexPositionCollection("K" & 4)
        End If
        '
        Dim OddsText9 As String = PrepareHexData(Me.txtOdds9.Text.Trim)
        Dim OddsText9a As String = PrepareHexData(Me.txtOdds9A.Text.Trim)
        If ArrayOnOff(8) = 1 Then
            'OddsText9a is in segment 6, if "/" on then 00100000 or 32d
            intHsegment1to8_2 = intHsegment1to8_2 + Me.intSegmentHexPositionCollection("K" & 6)
        End If
        '
        Dim OddsText10 As String = PrepareHexData(Me.txtOdds10.Text.Trim)
        Dim OddsText10a As String = PrepareHexData(Me.txtOdds10A.Text.Trim)
        If ArrayOnOff(9) = 1 Then
            'OddsText10a is in segment 8, if "/" on then 10000000 or 128d
            intHsegment1to8_2 = intHsegment1to8_2 + Me.intSegmentHexPositionCollection("K" & 8)
        End If
        '
        Dim OddsText11 As String = PrepareHexData(Me.txtOdds11.Text.Trim)
        Dim OddsText11a As String = PrepareHexData(Me.txtOdds11A.Text.Trim)
        If ArrayOnOff(10) = 1 Then
            'OddsText11a is in segment 2, if "/" on then 00000010 or 2d
            intHsegment9to16_2 = intHsegment9to16_2 + Me.intSegmentHexPositionCollection("K" & 2)
        End If
        '
        Dim OddsText12 As String = PrepareHexData(Me.txtOdds12.Text.Trim)
        Dim OddsText12a As String = PrepareHexData(Me.txtOdds12A.Text.Trim)
        If ArrayOnOff(11) = 1 Then
            'OddsText12a is in segment 4, if "/" on then 00001000 or 8d
            intHsegment9to16_2 = intHsegment9to16_2 + Me.intSegmentHexPositionCollection("K" & 4)
        End If

        Try
            Dim MessageToSend As String = ""
            Dim myDataToSend As String = ""

            strPort = ("0" & Hex(26))
            strPort = strPort.Substring(strPort.Length - 2)
            BytesInPayload = ("0" & Hex(17))
            BytesInPayload = BytesInPayload.Substring(BytesInPayload.Length - 2)
            Hsegment9to16_1 = ("0" & Hex(intHsegment9to16_1))
            Hsegment9to16_1 = Hsegment9to16_1.Substring(Hsegment9to16_1.Length - 2)
            Hsegment1to8_1 = ("0" & Hex(intHsegment1to8_1))
            Hsegment1to8_1 = Hsegment1to8_1.Substring(Hsegment1to8_1.Length - 2)

            BoardControl = "01"

            'myDataToSend = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}00000000{10}{11}{12}{13}{14}{15}{16}{17}{18}{19}{20}{21}{22}0000",
            '                                     SOH, strPort, BoardControl, BoardDimming, PayloadType,
            '                                     BytesInPayload, Flash9to16, Flash1to8, Hsegment9to16_1, Hsegment1to8_1,
            '                                     OddsText1, OddsText1a, OddsText2, OddsText2a, OddsText3, OddsText3a,
            '                                     OddsText4, OddsText4a, OddsText5, OddsText5a, OddsText6, OddsText6a,
            '                                     EOT)

            myDataToSend = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}{19}{20}{21}{22}0000",
                                                 SOH, strPort, BoardControl, BoardDimming, PayloadType,
                                                 BytesInPayload, Flash9to16, Flash1to8, Hsegment9to16_1, Hsegment1to8_1,
                                                 OddsText1, OddsText1a, OddsText2, OddsText2a, OddsText3, OddsText3a,
                                                 OddsText4, OddsText4a, OddsText5, OddsText5a, OddsText6, OddsText6a,
                                                 EOT)

            'work it out
            MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
            '
            Me.myColMessages.Add(MessageToSend)

            'now complete string with odds 7-14 for inside board
            MessageToSend = ""
            myDataToSend = ""

            strPort = ("0" & Hex(27))
            strPort = strPort.Substring(strPort.Length - 2)
            BytesInPayload = ("0" & Hex(17))
            BytesInPayload = BytesInPayload.Substring(BytesInPayload.Length - 2)
            Hsegment9to16_2 = ("0" & Hex(intHsegment9to16_2))
            Hsegment9to16_2 = Hsegment9to16_2.Substring(Hsegment9to16_2.Length - 2)
            Hsegment1to8_2 = ("0" & Hex(intHsegment1to8_2))
            Hsegment1to8_2 = Hsegment1to8_2.Substring(Hsegment1to8_2.Length - 2)

            BoardControl = "01"

            myDataToSend = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}{19}{20}{21}{22}0000",
                                                 SOH, strPort, BoardControl, BoardDimming, PayloadType,
                                                 BytesInPayload, Flash9to16, Flash1to8, Hsegment9to16_2, Hsegment1to8_2,
                                                 OddsText7, OddsText7a, OddsText8, OddsText8a, OddsText9, OddsText9a,
                                                 OddsText10, OddsText10a, OddsText11, OddsText11a, OddsText12, OddsText12a,
                                                 EOT)

            'work it out
            MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
            '
            Me.myColMessages.Add(MessageToSend)

        Catch ex As Exception
            'nothing
        End Try
        ''
    End Sub

    Private Function PrepareHexData(ByVal txtFromTextBox As String) As String
        Dim txtText As String
        Try
            txtText = IIf(txtFromTextBox = "", Me.strToteDisplayCollection("KS"), Me.strToteDisplayCollection("K" & txtFromTextBox))
        Catch ex As Exception
            txtText = Me.strToteDisplayCollection("KS")
        End Try
        Return txtText
    End Function

    'Private Sub PrepareMiniBoardOddsDataToSend(ByVal ArrayOnOff As Array)
    '    'clear all messages
    '    Me.myMiniColMessages.Clear()

    '    'set constants

    '    'start of header
    '    Dim SOH As String = "01"
    '    'address, lets get it fom the settings, tania will later decide how to approach this
    '    Dim BoardAddress As String = "00" 'My.Settings.MiniBoardAddress
    '    'control, will send it always "on" from here
    '    Dim BoardControl As String = "01"
    '    'dimming, it should be done through the settings as well
    '    Dim BoardDimming As String = "0" & Hex(My.Settings.MiniBoardDimming)
    '    'pay load type, this value will be always 16 from here (hex(10))
    '    Dim PayloadType As String = "10"
    '    'end of transmission
    '    Dim EOT As String = "04"

    '    'convert values to hex
    '    Dim RaceText1 As String = If(Me.txtRaceA.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                             Me.strMiniBoardCollectionOff("K" & Me.txtRaceA.Text.Trim))
    '    Dim RaceText2 As String = If(Me.txtRaceB.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                             Me.strMiniBoardCollectionOff("K" & Me.txtRaceB.Text.Trim))
    '    'MTP
    '    Dim mtpText1 As String = If(Me.txtMTPA.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                           Me.strMiniBoardCollectionOff("K" & Me.txtMTPA.Text.Trim))
    '    Dim mtpText2 As String = If(Me.txtMTPB.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                           Me.strMiniBoardCollectionOff("K" & Me.txtMTPB.Text.Trim))
    '    'time of day (TOD)
    '    Dim TODTexta As String = If(Me.txtTODA.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                           Me.strMiniBoardCollectionOff("K" & Me.txtTODA.Text.Trim))
    '    'if we have time, turn on the colon
    '    Dim TODTextb As String = If(Me.txtTODB.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                           Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtTODB.Text.Trim)) + &H20))

    '    Dim TODTextc As String = If(Me.txtTODC.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                           Me.strMiniBoardCollectionOff("K" & Me.txtTODC.Text.Trim))
    '    Dim TODTextd As String = If(Me.txtTODD.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                           Me.strMiniBoardCollectionOff("K" & Me.txtTODD.Text.Trim))
    '    'post time
    '    Dim PostTimea As String = If(Me.txtPostTimea.Text.Trim = "" OrElse Me.txtPostTimea.Text.Trim = "0", Me.strMiniBoardCollectionOff("KS"),
    '                                                                 Me.strMiniBoardCollectionOff("K" & Me.txtPostTimea.Text.Trim))
    '    'turn on the colon here as well
    '    Dim PostTimeb As String = If(Me.txtPostTimeb.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                                 Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtPostTimeb.Text.Trim)) + &H20))
    '    Dim PostTimec As String = If(Me.txtPostTimec.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                                 Me.strMiniBoardCollectionOff("K" & Me.txtPostTimec.Text.Trim))
    '    Dim PostTimed As String = If(Me.txtPostTimed.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                                 Me.strMiniBoardCollectionOff("K" & Me.txtPostTimed.Text.Trim))

    '    'build the string to send
    '    Dim strToSend As String = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}{19}",
    '                                             SOH, BoardAddress, BoardControl, BoardDimming, PayloadType,
    '                                             RaceText1, RaceText2, mtpText1, mtpText2,
    '                                             TODTexta, TODTextb, TODTextc, TODTextd,
    '                                             PostTimea, PostTimeb, PostTimec, PostTimed, EOT, Me.CalculateCRC(1), Me.CalculateCRC(2))
    '    '
    '    Me.myMiniColMessages.Add(Me.myCommSvr.oCommServerNet.GetMessageToSend(strToSend))


    '    '***************************************************************************************************************************
    '    PayloadType = "11"
    '    'odds 1 - 16
    '    Dim OddsText1a As String = If(Me.txtOdds1A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                             Me.strMiniBoardCollectionOff("K" & Me.txtOdds1A.Text.Trim))
    '    Dim OddsText1 As String
    '    If ArrayOnOff(0) = 1 Then
    '        'use on collection
    '        OddsText1 = If(Me.txtOdds1.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                     Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtOdds1.Text.Trim)) + &H20))
    '    Else
    '        'use off collection
    '        OddsText1 = If(Me.txtOdds1.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                     Me.strMiniBoardCollectionOff("K" & Me.txtOdds1.Text.Trim))
    '    End If
    '    '
    '    Dim OddsText2a As String = If(Me.txtOdds2A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                             Me.strMiniBoardCollectionOff("K" & Me.txtOdds2A.Text.Trim))
    '    Dim OddsText2 As String
    '    If ArrayOnOff(1) = 1 Then
    '        'use on collection
    '        OddsText2 = If(Me.txtOdds2.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                   Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtOdds2.Text.Trim)) + &H20))
    '    Else
    '        'use off collection
    '        OddsText2 = If(Me.txtOdds2.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                      Me.strMiniBoardCollectionOff("K" & Me.txtOdds2.Text.Trim))
    '    End If
    '    '
    '    Dim OddsText3a As String = If(Me.txtOdds3A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                             Me.strMiniBoardCollectionOff("K" & Me.txtOdds3A.Text.Trim))
    '    Dim OddsText3 As String
    '    If ArrayOnOff(2) = 1 Then
    '        'use on collection
    '        OddsText3 = If(Me.txtOdds3.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                     Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtOdds3.Text.Trim)) + &H20))
    '    Else
    '        'use off collection
    '        OddsText3 = If(Me.txtOdds3.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                     Me.strMiniBoardCollectionOff("K" & Me.txtOdds3.Text.Trim))
    '    End If
    '    '
    '    Dim OddsText4a As String = If(Me.txtOdds4A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                             Me.strMiniBoardCollectionOff("K" & Me.txtOdds4A.Text.Trim))
    '    Dim OddsText4 As String
    '    If ArrayOnOff(3) = 1 Then
    '        'use on collection
    '        OddsText4 = If(Me.txtOdds4.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                     Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtOdds4.Text.Trim)) + &H20))
    '    Else
    '        'use off collection
    '        OddsText4 = If(Me.txtOdds4.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                     Me.strMiniBoardCollectionOff("K" & Me.txtOdds4.Text.Trim))
    '    End If

    '    Dim OddsText5a As String = If(Me.txtOdds5A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                             Me.strMiniBoardCollectionOff("K" & Me.txtOdds5A.Text.Trim))
    '    Dim OddsText5 As String
    '    If ArrayOnOff(4) = 1 Then
    '        'use on collection
    '        OddsText5 = If(Me.txtOdds5.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                     Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtOdds5.Text.Trim)) + &H20))
    '    Else
    '        'use off collection
    '        OddsText5 = If(Me.txtOdds5.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                     Me.strMiniBoardCollectionOff("K" & Me.txtOdds5.Text.Trim))
    '    End If
    '    '
    '    Dim OddsText6a As String = If(Me.txtOdds6A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                             Me.strMiniBoardCollectionOff("K" & Me.txtOdds6A.Text.Trim))
    '    Dim OddsText6 As String
    '    If ArrayOnOff(5) = 1 Then
    '        'use on collection
    '        OddsText6 = If(Me.txtOdds6.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                     Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtOdds6.Text.Trim)) + &H20))
    '    Else
    '        'use off collection
    '        OddsText6 = If(Me.txtOdds6.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                     Me.strMiniBoardCollectionOff("K" & Me.txtOdds6.Text.Trim))
    '    End If
    '    '
    '    Dim OddsText7a As String = If(Me.txtOdds7A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                             Me.strMiniBoardCollectionOff("K" & Me.txtOdds7A.Text.Trim))
    '    Dim OddsText7 As String
    '    If ArrayOnOff(6) = 1 Then
    '        'use on collection
    '        OddsText7 = If(Me.txtOdds7.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                     Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtOdds7.Text.Trim)) + &H20))
    '    Else
    '        'use off collection
    '        OddsText7 = If(Me.txtOdds7.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                     Me.strMiniBoardCollectionOff("K" & Me.txtOdds7.Text.Trim))
    '    End If
    '    '
    '    Dim OddsText8a As String = If(Me.txtOdds8A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                             Me.strMiniBoardCollectionOff("K" & Me.txtOdds8A.Text.Trim))
    '    Dim OddsText8 As String
    '    If ArrayOnOff(7) = 1 Then
    '        'use on collection
    '        OddsText8 = If(Me.txtOdds8.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                     Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtOdds8.Text.Trim)) + &H20))
    '    Else
    '        'use off collection
    '        OddsText8 = If(Me.txtOdds8.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                     Me.strMiniBoardCollectionOff("K" & Me.txtOdds8.Text.Trim))
    '    End If
    '    '
    '    Dim OddsText9a As String = If(Me.txtOdds9A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                             Me.strMiniBoardCollectionOff("K" & Me.txtOdds9A.Text.Trim))
    '    Dim OddsText9 As String
    '    If ArrayOnOff(8) = 1 Then
    '        'use on collection
    '        OddsText9 = If(Me.txtOdds9.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                     Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtOdds9.Text.Trim)) + &H20))
    '    Else
    '        'use off collection
    '        OddsText9 = If(Me.txtOdds9.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                     Me.strMiniBoardCollectionOff("K" & Me.txtOdds9.Text.Trim))
    '    End If
    '    '
    '    Dim OddsText10a As String = If(Me.txtOdds10A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                               Me.strMiniBoardCollectionOff("K" & Me.txtOdds10A.Text.Trim))
    '    Dim OddsText10 As String
    '    If ArrayOnOff(9) = 1 Then
    '        'use on collection
    '        OddsText10 = If(Me.txtOdds10.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                       Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtOdds10.Text.Trim)) + &H20))
    '    Else
    '        'use off collection
    '        OddsText10 = If(Me.txtOdds10.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                       Me.strMiniBoardCollectionOff("K" & Me.txtOdds10.Text.Trim))
    '    End If
    '    '
    '    Dim OddsText11a As String = If(Me.txtOdds11A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                               Me.strMiniBoardCollectionOff("K" & Me.txtOdds11A.Text.Trim))
    '    Dim OddsText11 As String
    '    If ArrayOnOff(10) = 1 Then
    '        'use on collection
    '        OddsText11 = If(Me.txtOdds11.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                       Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtOdds11.Text.Trim)) + &H20))
    '    Else
    '        'use off collection
    '        OddsText11 = If(Me.txtOdds11.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                       Me.strMiniBoardCollectionOff("K" & Me.txtOdds11.Text.Trim))
    '    End If
    '    '
    '    Dim OddsText12a As String = If(Me.txtOdds12A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                               Me.strMiniBoardCollectionOff("K" & Me.txtOdds12A.Text.Trim))
    '    Dim OddsText12 As String
    '    If ArrayOnOff(11) = 1 Then
    '        'use on collection
    '        OddsText12 = If(Me.txtOdds12.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                       Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtOdds12.Text.Trim)) + &H20))
    '    Else
    '        'use off collection
    '        OddsText12 = If(Me.txtOdds12.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                        Me.strMiniBoardCollectionOff("K" & Me.txtOdds12.Text.Trim))
    '    End If
    '    '
    '    Dim OddsText13a As String = If(Me.txtOdds13A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                               Me.strMiniBoardCollectionOff("K" & Me.txtOdds13A.Text.Trim))
    '    Dim OddsText13 As String
    '    If ArrayOnOff(12) = 1 Then
    '        'use on collection
    '        OddsText13 = If(Me.txtOdds13.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                       Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtOdds13.Text.Trim)) + &H20))
    '    Else
    '        'use off collection
    '        OddsText13 = If(Me.txtOdds13.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                        Me.strMiniBoardCollectionOff("K" & Me.txtOdds13.Text.Trim))
    '    End If
    '    '
    '    Dim OddsText14a As String = If(Me.txtOdds14A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                               Me.strMiniBoardCollectionOff("K" & Me.txtOdds14A.Text.Trim))
    '    Dim OddsText14 As String
    '    If ArrayOnOff(13) = 1 Then
    '        'use on collection
    '        OddsText14 = If(Me.txtOdds14.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                       Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtOdds14.Text.Trim)) + &H20))
    '    Else
    '        'use off collection
    '        OddsText14 = If(Me.txtOdds14.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                        Me.strMiniBoardCollectionOff("K" & Me.txtOdds14.Text.Trim))
    '    End If
    '    '
    '    Dim OddsText15a As String = If(Me.txtOdds15A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                               Me.strMiniBoardCollectionOff("K" & Me.txtOdds15A.Text.Trim))
    '    Dim OddsText15 As String
    '    If ArrayOnOff(14) = 1 Then
    '        'use on collection
    '        OddsText15 = If(Me.txtOdds15.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                       Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtOdds15.Text.Trim)) + &H20))
    '    Else
    '        'use off collection
    '        OddsText15 = If(Me.txtOdds15.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                        Me.strMiniBoardCollectionOff("K" & Me.txtOdds15.Text.Trim))
    '    End If
    '    '
    '    Dim OddsText16a As String = If(Me.txtOdds16A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                               Me.strMiniBoardCollectionOff("K" & Me.txtOdds16A.Text.Trim))
    '    Dim OddsText16 As String
    '    If ArrayOnOff(15) = 1 Then
    '        'use on collection
    '        OddsText16 = If(Me.txtOdds16.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                       Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtOdds16.Text.Trim)) + &H20))
    '    Else
    '        'use off collection
    '        OddsText16 = If(Me.txtOdds16.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                                                        Me.strMiniBoardCollectionOff("K" & Me.txtOdds16.Text.Trim))
    '    End If
    '    '

    '    '****************************************************************************************************************************************************
    '    'build the string to send

    '    strToSend = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}{19}{20}{21}{22}{23}{24}{25}{26}{27}" _
    '                              & "{28}{29}{30}{31}{32}{33}{34}{35}{36}{37}{38}{39}",
    '                                             SOH, BoardAddress, BoardControl, BoardDimming, PayloadType,
    '                                             OddsText1, OddsText1a, OddsText2, OddsText2a,
    '                                             OddsText3, OddsText3a, OddsText4, OddsText4a,
    '                                             OddsText5, OddsText5a, OddsText6, OddsText6a,
    '                                             OddsText7, OddsText7a, OddsText8, OddsText8a,
    '                                             OddsText9, OddsText9a, OddsText10, OddsText10a,
    '                                             OddsText11, OddsText11a, OddsText12, OddsText12a,
    '                                             OddsText13, OddsText13a, OddsText14, OddsText14a,
    '                                             OddsText15, OddsText15a, OddsText16, OddsText16a,
    '                                             EOT, Me.CalculateCRC(1), Me.CalculateCRC(2))
    '    '
    '    Me.myMiniColMessages.Add(Me.myCommSvr.oCommServerNet.GetMessageToSend(strToSend))
    '    ''
    'End Sub

    Private Sub PrepareRunningOrderDataToSend()
        Dim ConstantDataToSend As String = ""
        Dim myDataToSend As String = ""
        Dim MessageToSend As String = ""

        'start of header
        Dim SOH As String = "01"
        Dim strPort As String
        'control, will send it always "on" from here
        Dim BoardControl As String = "01"
        'dimming, it should be done through the settings as well
        Dim BoardDimming As String = "0" & Hex(My.Settings.MiniBoardDimming)
        'pay load type, this value will be always 16 from here (hex(10))
        Dim PayloadType As String = "01"
        'end of transmission
        Dim EOT As String = "04"
        Dim BytesInPayload As String

        BytesInPayload = ("0" & Hex(21))
        BytesInPayload = BytesInPayload.Substring(BytesInPayload.Length - 2)

        Dim Flash9to16 As String = "00"
        Dim Flash1to8 As String = "00"
        Dim intFlash1to8 As Integer = 0

        'H segment 16-9 hhhhhhhh
        'H segment  8-1 hhhhhhhh

        Dim Hsegment9to16_1 As String = "00"
        Dim intHsegment1to8_1 As Integer = 0
        Dim Hsegment1to8_1 As String = "00"
        '

        Dim ro1 As String = PrepareHexData(Me.txtRunning1.Text.Trim)
        Dim ro1a As String = PrepareHexData(Me.txtRunning1A.Text.Trim)
        If Me.chkResults1.Checked Then
            'ro1 is in segment 1, 00000001 or 1d
            'intHsegment1to8_1 = intHsegment1to8_1 + Me.intSegmentHexPositionCollection("K" & 1)

            intFlash1to8 += 3 '3 = first (1) and second (2) digits

        End If
        '
        Dim ro2 As String = PrepareHexData(Me.txtRunning2.Text.Trim)
        Dim ro2a As String = PrepareHexData(Me.txtRunning2A.Text.Trim)
        If Me.chkResults2.Checked Then
            'ro2 is in segment 3, 00000100 or 4d
            'intHsegment1to8_1 = intHsegment1to8_1 + Me.intSegmentHexPositionCollection("K" & 3)

            intFlash1to8 += 12 '12 = third (4) and fourth (8) digits
        End If
        '        
        Dim ro3 As String = PrepareHexData(Me.txtRunning3.Text.Trim)
        Dim ro3a As String = PrepareHexData(Me.txtRunning3A.Text.Trim)
        If Me.chkResults3.Checked Then
            'ro3 is in segment 5, 00010000 or 16d
            'intHsegment1to8_1 = intHsegment1to8_1 + Me.intSegmentHexPositionCollection("K" & 5)

            intFlash1to8 += 48 '48 = fifth (16) and sixth (32) digits
        End If
        'n
        Dim ro4 As String = PrepareHexData(Me.txtRunning4.Text.Trim)
        Dim ro4a As String = PrepareHexData(Me.txtRunning4A.Text.Trim)
        If Me.chkResults4.Checked Then
            'ro4 is in segment 7, 01000000 or 64d
            'intHsegment1to8_1 = intHsegment1to8_1 + Me.intSegmentHexPositionCollection("K" & 7)

            intFlash1to8 += 172 '172 = seventh (64) and eighth (128) digits
        End If
        '
        Try
            strPort = ("0" & Hex(14))
            strPort = strPort.Substring(strPort.Length - 2)

            Hsegment1to8_1 = ("0" & Hex(intHsegment1to8_1))

            Flash1to8 = ("0" & Hex(intFlash1to8))

            myDataToSend = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}0000000000000000{18}0000",
                                                 SOH, strPort, BoardControl, BoardDimming, PayloadType,
                                                 BytesInPayload, Flash9to16, Flash1to8, Hsegment9to16_1, Hsegment1to8_1,
                                                 ro1, ro1a, ro2, ro2a,
                                                 ro3, ro3a, ro4, ro4a,
                                                 EOT)

            'convert it to char
            MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
            '
            Me.myComPort.Output(MessageToSend)
            If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()
            '
        Catch ex As Exception
            '
        End Try

        If myCommSvr.m_blnClearExotics Then
            'If Me.timerOfficial.Enabled Then
            If (myCommSvr.p_blnResultsOut) Then
                ClearResults(False, True)
            End If
            'InitializeCommPort() 'Tania ClearToteBoard
            'Else
            '    ClearResults(False, False)
            'End If
        End If
        myCommSvr.m_blnClearExotics = False
        ''
    End Sub

    'Private Sub PrepareMiniRunningOrderDataToSend()
    '    'set constants

    '    'start of header
    '    Dim SOH As String = "01"
    '    'address, lets get it fom the settings, tania will later decide how to approach this
    '    Dim BoardAddress As String = "00" 'My.Settings.MiniBoardAddress
    '    'control, will send it always "on" from here
    '    Dim BoardControl As String = "01"
    '    'dimming, it should be done through the settings as well
    '    Dim BoardDimming As String = "0" & Hex(My.Settings.MiniBoardDimming)
    '    'pay load type, this value will be always 16 from here (hex(10))
    '    Dim PayloadType As String = "12"
    '    'end of transmission
    '    Dim EOT As String = "04"
    '    '
    '    Dim ro1 As String = If(Me.txtRunning1.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                           Me.strMiniBoardCollectionOff("K" & Me.txtRunning1.Text.Trim))
    '    Dim ro2 As String = If(Me.txtRunning2.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                           Me.strMiniBoardCollectionOff("K" & Me.txtRunning2.Text.Trim))
    '    Dim ro3 As String = If(Me.txtRunning3.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                           Me.strMiniBoardCollectionOff("K" & Me.txtRunning3.Text.Trim))
    '    Dim ro4 As String = If(Me.txtRunning4.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                           Me.strMiniBoardCollectionOff("K" & Me.txtRunning4.Text.Trim))
    '    '
    '    Dim ro1a As String
    '    Dim ro2a As String
    '    Dim ro3a As String
    '    Dim ro4a As String
    '    If Me.chkResults1.Checked Then
    '        ro1a = If(Me.txtRunning1A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '              Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtRunning1A.Text.Trim)) + &H20))
    '    Else
    '        ro1a = If(Me.txtRunning1A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '              Me.strMiniBoardCollectionOff("K" & Me.txtRunning1A.Text.Trim))
    '    End If
    '    '
    '    If Me.chkResults2.Checked Then
    '        ro2a = If(Me.txtRunning2A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '              Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtRunning2A.Text.Trim)) + &H20))
    '    Else
    '        ro2a = If(Me.txtRunning2A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '              Me.strMiniBoardCollectionOff("K" & Me.txtRunning2A.Text.Trim))
    '    End If
    '    'this will turn on "official" on board
    '    If Me.chkResults3.Checked Then
    '        ro3a = If(Me.txtRunning3A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '              Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtRunning3A.Text.Trim)) + &H20))
    '    Else
    '        ro3a = If(Me.txtRunning3A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '              Me.strMiniBoardCollectionOff("K" & Me.txtRunning3A.Text.Trim))
    '    End If
    '    '
    '    If Me.chkResults4.Checked Then
    '        ro4a = If(Me.txtRunning4A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '              Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtRunning4A.Text.Trim)) + &H20))
    '    Else
    '        ro4a = If(Me.txtRunning4A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '              Me.strMiniBoardCollectionOff("K" & Me.txtRunning4A.Text.Trim))
    '    End If
    '    '
    '    'build the string to send
    '    Dim MessageToSend As String = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}",
    '                                             SOH, BoardAddress, BoardControl, BoardDimming, PayloadType,
    '                                             ro1, ro1a, ro2, ro2a,
    '                                             ro3, ro3a, ro4, ro4a,
    '                                             EOT, Me.CalculateCRC(1), Me.CalculateCRC(2))
    '    '
    '    Me.myMiniComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend(MessageToSend))
    '    If Not Me.myMiniComPort.ErrorMessage = "" Then Me.ShowError(True)
    '    'send it to WIN for official, dh,inq
    '    Me.PrepareMiniWPSDataToSend(True)

    '    ''
    'End Sub

    'Private Sub PrepareWPSDataToSend()
    '    Dim ConstantDataToSend As String = ""
    '    Dim myDataToSend As String = ""
    '    Dim MessageToSend As String = ""
    '    Dim strCheckSum As String = ""

    '    Dim strPort As String
    '    strPort = ("0" & Hex(15))
    '    strPort = strPort.Substring(strPort.Length - 2)

    '    'constant data to send
    '    ConstantDataToSend = String.Format("55AA{0}11", strPort)
    '    'first win
    '    Dim WIN1a As String = IIf(Me.txtWin1A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtWin1A.Text.Trim = "", "S", Me.txtWin1A.Text.Trim)))
    '    Dim WIN1b As String = IIf(Me.txtWin1B.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtWin1B.Text.Trim = "", "S", Me.txtWin1B.Text.Trim)))
    '    Dim WIN1c As String = IIf(Me.txtWin1C.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtWin1C.Text.Trim = "", "S", Me.txtWin1C.Text.Trim)))
    '    Dim WIN1d As String = IIf(Me.txtWin1D.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtWin1D.Text.Trim = "", "S", Me.txtWin1D.Text.Trim)))
    '    Dim WIN1e As String = IIf(Me.txtWin1E.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtWin1E.Text.Trim = "", "S", Me.txtWin1E.Text.Trim)))
    '    Dim WIN1f As String = IIf(Me.txtWin1F.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtWin1F.Text.Trim = "", "S", Me.txtWin1F.Text.Trim)))
    '    'second win
    '    Dim WIN2a As String = IIf(Me.txtWin2A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtWin2A.Text.Trim = "", "S", Me.txtWin2A.Text.Trim)))
    '    Dim WIN2b As String = IIf(Me.txtWin2B.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtWin2B.Text.Trim = "", "S", Me.txtWin2B.Text.Trim)))
    '    Dim WIN2c As String = IIf(Me.txtWin2C.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtWin2C.Text.Trim = "", "S", Me.txtWin2C.Text.Trim)))
    '    Dim WIN2d As String = IIf(Me.txtWin2D.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtWin2D.Text.Trim = "", "S", Me.txtWin2D.Text.Trim)))
    '    Dim WIN2e As String = IIf(Me.txtWin2E.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtWin2E.Text.Trim = "", "S", Me.txtWin2E.Text.Trim)))
    '    Dim WIN2f As String = IIf(Me.txtWin2F.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtWin2F.Text.Trim = "", "S", Me.txtWin2F.Text.Trim)))
    '    '
    '    strCheckSum = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}00000000", _
    '                                              WIN1a, WIN1b, WIN1c, WIN1d, WIN1e, WIN1f, _
    '                                              WIN2a, WIN2b, WIN2c, WIN2d, WIN2e, WIN2f)
    '    '
    '    Dim strMsgCount As String = Me.Getm_Ctr
    '    '
    '    myDataToSend = String.Format("55AA{0}11{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}00000000{16}" _
    '                          , strPort, strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myOutsideIntensity _
    '                          , WIN1a, WIN1b, WIN1c, WIN1d, WIN1e, WIN1f _
    '                          , WIN2a, WIN2b, WIN2c, WIN2d, WIN2e, WIN2f _
    '                          , Me.CalcChecksumToAsciiHexString(Me.myOutsideIntensity & strCheckSum))
    '    'convert it to char
    '    MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
    '    '
    '    Me.myComPort.Output(MessageToSend)
    '    If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()

    '    'inside board
    '    If (My.Settings.HasInsideBoard) Then
    '        strPort = ("0" & Hex(31))
    '        strPort = strPort.Substring(strPort.Length - 2)
    '        ConstantDataToSend = String.Format("55AA{0}11", strPort)
    '        strMsgCount = Me.Getm_Ctr
    '        '
    '        myDataToSend = String.Format("55AA{0}11{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}00000000{16}" _
    '                              , strPort, strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myInsideIntensity _
    '                              , WIN1a, WIN1b, WIN1c, WIN1d, WIN1e, WIN1f _
    '                              , WIN2a, WIN2b, WIN2c, WIN2d, WIN2e, WIN2f _
    '                              , Me.CalcChecksumToAsciiHexString(Me.myInsideIntensity & strCheckSum))
    '        'convert it to char
    '        MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
    '        '
    '        Me.myComPort.Output(MessageToSend)
    '        If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()
    '    End If

    '    'PLACE
    '    strPort = ("0" & Hex(16))
    '    strPort = strPort.Substring(strPort.Length - 2)
    '    '
    '    'constant data to send
    '    ConstantDataToSend = String.Format("55AA{0}11", strPort)
    '    'first place
    '    Dim PLACE1a As String = IIf(Me.txtPlace1A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPlace1A.Text.Trim = "", "S", Me.txtPlace1A.Text.Trim)))
    '    Dim PLACE1b As String = IIf(Me.txtPlace1B.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPlace1B.Text.Trim = "", "S", Me.txtPlace1B.Text.Trim)))
    '    Dim PLACE1c As String = IIf(Me.txtPlace1C.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPlace1C.Text.Trim = "", "S", Me.txtPlace1C.Text.Trim)))
    '    Dim PLACE1d As String = IIf(Me.txtPlace1D.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPlace1D.Text.Trim = "", "S", Me.txtPlace1D.Text.Trim)))
    '    Dim PLACE1e As String = IIf(Me.txtPlace1E.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPlace1E.Text.Trim = "", "S", Me.txtPlace1E.Text.Trim)))
    '    'second place
    '    Dim PLACE2a As String = IIf(Me.txtPlace2A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPlace2A.Text.Trim = "", "S", Me.txtPlace2A.Text.Trim)))
    '    Dim PLACE2b As String = IIf(Me.txtPlace2B.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPlace2B.Text.Trim = "", "S", Me.txtPlace2B.Text.Trim)))
    '    Dim PLACE2c As String = IIf(Me.txtPlace2C.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPlace2C.Text.Trim = "", "S", Me.txtPlace2C.Text.Trim)))
    '    Dim PLACE2d As String = IIf(Me.txtPlace2D.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPlace2D.Text.Trim = "", "S", Me.txtPlace2D.Text.Trim)))
    '    Dim PLACE2e As String = IIf(Me.txtPlace2E.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPlace2E.Text.Trim = "", "S", Me.txtPlace2E.Text.Trim)))
    '    '3rd place
    '    Dim PLACE3a As String = IIf(Me.txtPlace3A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPlace3A.Text.Trim = "", "S", Me.txtPlace3A.Text.Trim)))
    '    Dim PLACE3b As String = IIf(Me.txtPlace3B.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPlace3B.Text.Trim = "", "S", Me.txtPlace3B.Text.Trim)))
    '    Dim PLACE3c As String = IIf(Me.txtPlace3C.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPlace3C.Text.Trim = "", "S", Me.txtPlace3C.Text.Trim)))
    '    Dim PLACE3d As String = IIf(Me.txtPlace3D.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPlace3D.Text.Trim = "", "S", Me.txtPlace3D.Text.Trim)))
    '    Dim PLACE3e As String = IIf(Me.txtPlace3E.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPlace3E.Text.Trim = "", "S", Me.txtPlace3E.Text.Trim)))
    '    '
    '    strCheckSum = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}00", _
    '                                              PLACE1a, PLACE1b, PLACE1c, PLACE1d, PLACE1e, _
    '                                              PLACE2a, PLACE2b, PLACE2c, PLACE2d, PLACE2e, _
    '                                              PLACE3a, PLACE3b, PLACE3c, PLACE3d, PLACE3e)
    '    '
    '    strMsgCount = Me.Getm_Ctr
    '    '
    '    myDataToSend = String.Format("55AA{0}11{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}00{19}" _
    '                          , strPort, strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myOutsideIntensity _
    '                          , PLACE1a, PLACE1b, PLACE1c, PLACE1d, PLACE1e _
    '                          , PLACE2a, PLACE2b, PLACE2c, PLACE2d, PLACE2e _
    '                          , PLACE3a, PLACE3b, PLACE3c, PLACE3d, PLACE3e _
    '                          , Me.CalcChecksumToAsciiHexString(Me.myOutsideIntensity & strCheckSum))
    '    'convert it to char
    '    MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
    '    '
    '    Me.myComPort.Output(MessageToSend)
    '    If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()

    '    'inside board
    '    If (My.Settings.HasInsideBoard) Then
    '        strPort = ("0" & Hex(32))
    '        strPort = strPort.Substring(strPort.Length - 2)
    '        ConstantDataToSend = String.Format("55AA{0}11", strPort)
    '        strMsgCount = Me.Getm_Ctr
    '        '
    '        myDataToSend = String.Format("55AA{0}11{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}00{19}" _
    '                              , strPort, strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myInsideIntensity _
    '                              , PLACE1a, PLACE1b, PLACE1c, PLACE1d, PLACE1e _
    '                              , PLACE2a, PLACE2b, PLACE2c, PLACE2d, PLACE2e _
    '                              , PLACE3a, PLACE3b, PLACE3c, PLACE3d, PLACE3e _
    '                              , Me.CalcChecksumToAsciiHexString(Me.myInsideIntensity & strCheckSum))
    '        'convert it to char
    '        MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
    '        '
    '        Me.myComPort.Output(MessageToSend)
    '        If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()
    '    End If

    '    'SHOW
    '    strPort = ("0" & Hex(17))
    '    strPort = strPort.Substring(strPort.Length - 2)

    '    ConstantDataToSend = String.Format("55AA{0}11", strPort)
    '    'show1
    '    Dim SHOW1a As String = IIf(Me.txtShow1A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtShow1A.Text.Trim = "", "S", Me.txtShow1A.Text.Trim)))
    '    Dim SHOW1b As String = IIf(Me.txtShow1B.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtShow1B.Text.Trim = "", "S", Me.txtShow1B.Text.Trim)))
    '    Dim SHOW1c As String = IIf(Me.txtShow1C.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtShow1C.Text.Trim = "", "S", Me.txtShow1C.Text.Trim)))
    '    Dim SHOW1d As String = IIf(Me.txtShow1D.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtShow1D.Text.Trim = "", "S", Me.txtShow1D.Text.Trim)))
    '    Dim SHOW1e As String = IIf(Me.txtShow1E.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtShow1E.Text.Trim = "", "S", Me.txtShow1E.Text.Trim)))
    '    'show2
    '    Dim SHOW2a As String = IIf(Me.txtShow2A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtShow2A.Text.Trim = "", "S", Me.txtShow2A.Text.Trim)))
    '    Dim SHOW2b As String = IIf(Me.txtShow2B.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtShow2B.Text.Trim = "", "S", Me.txtShow2B.Text.Trim)))
    '    Dim SHOW2c As String = IIf(Me.txtShow2C.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtShow2C.Text.Trim = "", "S", Me.txtShow2C.Text.Trim)))
    '    Dim SHOW2d As String = IIf(Me.txtShow2D.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtShow2D.Text.Trim = "", "S", Me.txtShow2D.Text.Trim)))
    '    Dim SHOW2e As String = IIf(Me.txtShow2E.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtShow2E.Text.Trim = "", "S", Me.txtShow2E.Text.Trim)))
    '    'show3
    '    Dim SHOW3a As String = IIf(Me.txtShow3A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtShow3A.Text.Trim = "", "S", Me.txtShow3A.Text.Trim)))
    '    Dim SHOW3b As String = IIf(Me.txtShow3B.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtShow3B.Text.Trim = "", "S", Me.txtShow3B.Text.Trim)))
    '    Dim SHOW3c As String = IIf(Me.txtShow3C.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtShow3C.Text.Trim = "", "S", Me.txtShow3C.Text.Trim)))
    '    Dim SHOW3d As String = IIf(Me.txtShow3D.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtShow3D.Text.Trim = "", "S", Me.txtShow3D.Text.Trim)))
    '    Dim SHOW3e As String = IIf(Me.txtShow3E.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtShow3E.Text.Trim = "", "S", Me.txtShow3E.Text.Trim)))
    '    'show4 (Used on controler 34.)
    '    Dim SHOW4a As String = IIf(Me.txtShow4A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtShow4A.Text.Trim = "", "S", Me.txtShow4A.Text.Trim)))
    '    Dim SHOW4b As String = IIf(Me.txtShow4B.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtShow4B.Text.Trim = "", "S", Me.txtShow4B.Text.Trim)))
    '    Dim SHOW4c As String = IIf(Me.txtShow4C.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtShow4C.Text.Trim = "", "S", Me.txtShow4C.Text.Trim)))
    '    Dim SHOW4d As String = IIf(Me.txtShow4D.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtShow4D.Text.Trim = "", "S", Me.txtShow4D.Text.Trim)))
    '    Dim SHOW4e As String = IIf(Me.txtShow4E.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtShow4E.Text.Trim = "", "S", Me.txtShow4E.Text.Trim)))
    '    '
    '    strCheckSum = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}00", _
    '                                              SHOW1a, SHOW1b, SHOW1c, SHOW1d, SHOW1e, _
    '                                              SHOW2a, SHOW2b, SHOW2c, SHOW2d, SHOW2e, _
    '                                              SHOW3a, SHOW3b, SHOW3c, SHOW3d, SHOW3e)
    '    strMsgCount = Me.Getm_Ctr
    '    myDataToSend = String.Format("55AA{0}11{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}00{19}" _
    '                          , strPort, strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myOutsideIntensity _
    '                          , SHOW1a, SHOW1b, SHOW1c, SHOW1d, SHOW1e _
    '                          , SHOW2a, SHOW2b, SHOW2c, SHOW2d, SHOW2e _
    '                          , SHOW3a, SHOW3b, SHOW3c, SHOW3d, SHOW3e _
    '                          , Me.CalcChecksumToAsciiHexString(Me.myOutsideIntensity & strCheckSum))
    '    'convert it to char
    '    MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
    '    Me.myComPort.Output(MessageToSend)

    '    If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()

    '    'inside board
    '    If (My.Settings.HasInsideBoard) Then
    '        '
    '    End If

    '    '******************************************************************
    '    'SHOW4- controller 33 is maxed out.  therefore, 34 is used here to show 5 digits only.
    '    'enrique: ctrl#34: fourth row of the show race results (WPS)
    '    strPort = ("0" & Hex(34))
    '    strPort = strPort.Substring(strPort.Length - 2)
    '    ConstantDataToSend = String.Format("55AA{0}11", strPort)

    '    strCheckSum = String.Format("{0}{1}{2}{3}{4}0000000000000000000000", _
    '                                SHOW4a, SHOW4b, SHOW4c, SHOW4d, SHOW4e)
    '    strMsgCount = Me.Getm_Ctr
    '    myDataToSend = String.Format("55AA{0}11{1}{2}{3}{4}{5}{6}{7}{8}0000000000000000000000{9}" _
    '                                 , strPort, strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myOutsideIntensity _
    '                                 , SHOW4a, SHOW4b, SHOW4c, SHOW4d, SHOW4e _
    '                                 , Me.CalcChecksumToAsciiHexString(Me.myOutsideIntensity & strCheckSum))

    '    'convert it to char
    '    MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
    '    '******************************************************************
    '    Me.myComPort.Output(MessageToSend)

    '    If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()

    '    'inside board
    '    If (My.Settings.HasInsideBoard) Then
    '        '
    '    End If
    '    'strPort = ("0" & Hex(33))
    '    'strPort = strPort.Substring(strPort.Length - 2)
    '    'ConstantDataToSend = String.Format("55AA{0}11", strPort)
    '    'strMsgCount = Me.Getm_Ctr
    '    'myDataToSend = String.Format("55AA{0}11{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}00" _
    '    '                      , strPort, strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myInsideIntensity _
    '    '                      , SHOW1a, SHOW1b, SHOW1c, SHOW1d, SHOW1e _
    '    '                      , SHOW2a, SHOW2b, SHOW2c, SHOW2d, SHOW2e _
    '    '                      , SHOW3a, SHOW3b, SHOW3c, SHOW3d, SHOW3e _
    '    '                      , Me.CalcChecksumToAsciiHexString(Me.myInsideIntensity & strCheckSum))
    '    ''convert it to char
    '    'MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
    '    'Me.myComPort.Output(MessageToSend)


    '    'If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()
    '    ''

    'End Sub

    Private Sub PrepareWPSDataToSend()
        Dim myDataToSend As String = ""
        Dim MessageToSend As String = ""

        'start of header
        Dim SOH As String = "01"
        Dim strPort As String
        'control, will send it always "on" from here
        Dim BoardControl As String = "01"
        'dimming, it should be done through the settings as well
        Dim BoardDimming As String = "0" & Hex(My.Settings.MiniBoardDimming)
        'pay load type, this value will be always 16 from here (hex(10))
        Dim PayloadType As String = "01"
        'end of transmission
        Dim EOT As String = "04"
        Dim BytesInPayload As String

        BytesInPayload = ("0" & Hex(21))
        BytesInPayload = BytesInPayload.Substring(BytesInPayload.Length - 2)

        'first win
        Dim WIN1a As String = PrepareHexData(Me.txtWin1A.Text.Trim)
        Dim WIN1b As String = PrepareHexData(Me.txtWin1B.Text.Trim)
        Dim WIN1c As String = PrepareHexData(Me.txtWin1C.Text.Trim)
        Dim WIN1d As String = PrepareHexData(Me.txtWin1D.Text.Trim)
        Dim WIN1e As String = PrepareHexData(Me.txtWin1E.Text.Trim)
        Dim WIN1f As String = PrepareHexData(Me.txtWin1F.Text.Trim)
        'second win
        Dim WIN2a As String = PrepareHexData(Me.txtWin2A.Text.Trim)
        Dim WIN2b As String = PrepareHexData(Me.txtWin2B.Text.Trim)
        Dim WIN2c As String = PrepareHexData(Me.txtWin2C.Text.Trim)
        Dim WIN2d As String = PrepareHexData(Me.txtWin2D.Text.Trim)
        Dim WIN2e As String = PrepareHexData(Me.txtWin2E.Text.Trim)
        Dim WIN2f As String = PrepareHexData(Me.txtWin2F.Text.Trim)
        '
        'first place
        Dim PLACE1a As String = PrepareHexData(Me.txtPlace1A.Text.Trim)
        Dim PLACE1b As String = PrepareHexData(Me.txtPlace1B.Text.Trim)
        Dim PLACE1c As String = PrepareHexData(Me.txtPlace1C.Text.Trim)
        Dim PLACE1d As String = PrepareHexData(Me.txtPlace1D.Text.Trim)
        Dim PLACE1e As String = PrepareHexData(Me.txtPlace1E.Text.Trim)
        'second place
        Dim PLACE2a As String = PrepareHexData(Me.txtPlace2A.Text.Trim)
        Dim PLACE2b As String = PrepareHexData(Me.txtPlace2B.Text.Trim)
        Dim PLACE2c As String = PrepareHexData(Me.txtPlace2C.Text.Trim)
        Dim PLACE2d As String = PrepareHexData(Me.txtPlace2D.Text.Trim)
        Dim PLACE2e As String = PrepareHexData(Me.txtPlace2E.Text.Trim)
        '3rd place
        Dim PLACE3a As String = PrepareHexData(Me.txtPlace3A.Text.Trim)
        Dim PLACE3b As String = PrepareHexData(Me.txtPlace3B.Text.Trim)
        Dim PLACE3c As String = PrepareHexData(Me.txtPlace3C.Text.Trim)
        Dim PLACE3d As String = PrepareHexData(Me.txtPlace3D.Text.Trim)
        Dim PLACE3e As String = PrepareHexData(Me.txtPlace3E.Text.Trim)
        '
        'show1
        Dim SHOW1a As String = PrepareHexData(Me.txtShow1A.Text.Trim)
        Dim SHOW1b As String = PrepareHexData(Me.txtShow1B.Text.Trim)
        Dim SHOW1c As String = PrepareHexData(Me.txtShow1C.Text.Trim)
        Dim SHOW1d As String = PrepareHexData(Me.txtShow1D.Text.Trim)
        Dim SHOW1e As String = PrepareHexData(Me.txtShow1E.Text.Trim)
        'show2
        Dim SHOW2a As String = PrepareHexData(Me.txtShow2A.Text.Trim)
        Dim SHOW2b As String = PrepareHexData(Me.txtShow2B.Text.Trim)
        Dim SHOW2c As String = PrepareHexData(Me.txtShow2C.Text.Trim)
        Dim SHOW2d As String = PrepareHexData(Me.txtShow2D.Text.Trim)
        Dim SHOW2e As String = PrepareHexData(Me.txtShow2E.Text.Trim)
        'show3
        Dim SHOW3a As String = PrepareHexData(Me.txtShow3A.Text.Trim)
        Dim SHOW3b As String = PrepareHexData(Me.txtShow3B.Text.Trim)
        Dim SHOW3c As String = PrepareHexData(Me.txtShow3C.Text.Trim)
        Dim SHOW3d As String = PrepareHexData(Me.txtShow3D.Text.Trim)
        Dim SHOW3e As String = PrepareHexData(Me.txtShow3E.Text.Trim)
        'show4 (Used on controler 34.)
        Dim SHOW4a As String = PrepareHexData(Me.txtShow4A.Text.Trim)
        Dim SHOW4b As String = PrepareHexData(Me.txtShow4B.Text.Trim)
        Dim SHOW4c As String = PrepareHexData(Me.txtShow4C.Text.Trim)
        Dim SHOW4d As String = PrepareHexData(Me.txtShow4D.Text.Trim)
        Dim SHOW4e As String = PrepareHexData(Me.txtShow4E.Text.Trim)
        '

        Try
            'Win
            strPort = ("0" & Hex(15))
            strPort = strPort.Substring(strPort.Length - 2)

            myDataToSend = String.Format("{0}{1}{2}{3}{4}{5}00000000{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}00000000{18}0000",
                                                 SOH, strPort, BoardControl, BoardDimming, PayloadType, BytesInPayload,
                                                 WIN1a, WIN1b, WIN1c, WIN1d, WIN1e, WIN1f,
                                                 WIN2a, WIN2b, WIN2c, WIN2d, WIN2e, WIN2f,
                                                 EOT)

            'convert it to char
            MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
            '
            Me.myComPort.Output(MessageToSend)
            If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()
            '
            'PLACE
            strPort = ("0" & Hex(16))
            strPort = strPort.Substring(strPort.Length - 2)

            myDataToSend = String.Format("{0}{1}{2}{3}{4}{5}00000000{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}{19}{20}00{21}0000",
                                                 SOH, strPort, BoardControl, BoardDimming, PayloadType, BytesInPayload,
                                                 PLACE1a, PLACE1b, PLACE1c, PLACE1d, PLACE1e,
                                                 PLACE2a, PLACE2b, PLACE2c, PLACE2d, PLACE2e,
                                                 PLACE3a, PLACE3b, PLACE3c, PLACE3d, PLACE3e,
                                                 EOT)

            'convert it to char
            MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
            '
            Me.myComPort.Output(MessageToSend)
            If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()

            'SHOW
            strPort = ("0" & Hex(17))
            strPort = strPort.Substring(strPort.Length - 2)

            myDataToSend = String.Format("{0}{1}{2}{3}{4}{5}00000000{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}{19}{20}00{21}0000",
                                                 SOH, strPort, BoardControl, BoardDimming, PayloadType, BytesInPayload,
                                                 SHOW1a, SHOW1b, SHOW1c, SHOW1d, SHOW1e,
                                                 SHOW2a, SHOW2b, SHOW2c, SHOW2d, SHOW2e,
                                                 SHOW3a, SHOW3b, SHOW3c, SHOW3d, SHOW3e,
                                                 EOT)
            'convert it to char
            MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
            '
            Me.myComPort.Output(MessageToSend)

            If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()
            '
            '******************************************************************
            'SHOW4- controller 33 is maxed out.  therefore, 34 is used here to show 5 digits only.
            'enrique: ctrl#34: fourth row of the show race results (WPS)
            strPort = ("0" & Hex(34))
            strPort = strPort.Substring(strPort.Length - 2)

            myDataToSend = String.Format("{0}{1}{2}{3}{4}{5}00000000{6}{7}{8}{9}{10}0000000000000000000000{11}0000",
                                                 SOH, strPort, BoardControl, BoardDimming, PayloadType, BytesInPayload,
                                                 SHOW4a, SHOW4b, SHOW4c, SHOW4d, SHOW4e,
                                                 EOT)

            'convert it to char
            MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
            '
            Me.myComPort.Output(MessageToSend)

            If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()
            '
        Catch ex As Exception
            '
        End Try

    End Sub

    Private Sub PrepareTimeOfDayToSend()
        Dim myDataToSend As String = ""
        Dim MessageToSend As String = ""

        'start of header
        Dim SOH As String = "01"
        Dim strPort As String
        'control, will send it always "on" from here
        Dim BoardControl As String = "01"
        'dimming, it should be done through the settings as well
        Dim BoardDimming As String = "0" & Hex(My.Settings.MiniBoardDimming)
        'pay load type, this value will be always 16 from here (hex(10))
        Dim PayloadType As String = "01"
        'end of transmission
        Dim EOT As String = "04"
        Dim BytesInPayload As String

        'Race
        Dim RaceText1 As String = PrepareHexData(Me.txtRaceA.Text.Trim)
        Dim RaceText2 As String = PrepareHexData(Me.txtRaceB.Text.Trim)
        'MTP
        Dim mtpText1 As String = PrepareHexData(Me.txtMTPA.Text.Trim)
        Dim mtpText2 As String = PrepareHexData(Me.txtMTPB.Text.Trim)

        Dim TIMEa As String = PrepareHexData(Me.txtTODA.Text.Trim)
        Dim TIMEb As String = PrepareHexData(Me.txtTODB.Text.Trim)
        Dim TIMEc As String = PrepareHexData(Me.txtTODC.Text.Trim)
        Dim TIMEd As String = PrepareHexData(Me.txtTODD.Text.Trim)

        Try

            'BytesInPayload = ("0" & Hex(9))
            'BytesInPayload = BytesInPayload.Substring(BytesInPayload.Length - 2)

            'strPort = ("0" & Hex(26))
            'strPort = strPort.Substring(strPort.Length - 2)

            'myDataToSend = String.Format("{0}{1}{2}{3}{4}{5}00000000{6}{7}{8}{9}{10}0000",
            '                                     SOH, strPort, BoardControl, BoardDimming, PayloadType, BytesInPayload,
            '                                     RaceText1, RaceText2, mtpText1, mtpText2,
            '                                     EOT)

            ''convert it to char
            'MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)

            'Me.myComPort.Output(MessageToSend)

            'If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()

            BytesInPayload = ("0" & Hex(21))
            BytesInPayload = BytesInPayload.Substring(BytesInPayload.Length - 2)

            strPort = ("0" & Hex(50))
            strPort = strPort.Substring(strPort.Length - 2)

            myDataToSend = String.Format("{0}{1}{2}{3}{4}{5}00000000{6}{7}{8}{9}{10}{11}{12}{13}0000000000000000{14}0000",
                                                 SOH, strPort, BoardControl, BoardDimming, PayloadType, BytesInPayload,
                                                 TIMEa, TIMEb, TIMEc, TIMEd,
                                                 RaceText1, RaceText2, mtpText1, mtpText2,
                                                 EOT)


            'convert it to char
            MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)

            Me.myComPort.Output(MessageToSend)

            If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()

        Catch ex As Exception
            '
        End Try

    End Sub

    'Private Sub PrepareTeletimerToSend()
    '    Dim ConstantDataToSend As String = ""
    '    Dim myDataToSend As String = ""
    '    Dim MessageToSend As String = ""
    '    Dim strCheckSum As String = ""
    '    Dim strMsgCount As String = ""
    '    Dim strPort As String
    '    strPort = ("0" & Hex(12))
    '    strPort = strPort.Substring(strPort.Length - 2)

    '    'constant data to send
    '    ConstantDataToSend = String.Format("55AA{0}11", strPort)

    '    Dim FINISHa As String = IIf(Me.txtFinisha.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtFinisha.Text.Trim = "", "S", Me.txtFinisha.Text.Trim)))
    '    Dim FINISHb As String = IIf(Me.txtFinishb.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtFinishb.Text.Trim = "", "S", Me.txtFinishb.Text.Trim)))
    '    Dim FINISHc As String = IIf(Me.txtFinishc.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtFinishc.Text.Trim = "", "S", Me.txtFinishc.Text.Trim)))
    '    Dim FINISHd As String = IIf(Me.txtFinishd.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtFinishd.Text.Trim = "", "S", Me.txtFinishd.Text.Trim)))
    '    Dim FINISHe As String = IIf(Me.txtFinishe.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtFinishe.Text.Trim = "", "S", Me.txtFinishe.Text.Trim)))

    '    Dim MILEa As String = IIf(Me.txtMilea.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtMilea.Text.Trim = "", "S", Me.txtMilea.Text.Trim)))
    '    Dim MILEb As String = IIf(Me.txtMileb.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtMileb.Text.Trim = "", "S", Me.txtMileb.Text.Trim)))
    '    Dim MILEc As String = IIf(Me.txtMilec.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtMilec.Text.Trim = "", "S", Me.txtMilec.Text.Trim)))
    '    Dim MILEd As String = IIf(Me.txtMiled.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtMiled.Text.Trim = "", "S", Me.txtMiled.Text.Trim)))
    '    Dim MILEe As String = IIf(Me.txtMilee.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtMilee.Text.Trim = "", "S", Me.txtMilee.Text.Trim)))

    '    Dim S34a As String = IIf(Me.txt34a.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txt34a.Text.Trim = "", "S", Me.txt34a.Text.Trim)))
    '    Dim S34b As String = IIf(Me.txt34b.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txt34b.Text.Trim = "", "S", Me.txt34b.Text.Trim)))
    '    Dim S34c As String = IIf(Me.txt34c.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txt34c.Text.Trim = "", "S", Me.txt34c.Text.Trim)))
    '    Dim S34d As String = IIf(Me.txt34d.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txt34d.Text.Trim = "", "S", Me.txt34d.Text.Trim)))
    '    Dim S34e As String = IIf(Me.txt34e.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txt34e.Text.Trim = "", "S", Me.txt34e.Text.Trim)))

    '    strCheckSum = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}00", _
    '                                              FINISHa, FINISHb, FINISHc, FINISHd, FINISHe, _
    '                                              MILEa, MILEb, MILEc, MILEd, MILEe, _
    '                                              S34a, S34b, S34c, S34d, S34e)
    '    strMsgCount = Me.Getm_Ctr
    '    myDataToSend = String.Format("55AA{0}11{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}00{19}" _
    '                          , strPort, strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myOutsideIntensity _
    '                          , FINISHa, FINISHb, FINISHc, FINISHd, FINISHe _
    '                          , MILEa, MILEb, MILEc, MILEd, MILEe _
    '                          , S34a, S34b, S34c, S34d, S34e _
    '                          , Me.CalcChecksumToAsciiHexString(Me.myOutsideIntensity & strCheckSum))
    '    'convert it to char
    '    MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
    '    Me.myComPort.Output(MessageToSend)

    '    If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()

    '    strPort = ("0" & Hex(13))
    '    strPort = strPort.Substring(strPort.Length - 2)

    '    'constant data to send
    '    ConstantDataToSend = String.Format("55AA{0}11", strPort)

    '    Dim S12a As String = IIf(Me.txt12a.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txt12a.Text.Trim = "", "S", Me.txt12a.Text.Trim)))
    '    Dim S12b As String = IIf(Me.txt12b.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txt12b.Text.Trim = "", "S", Me.txt12b.Text.Trim)))
    '    Dim S12c As String = IIf(Me.txt12c.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txt12c.Text.Trim = "", "S", Me.txt12c.Text.Trim)))
    '    Dim S12d As String = IIf(Me.txt12d.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txt12d.Text.Trim = "", "S", Me.txt12d.Text.Trim)))
    '    Dim S12e As String = IIf(Me.txt12e.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txt12e.Text.Trim = "", "S", Me.txt12e.Text.Trim)))

    '    Dim S14a As String = IIf(Me.txt14a.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txt14a.Text.Trim = "", "S", Me.txt14a.Text.Trim)))
    '    Dim S14b As String = IIf(Me.txt14b.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txt14b.Text.Trim = "", "S", Me.txt14b.Text.Trim)))
    '    Dim S14c As String = IIf(Me.txt14c.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txt14c.Text.Trim = "", "S", Me.txt14c.Text.Trim)))
    '    Dim S14d As String = IIf(Me.txt14d.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txt14d.Text.Trim = "", "S", Me.txt14d.Text.Trim)))
    '    Dim S14e As String = IIf(Me.txt14e.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txt14e.Text.Trim = "", "S", Me.txt14e.Text.Trim)))

    '    strCheckSum = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}000000000000", _
    '                                              S12a, S12b, S12c, S12d, S12e, _
    '                                              S14a, S14b, S14c, S14d, S14e)
    '    strMsgCount = Me.Getm_Ctr
    '    myDataToSend = String.Format("55AA{0}11{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}000000000000{14}" _
    '                          , strPort, strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myOutsideIntensity _
    '                          , S12a, S12b, S12c, S12d, S12e _
    '                          , S14a, S14b, S14c, S14d, S14e _
    '                          , Me.CalcChecksumToAsciiHexString(Me.myOutsideIntensity & strCheckSum))
    '    'convert it to char
    '    MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
    '    Me.myComPort.Output(MessageToSend)

    '    If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()

    '    If (My.Settings.HasInsideBoard) Then
    '        '
    '    End If
    'End Sub

    Private Sub PrepareTeletimerToSend()
        Dim myDataToSend As String = ""
        Dim MessageToSend As String = ""

        'start of header
        Dim SOH As String = "01"
        Dim strPort As String
        'control, will send it always "on" from here
        Dim BoardControl As String = "01"
        'dimming, it should be done through the settings as well
        Dim BoardDimming As String = "0" & Hex(My.Settings.MiniBoardDimming)
        'pay load type, this value will be always 16 from here (hex(10))
        Dim PayloadType As String = "01"
        'end of transmission
        Dim EOT As String = "04"
        Dim BytesInPayload As String

        Dim FINISHa As String = PrepareHexData(Me.txtFinisha.Text.Trim)
        Dim FINISHb As String = PrepareHexData(Me.txtFinishb.Text.Trim)
        Dim FINISHc As String = PrepareHexData(Me.txtFinishc.Text.Trim)
        Dim FINISHd As String = PrepareHexData(Me.txtFinishd.Text.Trim)
        Dim FINISHe As String = PrepareHexData(Me.txtFinishe.Text.Trim)

        Dim MILEa As String = PrepareHexData(Me.txtMilea.Text.Trim)
        Dim MILEb As String = PrepareHexData(Me.txtMileb.Text.Trim)
        Dim MILEc As String = PrepareHexData(Me.txtMilec.Text.Trim)
        Dim MILEd As String = PrepareHexData(Me.txtMiled.Text.Trim)
        Dim MILEe As String = PrepareHexData(Me.txtMilee.Text.Trim)

        Dim S34a As String = PrepareHexData(Me.txt34a.Text.Trim)
        Dim S34b As String = PrepareHexData(Me.txt34b.Text.Trim)
        Dim S34c As String = PrepareHexData(Me.txt34c.Text.Trim)
        Dim S34d As String = PrepareHexData(Me.txt34d.Text.Trim)
        Dim S34e As String = PrepareHexData(Me.txt34e.Text.Trim)

        Dim S12a As String = PrepareHexData(Me.txt12a.Text.Trim)
        Dim S12b As String = PrepareHexData(Me.txt12b.Text.Trim)
        Dim S12c As String = PrepareHexData(Me.txt12c.Text.Trim)
        Dim S12d As String = PrepareHexData(Me.txt12d.Text.Trim)
        Dim S12e As String = PrepareHexData(Me.txt12e.Text.Trim)

        Dim S14a As String = PrepareHexData(Me.txt14a.Text.Trim)
        Dim S14b As String = PrepareHexData(Me.txt14b.Text.Trim)
        Dim S14c As String = PrepareHexData(Me.txt14c.Text.Trim)
        Dim S14d As String = PrepareHexData(Me.txt14d.Text.Trim)
        Dim S14e As String = PrepareHexData(Me.txt14e.Text.Trim)

        Try
            BytesInPayload = ("0" & Hex(21))
            BytesInPayload = BytesInPayload.Substring(BytesInPayload.Length - 2)

            strPort = ("0" & Hex(12))
            strPort = strPort.Substring(strPort.Length - 2)

            myDataToSend = String.Format("{0}{1}{2}{3}{4}{5}00000000{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}{19}{20}00{21}0000",
                                                 SOH, strPort, BoardControl, BoardDimming, PayloadType, BytesInPayload,
                                                 FINISHa, FINISHb, FINISHc, FINISHd, FINISHe,
                                                 MILEa, MILEb, MILEc, MILEd, MILEe,
                                                 S34a, S34b, S34c, S34d, S34e,
                                                 EOT)

            'convert it to char
            MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
            Me.myComPort.Output(MessageToSend)

            If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()

            strPort = ("0" & Hex(13))
            strPort = strPort.Substring(strPort.Length - 2)

            myDataToSend = String.Format("{0}{1}{2}{3}{4}{5}00000000{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}000000000000{16}0000",
                                                 SOH, strPort, BoardControl, BoardDimming, PayloadType, BytesInPayload,
                                                 S12a, S12b, S12c, S12d, S12e,
                                                 S14a, S14b, S14c, S14d, S14e,
                                                 EOT)
            'convert it to char
            MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
            Me.myComPort.Output(MessageToSend)

            If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()

        Catch ex As Exception
            '
        End Try
    End Sub

    'Private Sub PreparePoolsToSend()
    '    Dim ConstantDataToSend As String = ""
    '    Dim myDataToSend As String = ""
    '    Dim MessageToSend As String = ""
    '    Dim strCheckSum As String = ""
    '    Dim strPort As String

    '    'Pool Total
    '    strPort = ("0" & Hex(48))
    '    strPort = strPort.Substring(strPort.Length - 2)

    '    'constant data to send
    '    ConstantDataToSend = String.Format("55AA{0}11", strPort)
    '    Dim PoolTotA As String = IIf(Me.txtPoolTotA.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPoolTotA.Text.Trim = "", "S", Me.txtPoolTotA.Text.Trim)))
    '    Dim PoolTotB As String = IIf(Me.txtPoolTotB.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPoolTotB.Text.Trim = "", "S", Me.txtPoolTotB.Text.Trim)))
    '    Dim PoolTotC As String = IIf(Me.txtPoolTotC.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPoolTotC.Text.Trim = "", "S", Me.txtPoolTotC.Text.Trim)))
    '    Dim PoolTotD As String = IIf(Me.txtPoolTotD.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPoolTotD.Text.Trim = "", "S", Me.txtPoolTotD.Text.Trim)))
    '    Dim PoolTotE As String = IIf(Me.txtPoolTotE.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPoolTotE.Text.Trim = "", "S", Me.txtPoolTotE.Text.Trim)))
    '    Dim PoolTotF As String = IIf(Me.txtPoolTotF.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPoolTotF.Text.Trim = "", "S", Me.txtPoolTotF.Text.Trim)))
    '    '
    '    strCheckSum = String.Format("{0}{1}{2}{3}{4}{5}00000000000000000000", _
    '                                              PoolTotA, PoolTotB, PoolTotC, PoolTotD, _
    '                                              PoolTotE, PoolTotF)
    '    '
    '    Dim strMsgCount As String = Me.Getm_Ctr
    '    '
    '    myDataToSend = String.Format("55AA{0}11{1}{2}{3}{4}{5}{6}{7}{8}{9}00000000000000000000{10}" _
    '                          , strPort, strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myOutsideIntensity _
    '                          , PoolTotA, PoolTotB, PoolTotC, PoolTotD _
    '                          , PoolTotE, PoolTotF, Me.CalcChecksumToAsciiHexString(Me.myOutsideIntensity & strCheckSum))

    '    'convert it to char
    '    MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)

    '    Me.myComPort.Output(MessageToSend)

    '    If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()

    '    strPort = ("0" & Hex(36))
    '    strPort = strPort.Substring(strPort.Length - 2)

    '    'constant data to send
    '    ConstantDataToSend = String.Format("55AA{0}11", strPort)
    '    Dim txtPool1A As String = IIf(Me.txtPool1A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool1A.Text.Trim = "", "S", Me.txtPool1A.Text.Trim)))
    '    Dim txtPool1B As String = IIf(Me.txtPool1B.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool1B.Text.Trim = "", "S", Me.txtPool1B.Text.Trim)))
    '    Dim txtPool1C As String = IIf(Me.txtPool1C.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool1C.Text.Trim = "", "S", Me.txtPool1C.Text.Trim)))
    '    Dim txtPool1D As String = IIf(Me.txtPool1D.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool1D.Text.Trim = "", "S", Me.txtPool1D.Text.Trim)))
    '    Dim txtPool1E As String = IIf(Me.txtPool1E.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool1E.Text.Trim = "", "S", Me.txtPool1E.Text.Trim)))
    '    Dim txtPool2A As String = IIf(Me.txtPool2A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool2A.Text.Trim = "", "S", Me.txtPool2A.Text.Trim)))
    '    Dim txtPool2B As String = IIf(Me.txtPool2B.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool2B.Text.Trim = "", "S", Me.txtPool2B.Text.Trim)))
    '    Dim txtPool2C As String = IIf(Me.txtPool2C.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool2C.Text.Trim = "", "S", Me.txtPool2C.Text.Trim)))
    '    Dim txtPool2D As String = IIf(Me.txtPool2D.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool2D.Text.Trim = "", "S", Me.txtPool2D.Text.Trim)))
    '    Dim txtPool2E As String = IIf(Me.txtPool2E.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool2E.Text.Trim = "", "S", Me.txtPool2E.Text.Trim)))
    '    Dim txtPool3A As String = IIf(Me.txtPool3A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool3A.Text.Trim = "", "S", Me.txtPool3A.Text.Trim)))
    '    Dim txtPool3B As String = IIf(Me.txtPool3B.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool3B.Text.Trim = "", "S", Me.txtPool3B.Text.Trim)))
    '    Dim txtPool3C As String = IIf(Me.txtPool3C.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool3C.Text.Trim = "", "S", Me.txtPool3C.Text.Trim)))
    '    Dim txtPool3D As String = IIf(Me.txtPool3D.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool3D.Text.Trim = "", "S", Me.txtPool3D.Text.Trim)))
    '    Dim txtPool3E As String = IIf(Me.txtPool3E.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool3E.Text.Trim = "", "S", Me.txtPool3E.Text.Trim)))

    '    strCheckSum = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}00", _
    '                                      txtPool1A, txtPool1B, txtPool1C, txtPool1D, txtPool1E _
    '                                      , txtPool2A, txtPool2B, txtPool2C, txtPool2D, txtPool2E _
    '                                      , txtPool3A, txtPool3B, txtPool3C, txtPool3D, txtPool3E)
    '    strMsgCount = Me.Getm_Ctr
    '    myDataToSend = String.Format("55AA{0}11{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}00{19}" _
    '                          , strPort, strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myOutsideIntensity _
    '                          , txtPool1A, txtPool1B, txtPool1C, txtPool1D, txtPool1E _
    '                          , txtPool2A, txtPool2B, txtPool2C, txtPool2D, txtPool2E _
    '                          , txtPool3A, txtPool3B, txtPool3C, txtPool3D, txtPool3E _
    '                          , Me.CalcChecksumToAsciiHexString(Me.myOutsideIntensity & strCheckSum))
    '    'convert it to char
    '    MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
    '    Me.myComPort.Output(MessageToSend)
    '    If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()

    '    strPort = ("0" & Hex(38))
    '    strPort = strPort.Substring(strPort.Length - 2)

    '    'constant data to send
    '    ConstantDataToSend = String.Format("55AA{0}11", strPort)

    '    Dim txtPool4A As String = IIf(Me.txtPool4A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool4A.Text.Trim = "", "S", Me.txtPool4A.Text.Trim)))
    '    Dim txtPool4B As String = IIf(Me.txtPool4B.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool4B.Text.Trim = "", "S", Me.txtPool4B.Text.Trim)))
    '    Dim txtPool4C As String = IIf(Me.txtPool4C.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool4C.Text.Trim = "", "S", Me.txtPool4C.Text.Trim)))
    '    Dim txtPool4D As String = IIf(Me.txtPool4D.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool4D.Text.Trim = "", "S", Me.txtPool4D.Text.Trim)))
    '    Dim txtPool4E As String = IIf(Me.txtPool4E.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool4E.Text.Trim = "", "S", Me.txtPool4E.Text.Trim)))
    '    Dim txtPool5A As String = IIf(Me.txtPool5A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool5A.Text.Trim = "", "S", Me.txtPool5A.Text.Trim)))
    '    Dim txtPool5B As String = IIf(Me.txtPool5B.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool5B.Text.Trim = "", "S", Me.txtPool5B.Text.Trim)))
    '    Dim txtPool5C As String = IIf(Me.txtPool5C.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool5C.Text.Trim = "", "S", Me.txtPool5C.Text.Trim)))
    '    Dim txtPool5D As String = IIf(Me.txtPool5D.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool5D.Text.Trim = "", "S", Me.txtPool5D.Text.Trim)))
    '    Dim txtPool5E As String = IIf(Me.txtPool5E.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool5E.Text.Trim = "", "S", Me.txtPool5E.Text.Trim)))
    '    Dim txtPool6A As String = IIf(Me.txtPool6A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool6A.Text.Trim = "", "S", Me.txtPool6A.Text.Trim)))
    '    Dim txtPool6B As String = IIf(Me.txtPool6B.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool6B.Text.Trim = "", "S", Me.txtPool6B.Text.Trim)))
    '    Dim txtPool6C As String = IIf(Me.txtPool6C.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool6C.Text.Trim = "", "S", Me.txtPool6C.Text.Trim)))
    '    Dim txtPool6D As String = IIf(Me.txtPool6D.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool6D.Text.Trim = "", "S", Me.txtPool6D.Text.Trim)))
    '    Dim txtPool6E As String = IIf(Me.txtPool6E.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool6E.Text.Trim = "", "S", Me.txtPool6E.Text.Trim)))

    '    strCheckSum = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}00", _
    '                                      txtPool4A, txtPool4B, txtPool4C, txtPool4D, txtPool4E _
    '                                      , txtPool5A, txtPool5B, txtPool5C, txtPool5D, txtPool5E _
    '                                      , txtPool6A, txtPool6B, txtPool6C, txtPool6D, txtPool6E)
    '    strMsgCount = Me.Getm_Ctr
    '    myDataToSend = String.Format("55AA{0}11{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}00{19}" _
    '                          , strPort, strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myOutsideIntensity _
    '                          , txtPool4A, txtPool4B, txtPool4C, txtPool4D, txtPool4E _
    '                          , txtPool5A, txtPool5B, txtPool5C, txtPool5D, txtPool5E _
    '                          , txtPool6A, txtPool6B, txtPool6C, txtPool6D, txtPool6E _
    '                          , Me.CalcChecksumToAsciiHexString(Me.myOutsideIntensity & strCheckSum))
    '    MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
    '    Me.myComPort.Output(MessageToSend)
    '    If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()

    '    strPort = ("0" & Hex(40))
    '    strPort = strPort.Substring(strPort.Length - 2)

    '    'constant data to send
    '    ConstantDataToSend = String.Format("55AA{0}11", strPort)

    '    Dim txtPool7A As String = IIf(Me.txtPool7A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool7A.Text.Trim = "", "S", Me.txtPool7A.Text.Trim)))
    '    Dim txtPool7B As String = IIf(Me.txtPool7B.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool7B.Text.Trim = "", "S", Me.txtPool7B.Text.Trim)))
    '    Dim txtPool7C As String = IIf(Me.txtPool7C.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool7C.Text.Trim = "", "S", Me.txtPool7C.Text.Trim)))
    '    Dim txtPool7D As String = IIf(Me.txtPool7D.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool7D.Text.Trim = "", "S", Me.txtPool7D.Text.Trim)))
    '    Dim txtPool7E As String = IIf(Me.txtPool7E.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool7E.Text.Trim = "", "S", Me.txtPool7E.Text.Trim)))
    '    Dim txtPool8A As String = IIf(Me.txtPool8A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool8A.Text.Trim = "", "S", Me.txtPool8A.Text.Trim)))
    '    Dim txtPool8B As String = IIf(Me.txtPool8B.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool8B.Text.Trim = "", "S", Me.txtPool8B.Text.Trim)))
    '    Dim txtPool8C As String = IIf(Me.txtPool8C.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool8C.Text.Trim = "", "S", Me.txtPool8C.Text.Trim)))
    '    Dim txtPool8D As String = IIf(Me.txtPool8D.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool8D.Text.Trim = "", "S", Me.txtPool8D.Text.Trim)))
    '    Dim txtPool8E As String = IIf(Me.txtPool8E.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool8E.Text.Trim = "", "S", Me.txtPool8E.Text.Trim)))
    '    Dim txtPool9A As String = IIf(Me.txtPool9A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool9A.Text.Trim = "", "S", Me.txtPool9A.Text.Trim)))
    '    Dim txtPool9B As String = IIf(Me.txtPool9B.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool9B.Text.Trim = "", "S", Me.txtPool9B.Text.Trim)))
    '    Dim txtPool9C As String = IIf(Me.txtPool9C.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool9C.Text.Trim = "", "S", Me.txtPool9C.Text.Trim)))
    '    Dim txtPool9D As String = IIf(Me.txtPool9D.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool9D.Text.Trim = "", "S", Me.txtPool9D.Text.Trim)))
    '    Dim txtPool9E As String = IIf(Me.txtPool9E.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool9E.Text.Trim = "", "S", Me.txtPool9E.Text.Trim)))

    '    strCheckSum = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}00", _
    '                                      txtPool7A, txtPool7B, txtPool7C, txtPool7D, txtPool7E _
    '                                      , txtPool8A, txtPool8B, txtPool8C, txtPool8D, txtPool8E _
    '                                      , txtPool9A, txtPool9B, txtPool9C, txtPool9D, txtPool9E)
    '    strMsgCount = Me.Getm_Ctr
    '    myDataToSend = String.Format("55AA{0}11{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}00{19}" _
    '                          , strPort, strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myOutsideIntensity _
    '                          , txtPool7A, txtPool7B, txtPool7C, txtPool7D, txtPool7E _
    '                          , txtPool8A, txtPool8B, txtPool8C, txtPool8D, txtPool8E _
    '                          , txtPool9A, txtPool9B, txtPool9C, txtPool9D, txtPool9E _
    '                          , Me.CalcChecksumToAsciiHexString(Me.myOutsideIntensity & strCheckSum))
    '    MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
    '    Me.myComPort.Output(MessageToSend)
    '    If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()

    '    strPort = ("0" & Hex(42))
    '    strPort = strPort.Substring(strPort.Length - 2)

    '    'constant data to send
    '    ConstantDataToSend = String.Format("55AA{0}11", strPort)

    '    Dim txtPool10A As String = IIf(Me.txtPool10A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool10A.Text.Trim = "", "S", Me.txtPool10A.Text.Trim)))
    '    Dim txtPool10B As String = IIf(Me.txtPool10B.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool10B.Text.Trim = "", "S", Me.txtPool10B.Text.Trim)))
    '    Dim txtPool10C As String = IIf(Me.txtPool10C.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool10C.Text.Trim = "", "S", Me.txtPool10C.Text.Trim)))
    '    Dim txtPool10D As String = IIf(Me.txtPool10D.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool10D.Text.Trim = "", "S", Me.txtPool10D.Text.Trim)))
    '    Dim txtPool10E As String = IIf(Me.txtPool10E.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool10E.Text.Trim = "", "S", Me.txtPool10E.Text.Trim)))
    '    Dim txtPool11A As String = IIf(Me.txtPool11A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool11A.Text.Trim = "", "S", Me.txtPool11A.Text.Trim)))
    '    Dim txtPool11B As String = IIf(Me.txtPool11B.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool11B.Text.Trim = "", "S", Me.txtPool11B.Text.Trim)))
    '    Dim txtPool11C As String = IIf(Me.txtPool11C.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool11C.Text.Trim = "", "S", Me.txtPool11C.Text.Trim)))
    '    Dim txtPool11D As String = IIf(Me.txtPool11D.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool11D.Text.Trim = "", "S", Me.txtPool11D.Text.Trim)))
    '    Dim txtPool11E As String = IIf(Me.txtPool11E.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool11E.Text.Trim = "", "S", Me.txtPool11E.Text.Trim)))
    '    Dim txtPool12A As String = IIf(Me.txtPool12A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool12A.Text.Trim = "", "S", Me.txtPool12A.Text.Trim)))
    '    Dim txtPool12B As String = IIf(Me.txtPool12B.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool12B.Text.Trim = "", "S", Me.txtPool12B.Text.Trim)))
    '    Dim txtPool12C As String = IIf(Me.txtPool12C.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool12C.Text.Trim = "", "S", Me.txtPool12C.Text.Trim)))
    '    Dim txtPool12D As String = IIf(Me.txtPool12D.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool12D.Text.Trim = "", "S", Me.txtPool12D.Text.Trim)))
    '    Dim txtPool12E As String = IIf(Me.txtPool12E.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPool12E.Text.Trim = "", "S", Me.txtPool12E.Text.Trim)))

    '    strCheckSum = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}00", _
    '                                      txtPool10A, txtPool10B, txtPool10C, txtPool10D, txtPool10E _
    '                                      , txtPool11A, txtPool11B, txtPool11C, txtPool11D, txtPool11E _
    '                                      , txtPool12A, txtPool12B, txtPool12C, txtPool12D, txtPool12E)
    '    strMsgCount = Me.Getm_Ctr
    '    myDataToSend = String.Format("55AA{0}11{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}00{19}" _
    '                          , strPort, strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myOutsideIntensity _
    '                          , txtPool10A, txtPool10B, txtPool10C, txtPool10D, txtPool10E _
    '                          , txtPool11A, txtPool11B, txtPool11C, txtPool11D, txtPool11E _
    '                          , txtPool12A, txtPool12B, txtPool12C, txtPool12D, txtPool12E _
    '                          , Me.CalcChecksumToAsciiHexString(Me.myOutsideIntensity & strCheckSum))
    '    MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
    '    Me.myComPort.Output(MessageToSend)
    '    If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()

    '    'Inside Board
    '    If (My.Settings.HasInsideBoard) Then
    '        '
    '    End If

    '    PrepareWPSPoolHeaderToSend(False)

    'End Sub

    Private Sub PreparePoolsToSend()
        Dim myDataToSend As String = ""
        Dim MessageToSend As String = ""

        'start of header
        Dim SOH As String = "01"
        Dim strPort As String
        'control, will send it always "on" from here
        Dim BoardControl As String = "01"
        'dimming, it should be done through the settings as well
        Dim BoardDimming As String = "0" & Hex(My.Settings.MiniBoardDimming)
        'pay load type, this value will be always 16 from here (hex(10))
        Dim PayloadType As String = "01"
        'end of transmission
        Dim EOT As String = "04"
        Dim BytesInPayload As String

        BytesInPayload = ("0" & Hex(21))
        BytesInPayload = BytesInPayload.Substring(BytesInPayload.Length - 2)

        Dim PoolTotA As String = PrepareHexData(Me.txtPoolTotA.Text.Trim)
        Dim PoolTotB As String = PrepareHexData(Me.txtPoolTotB.Text.Trim)
        Dim PoolTotC As String = PrepareHexData(Me.txtPoolTotC.Text.Trim)
        Dim PoolTotD As String = PrepareHexData(Me.txtPoolTotD.Text.Trim)
        Dim PoolTotE As String = PrepareHexData(Me.txtPoolTotE.Text.Trim)
        Dim PoolTotF As String = PrepareHexData(Me.txtPoolTotF.Text.Trim)
        '
        Dim txtPool1A As String = PrepareHexData(Me.txtPool1A.Text.Trim)
        Dim txtPool1B As String = PrepareHexData(Me.txtPool1B.Text.Trim)
        Dim txtPool1C As String = PrepareHexData(Me.txtPool1C.Text.Trim)
        Dim txtPool1D As String = PrepareHexData(Me.txtPool1D.Text.Trim)
        Dim txtPool1E As String = PrepareHexData(Me.txtPool1E.Text.Trim)
        Dim txtPool2A As String = PrepareHexData(Me.txtPool2A.Text.Trim)
        Dim txtPool2B As String = PrepareHexData(Me.txtPool2B.Text.Trim)
        Dim txtPool2C As String = PrepareHexData(Me.txtPool2C.Text.Trim)
        Dim txtPool2D As String = PrepareHexData(Me.txtPool2D.Text.Trim)
        Dim txtPool2E As String = PrepareHexData(Me.txtPool2E.Text.Trim)
        Dim txtPool3A As String = PrepareHexData(Me.txtPool3A.Text.Trim)
        Dim txtPool3B As String = PrepareHexData(Me.txtPool3B.Text.Trim)
        Dim txtPool3C As String = PrepareHexData(Me.txtPool3C.Text.Trim)
        Dim txtPool3D As String = PrepareHexData(Me.txtPool3D.Text.Trim)
        Dim txtPool3E As String = PrepareHexData(Me.txtPool3E.Text.Trim)
        '
        Dim txtPool4A As String = PrepareHexData(Me.txtPool4A.Text.Trim)
        Dim txtPool4B As String = PrepareHexData(Me.txtPool4B.Text.Trim)
        Dim txtPool4C As String = PrepareHexData(Me.txtPool4C.Text.Trim)
        Dim txtPool4D As String = PrepareHexData(Me.txtPool4D.Text.Trim)
        Dim txtPool4E As String = PrepareHexData(Me.txtPool4E.Text.Trim)
        Dim txtPool5A As String = PrepareHexData(Me.txtPool5A.Text.Trim)
        Dim txtPool5B As String = PrepareHexData(Me.txtPool5B.Text.Trim)
        Dim txtPool5C As String = PrepareHexData(Me.txtPool5C.Text.Trim)
        Dim txtPool5D As String = PrepareHexData(Me.txtPool5D.Text.Trim)
        Dim txtPool5E As String = PrepareHexData(Me.txtPool5E.Text.Trim)
        Dim txtPool6A As String = PrepareHexData(Me.txtPool6A.Text.Trim)
        Dim txtPool6B As String = PrepareHexData(Me.txtPool6B.Text.Trim)
        Dim txtPool6C As String = PrepareHexData(Me.txtPool6C.Text.Trim)
        Dim txtPool6D As String = PrepareHexData(Me.txtPool6D.Text.Trim)
        Dim txtPool6E As String = PrepareHexData(Me.txtPool6E.Text.Trim)
        '
        Dim txtPool7A As String = PrepareHexData(Me.txtPool7A.Text.Trim)
        Dim txtPool7B As String = PrepareHexData(Me.txtPool7B.Text.Trim)
        Dim txtPool7C As String = PrepareHexData(Me.txtPool7C.Text.Trim)
        Dim txtPool7D As String = PrepareHexData(Me.txtPool7D.Text.Trim)
        Dim txtPool7E As String = PrepareHexData(Me.txtPool7E.Text.Trim)
        Dim txtPool8A As String = PrepareHexData(Me.txtPool8A.Text.Trim)
        Dim txtPool8B As String = PrepareHexData(Me.txtPool8B.Text.Trim)
        Dim txtPool8C As String = PrepareHexData(Me.txtPool8C.Text.Trim)
        Dim txtPool8D As String = PrepareHexData(Me.txtPool8D.Text.Trim)
        Dim txtPool8E As String = PrepareHexData(Me.txtPool8E.Text.Trim)
        Dim txtPool9A As String = PrepareHexData(Me.txtPool9A.Text.Trim)
        Dim txtPool9B As String = PrepareHexData(Me.txtPool9B.Text.Trim)
        Dim txtPool9C As String = PrepareHexData(Me.txtPool9C.Text.Trim)
        Dim txtPool9D As String = PrepareHexData(Me.txtPool9D.Text.Trim)
        Dim txtPool9E As String = PrepareHexData(Me.txtPool9E.Text.Trim)
        '
        Dim txtPool10A As String = PrepareHexData(Me.txtPool10A.Text.Trim)
        Dim txtPool10B As String = PrepareHexData(Me.txtPool10B.Text.Trim)
        Dim txtPool10C As String = PrepareHexData(Me.txtPool10C.Text.Trim)
        Dim txtPool10D As String = PrepareHexData(Me.txtPool10D.Text.Trim)
        Dim txtPool10E As String = PrepareHexData(Me.txtPool10E.Text.Trim)
        Dim txtPool11A As String = PrepareHexData(Me.txtPool11A.Text.Trim)
        Dim txtPool11B As String = PrepareHexData(Me.txtPool11B.Text.Trim)
        Dim txtPool11C As String = PrepareHexData(Me.txtPool11C.Text.Trim)
        Dim txtPool11D As String = PrepareHexData(Me.txtPool11D.Text.Trim)
        Dim txtPool11E As String = PrepareHexData(Me.txtPool11E.Text.Trim)
        Dim txtPool12A As String = PrepareHexData(Me.txtPool12A.Text.Trim)
        Dim txtPool12B As String = PrepareHexData(Me.txtPool12B.Text.Trim)
        Dim txtPool12C As String = PrepareHexData(Me.txtPool12C.Text.Trim)
        Dim txtPool12D As String = PrepareHexData(Me.txtPool12D.Text.Trim)
        Dim txtPool12E As String = PrepareHexData(Me.txtPool12E.Text.Trim)
        '
        Try
            'Pool Total
            strPort = ("0" & Hex(48))
            strPort = strPort.Substring(strPort.Length - 2)
            '

            myDataToSend = String.Format("{0}{1}{2}{3}{4}{5}00000000{6}{7}{8}{9}{10}{11}00000000000000000000{12}0000",
                                                     SOH, strPort, BoardControl, BoardDimming, PayloadType, BytesInPayload,
                                                     PoolTotA, PoolTotB, PoolTotC, PoolTotD, PoolTotE, PoolTotF,
                                                     EOT)

            'convert it to char
            MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)

            Me.myComPort.Output(MessageToSend)

            If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()

            strPort = ("0" & Hex(36))
            strPort = strPort.Substring(strPort.Length - 2)

            myDataToSend = String.Format("{0}{1}{2}{3}{4}{5}00000000{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}{19}{20}00{21}0000",
                                                     SOH, strPort, BoardControl, BoardDimming, PayloadType, BytesInPayload,
                                                     txtPool1A, txtPool1B, txtPool1C, txtPool1D, txtPool1E,
                                                     txtPool2A, txtPool2B, txtPool2C, txtPool2D, txtPool2E,
                                                     txtPool3A, txtPool3B, txtPool3C, txtPool3D, txtPool3E, EOT)

            'convert it to char
            MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
            Me.myComPort.Output(MessageToSend)
            If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()

            strPort = ("0" & Hex(38))
            strPort = strPort.Substring(strPort.Length - 2)

            myDataToSend = String.Format("{0}{1}{2}{3}{4}{5}00000000{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}{19}{20}00{21}0000",
                                                     SOH, strPort, BoardControl, BoardDimming, PayloadType, BytesInPayload,
                                                     txtPool4A, txtPool4B, txtPool4C, txtPool4D, txtPool4E,
                                                     txtPool5A, txtPool5B, txtPool5C, txtPool5D, txtPool5E,
                                                     txtPool6A, txtPool6B, txtPool6C, txtPool6D, txtPool6E, EOT)

            MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
            Me.myComPort.Output(MessageToSend)
            If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()

            strPort = ("0" & Hex(40))
            strPort = strPort.Substring(strPort.Length - 2)

            myDataToSend = String.Format("{0}{1}{2}{3}{4}{5}00000000{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}{19}{20}00{21}0000",
                                                     SOH, strPort, BoardControl, BoardDimming, PayloadType, BytesInPayload,
                                                     txtPool7A, txtPool7B, txtPool7C, txtPool7D, txtPool7E,
                                                     txtPool8A, txtPool8B, txtPool8C, txtPool8D, txtPool8E,
                                                     txtPool9A, txtPool9B, txtPool9C, txtPool9D, txtPool9E, EOT)

            MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
            Me.myComPort.Output(MessageToSend)
            If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()

            strPort = ("0" & Hex(42))
            strPort = strPort.Substring(strPort.Length - 2)

            myDataToSend = String.Format("{0}{1}{2}{3}{4}{5}00000000{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}{19}{20}00{21}0000",
                                                     SOH, strPort, BoardControl, BoardDimming, PayloadType, BytesInPayload,
                                                     txtPool10A, txtPool10B, txtPool10C, txtPool10D, txtPool10E,
                                                     txtPool11A, txtPool11B, txtPool11C, txtPool11D, txtPool11E,
                                                     txtPool12A, txtPool12B, txtPool12C, txtPool12D, txtPool12E, EOT)

            MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
            Me.myComPort.Output(MessageToSend)
            If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()

            PrepareWPSPoolHeaderToSend(False)
        Catch ex As Exception
            'nothing
        End Try
    End Sub

    'Private Sub PrepareWPSPoolHeaderToSend(ByVal blnFlagClear As Boolean)
    '    'start of header
    '    Dim SOH As String = "01"
    '    'address, lets get it fom the settings, tania will later decide how to approach this
    '    Dim BoardAddress As String = "00" 'My.Settings.MiniBoardAddress
    '    'control, will send it always "on" from here
    '    Dim BoardControl As String = "00"
    '    'dimming, it should be done through the settings as well
    '    Dim BoardDimming As String = "00" '"0" & Hex(My.Settings.MiniBoardDimming)
    '    'pay load type, this value will be always 16 from here (hex(10))
    '    Dim PayloadType As String = "17"
    '    Dim NumbOfString As String = "00"
    '    Dim NumbOfBytes As String = "1E" '30 Bytes (6 Columns * 5 Characters)
    '    Dim StartColumn As String = "00"
    '    'end of transmission
    '    Dim EOT As String = "04"
    '    '
    '    Dim strTemp As String = " "
    '    If (blnFlagClear = False) Then
    '        strTemp = txtPoolType.Text
    '    End If
    '    strTemp = strTemp.ToString.PadRight(5, " ")

    '    Dim displaySegments() As Emac.DisplayMatrixUtils.DisplaySegment
    '    Dim strCol As String
    '    Dim strColTemp As String
    '    Dim col As String = ""
    '    Dim col1 As String = ""
    '    Dim col2 As String = ""
    '    Dim col3 As String = ""
    '    Dim col4 As String = ""
    '    Dim col5 As String = ""

    '    For intCtr As Integer = 1 To Len(strTemp)
    '        Try
    '            col = ""
    '            strColTemp = ""
    '            strCol = Mid(strTemp, intCtr, 1)
    '            displaySegments = Emac.DisplayMatrixUtils.DisplaySegmentDictionary.GetDisplaySegments(strCol)
    '            For colIdx As Integer = 0 To 5
    '                strColTemp = displaySegments(colIdx).HexRowValue
    '                col = col & strColTemp.PadLeft(2, "0")
    '            Next colIdx
    '        Catch ex As Exception
    '            col = "000000000000"
    '        End Try
    '        If intCtr = 1 Then
    '            col1 = col
    '        ElseIf intCtr = 2 Then
    '            col2 = col
    '        ElseIf intCtr = 3 Then
    '            col3 = col
    '        ElseIf intCtr = 4 Then
    '            col4 = col
    '        ElseIf intCtr = 5 Then
    '            col5 = col
    '        End If
    '    Next intCtr
    '    '
    '    'build the string to send
    '    Dim MessageToSend As String = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}", _
    '                                             SOH, BoardAddress, BoardControl, BoardDimming, PayloadType, _
    '                                             NumbOfString, StartColumn, NumbOfBytes, _
    '                                             col1, col2, col3, col4, col5, _
    '                                             EOT, Me.CalculateCRC(1), Me.CalculateCRC(2))

    '    'Me.myMiniComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend("010080001700001E008080FE8080008080FE8080008080FE8080008080FE8080008080FE8080040000"))

    '    Me.myMiniComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend(MessageToSend))
    '    If Not Me.myMiniComPort.ErrorMessage = "" Then Me.ShowError(True)

    '    ''
    'End Sub

    Private Sub PrepareWPSPoolHeaderToSend(ByVal blnFlagClear As Boolean)
        Dim myDataToSend As String = ""
        Dim MessageToSend As String = ""
        'start of header
        Dim SOH As String = "01"
        'address, lets get it fom the settings, tania will later decide how to approach this
        Dim BoardAddress As String = "02"
        'control, will send it always "on" from here
        Dim BoardControl As String = "00"
        'dimming, it should be done through the settings as well
        Dim BoardDimming As String = "00"
        'pay load type, this value will be always 24 from here (hex(18))
        Dim PayloadType As String = "18"
        'end of transmission
        Dim EOT As String = "04"
        '
        Dim StringNumber As String = "00" 'string number for WPS

        Dim strTemp As String = " "
        If (blnFlagClear = False) Then
            strTemp = txtPoolType.Text
        End If
        strTemp = strTemp.ToString.PadRight(5, " ")

        Dim strDatatosend = PrepareDataThreeColorSign(strTemp, Color.Yellow, 24)
        strDatatosend = strDatatosend.Replace(" ", "")
        'strDatatosend = "3132333435"
        'Me.myComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend(strDatatosend))

        Try
            If (strDatatosend <> "-1") Then
                '(between 1 and 240 bytes of Data
                Dim intNumbDataBytes As Integer
                intNumbDataBytes = (strDatatosend.Length / 2)
                Dim NumbDataBytes As String
                NumbDataBytes = ("0" & Hex(intNumbDataBytes))
                NumbDataBytes = NumbDataBytes.Substring(NumbDataBytes.Length - 2)

                'NumbDataBytes = "05"

                myDataToSend = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}0000",
                                                     SOH, BoardAddress, BoardControl, BoardDimming, PayloadType,
                                                     StringNumber, NumbDataBytes, strDatatosend,
                                                     EOT)

                'convert it to char
                MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
                '
                Me.myComPort.Output(MessageToSend)

                If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()
            End If

        Catch ex As Exception

        End Try
        ''
    End Sub

    Private Function PrepareDataThreeColorSign(txtText As String,
                                               color As Color,
                                               fontsize As Integer) As String
        Dim MessageToSend As String = "-1"

        Try
            'string places from 1-10 (constant) position 1 a5 is start of file so we dont need it for now
            MessageToSend = "68 32 01 74 11 00 00 00 00"
            'next byte or position 11 is the sum of constant 22 + incoming text length to hex
            MessageToSend += $" {Hex(22 + txtText.Length)}"
            'positions 12-13 never change, position 14 may change but it doesnt for our parameters 
            '(font size 24 color red/yellow speed 100 stay 255 loop 1)
            MessageToSend += " 00 77 01"
            'position 15 if the decimal sum of all positions from 16-34 to hex
            'so we have to find the values, some are constants, some not.
            'from here we will need another string
            Dim strtocalculate As String

            'position 16 changes with the loop, for now loop will be always 1
            'so this is the string from 16-24 in hex 01 00 00 00 00 00 64 00 14

            'position 25 is the combination of font size and color, here are the constants
            'red 8 = 10    green 8 = 20    yellow 8 = 30
            'red 12 = 11   green 12 = 21   yellow 12 = 31
            'red 16 = 12   green 16 = 22   yellow 16 = 32
            'red 24 = 13   green 24 = 23   yellow 24 = 33
            'red 32 = 14   green 32 = 24   yellow 32 = 34
            'so we are now working with 16 and 24 (passed as parameter) and the color is passed by the calling method
            Dim fontColorValue As String = String.Empty
            Select Case color
                Case Color.Red
                    fontColorValue = If(fontsize = 16, "12", "13")
                Case Color.Yellow
                    fontColorValue = If(fontsize = 16, "32", "33")
            End Select
            strtocalculate = $"00 00 00 00 00 00 64 00 14 {fontColorValue}"
            'strtocalculate = $"01 00 00 00 00 00 64 00 14 {fontColorValue}"
            'strtocalculate = $"01 00 00 00 00 00 64 00 14 {If(color = Color.Red, 13, 33)}"

            'positions 26-30 never changes, 31 represents the speed but since this value will be 
            'always 100, there's no need to do anything at this point. so from 36-31 the string is 00 00 00 00 00 00
            'position 32 displays the stay value, this is also a constant value for now, it has a fixed value of 255 or ff in hex
            'add position 32 to the string above 00 00 00 00 00 00 ff
            'position 33 never changes and has a 00 value, so our final string from 26-33 is 00 00 00 00 00 00 ff 00
            strtocalculate += " 00 00 00 00 00 00 ff 00"

            'position 34 represents the incoming text length in hex\
            strtocalculate += $" {Hex(txtText.Length).PadLeft(2, "0")}"

            'now extract the decimal values, add them and convert the sum to hex
            Dim a As String() = strtocalculate.Split
            Dim total As Integer = 0
            a.ToList.ForEach(Sub(pos)
                                 'this converts our hex string to int
                                 total += Convert.ToInt32(pos, 16)
                             End Sub)
            'now we have the total, convert it to hex and add it to position 15
            'notice that if len hex(total) is greater than 2, we just take the last 2 digits since 
            'position 14 already has the first
            'we can also add the rest here
            MessageToSend += $" {Hex(total).PadLeft(3).Substring(1)} {strtocalculate}"
            'the rest is all the text to hex
            Dim ch As Char() = txtText.ToArray()

            ch.ToList.ForEach(Sub(c)
                                  MessageToSend += $" {Hex(Asc(c))}"
                              End Sub)
            'now that we have the entire string, replace the following symbols..
            'A5 represents beginning of file, replace it with 05
            MessageToSend = MessageToSend.ToUpper.Replace(" A5", " 05")

            'AE Represents end of file, replace it with 0E
            MessageToSend = MessageToSend.ToUpper.Replace(" AE", " 0E")

            'AA represents escape character, replace it with 0A
            MessageToSend = MessageToSend.ToUpper.Replace(" AA", " 0A")

            'now calculate check sum, remember is a two byte (4 digits) check sum where the last 
            'Byte goes first on the string to be sent
            Dim chkSum As String = CalculateCheckSum(MessageToSend.Replace(" ", ""))

            'insert check sum, function CalculateCheckSum returns it exactly as we need it
            'insert beginning and EOF to return a complete string
            Return $"A5 {MessageToSend} {chkSum} AE"
            '
        Catch ex As Exception
            Return "-1"
        End Try

    End Function

    Private Function PrepareDataThreeColorSignStatus(txtText As String,
                                               color As Color,
                                               fontsize As Integer) As String
        Dim MessageToSend As String = "-1"

        Try
            'string places from 1-10 (constant) position 1 a5 is start of file so we dont need it for now
            MessageToSend = "68 32 01 74 11 00 00 00 00"
            'next byte or position 11 is the sum of constant 22 + incoming text length to hex
            MessageToSend += $" {Hex(22 + txtText.Length)}"
            'positions 12-13 never change, position 14 may change but it doesnt for our parameters 
            '(font size 24 color red/yellow speed 100 stay 255 loop 1)
            MessageToSend += " 00 77 01"
            'position 15 if the decimal sum of all positions from 16-34 to hex
            'so we have to find the values, some are constants, some not.
            'from here we will need another string
            Dim strtocalculate As String

            'position 16 changes with the loop, for now loop will be always 1
            'so this is the string from 16-24 in hex 01 00 00 00 00 00 64 00 14

            'position 25 is the combination of font size and color, here are the constants
            'red 8 = 10    green 8 = 20    yellow 8 = 30
            'red 12 = 11   green 12 = 21   yellow 12 = 31
            'red 16 = 12   green 16 = 22   yellow 16 = 32
            'red 24 = 13   green 24 = 23   yellow 24 = 33
            'red 32 = 14   green 32 = 24   yellow 32 = 34
            'so we are now working with 16 and 24 (passed as parameter) and the color is passed by the calling method
            Dim fontColorValue As String = String.Empty
            Select Case color
                Case Color.Red
                    fontColorValue = If(fontsize = 16, "12", "13")
                Case Color.Yellow
                    fontColorValue = If(fontsize = 16, "32", "33")
            End Select
            strtocalculate = $"00 00 00 00 00 00 8C 00 14 {fontColorValue}"
            'strtocalculate = $"01 00 00 00 00 00 8C 00 14 {fontColorValue}"
            'strtocalculate = $"01 00 00 00 00 00 64 00 14 {If(color = Color.Red, 13, 33)}"

            'positions 26-30 never changes, 31 represents the speed but since this value will be 
            'always 140, there's no need to do anything at this point. so from 36-31 the string is 00 00 00 00 00 00
            'position 32 displays the stay value, this is also a constant value for now, it has a fixed value of 255 or ff in hex
            'add position 32 to the string above 00 00 00 00 00 00 ff
            'position 33 never changes and has a 00 value, so our final string from 26-33 is 00 00 00 00 00 00 ff 00
            strtocalculate += " 00 00 00 00 00 00 ff 00"

            'position 34 represents the incoming text length in hex\
            strtocalculate += $" {Hex(txtText.Length).PadLeft(2, "0")}"

            'now extract the decimal values, add them and convert the sum to hex
            Dim a As String() = strtocalculate.Split
            Dim total As Integer = 0
            a.ToList.ForEach(Sub(pos)
                                 'this converts our hex string to int
                                 total += Convert.ToInt32(pos, 16)
                             End Sub)
            'now we have the total, convert it to hex and add it to position 15
            'notice that if len hex(total) is greater than 2, we just take the last 2 digits since 
            'position 14 already has the first
            'we can also add the rest here
            MessageToSend += $" {Hex(total).PadLeft(3).Substring(1)} {strtocalculate}"
            'the rest is all the text to hex
            Dim ch As Char() = txtText.ToArray()

            ch.ToList.ForEach(Sub(c)
                                  MessageToSend += $" {Hex(Asc(c))}"
                              End Sub)
            'now that we have the entire string, replace the following symbols..
            'A5 represents beginning of file, replace it with 05
            MessageToSend = MessageToSend.ToUpper.Replace(" A5", " 05")

            'AE Represents end of file, replace it with 0E
            MessageToSend = MessageToSend.ToUpper.Replace(" AE", " 0E")

            'AA represents escape character, replace it with 0A
            MessageToSend = MessageToSend.ToUpper.Replace(" AA", " 0A")

            'now calculate check sum, remember is a two byte (4 digits) check sum where the last 
            'Byte goes first on the string to be sent
            Dim chkSum As String = CalculateCheckSum(MessageToSend.Replace(" ", ""))

            'insert check sum, function CalculateCheckSum returns it exactly as we need it
            'insert beginning and EOF to return a complete string
            Return $"A5 {MessageToSend} {chkSum} AE"
            '
        Catch ex As Exception
            Return "-1"
        End Try

    End Function

    Private Function PrepareDataThreeColorSignExotics(txtText As String,
                                               color As Color,
                                               fontsize As Integer) As String
        Dim MessageToSend As String = "-1"

        Try
            'string places from 1-10 (constant) position 1 a5 is start of file so we dont need it for now
            MessageToSend = "68 32 01 74 11 00 00 00 00"
            'next byte or position 11 is the sum of constant 22 + incoming text length to hex
            MessageToSend += $" {Hex(22 + txtText.Length)}"
            'positions 12-13 never change, position 14 may change but it doesnt for our parameters 
            '(font size 24 color red/yellow speed 100 stay 255 loop 1)
            MessageToSend += " 00 77 01"
            'position 15 if the decimal sum of all positions from 16-34 to hex
            'so we have to find the values, some are constants, some not.
            'from here we will need another string
            Dim strtocalculate As String

            'position 16 changes with the loop, for now loop will be always 1
            'so this is the string from 16-24 in hex 01 00 00 00 00 00 64 00 14

            'position 25 is the combination of font size and color, here are the constants
            'red 8 = 10    green 8 = 20    yellow 8 = 30
            'red 12 = 11   green 12 = 21   yellow 12 = 31
            'red 16 = 12   green 16 = 22   yellow 16 = 32
            'red 24 = 13   green 24 = 23   yellow 24 = 33
            'red 32 = 14   green 32 = 24   yellow 32 = 34
            'so we are now working with 16 and 24 (passed as parameter) and the color is passed by the calling method
            Dim fontColorValue As String = String.Empty
            Select Case color
                Case Color.Red
                    fontColorValue = If(fontsize = 16, "12", "13")
                Case Color.Yellow
                    fontColorValue = If(fontsize = 16, "32", "33")
            End Select
            strtocalculate = $"00 00 00 00 00 01 E0 00 14 {fontColorValue}"
            'strtocalculate = $"01 00 00 00 00 00 8C 00 14 {fontColorValue}"
            'strtocalculate = $"01 00 00 00 00 00 64 00 14 {If(color = Color.Red, 13, 33)}"

            'positions 26-30 never changes, 31 represents the speed but since this value will be 
            'always 140, there's no need to do anything at this point. so from 36-31 the string is 00 00 00 00 00 00
            'position 32 displays the stay value, this is also a constant value for now, it has a fixed value of 255 or ff in hex
            'add position 32 to the string above 00 00 00 00 00 00 ff
            'position 33 never changes and has a 00 value, so our final string from 26-33 is 00 00 00 00 00 00 ff 00
            strtocalculate += " 00 00 00 00 00 00 ff 00"

            'position 34 represents the incoming text length in hex\
            strtocalculate += $" {Hex(txtText.Length).PadLeft(2, "0")}"

            'now extract the decimal values, add them and convert the sum to hex
            Dim a As String() = strtocalculate.Split
            Dim total As Integer = 0
            a.ToList.ForEach(Sub(pos)
                                 'this converts our hex string to int
                                 total += Convert.ToInt32(pos, 16)
                             End Sub)
            'now we have the total, convert it to hex and add it to position 15
            'notice that if len hex(total) is greater than 2, we just take the last 2 digits since 
            'position 14 already has the first
            'we can also add the rest here
            MessageToSend += $" {Hex(total).PadLeft(3).Substring(1)} {strtocalculate}"
            'the rest is all the text to hex
            Dim ch As Char() = txtText.ToArray()

            ch.ToList.ForEach(Sub(c)
                                  MessageToSend += $" {Hex(Asc(c))}"
                              End Sub)
            'now that we have the entire string, replace the following symbols..
            'A5 represents beginning of file, replace it with 05
            MessageToSend = MessageToSend.ToUpper.Replace(" A5", " 05")

            'AE Represents end of file, replace it with 0E
            MessageToSend = MessageToSend.ToUpper.Replace(" AE", " 0E")

            'AA represents escape character, replace it with 0A
            MessageToSend = MessageToSend.ToUpper.Replace(" AA", " 0A")

            'now calculate check sum, remember is a two byte (4 digits) check sum where the last 
            'Byte goes first on the string to be sent
            Dim chkSum As String = CalculateCheckSum(MessageToSend.Replace(" ", ""))

            'insert check sum, function CalculateCheckSum returns it exactly as we need it
            'insert beginning and EOF to return a complete string
            Return $"A5 {MessageToSend} {chkSum} AE"
            '
        Catch ex As Exception
            Return "-1"
        End Try

    End Function

    'Private Sub PrepareMiniWPSDataToSend(Optional ByVal sendWinDataOnly As Boolean = False)
    '    'set constants

    '    'start of header
    '    Dim SOH As String = "01"
    '    'address, lets get it fom the settings, tania will later decide how to approach this
    '    Dim BoardAddress As String = "00" 'My.Settings.MiniBoardAddress
    '    'control, will send it always "on" from here
    '    Dim BoardControl As String = "01"
    '    'dimming, it should be done through the settings as well
    '    Dim BoardDimming As String = "0" & Hex(My.Settings.MiniBoardDimming)
    '    'pay load type
    '    Dim PayloadType As String = "13"
    '    'end of transmission
    '    Dim EOT As String = "04"

    '    'first win
    '    Dim WIN1a As String = If(Me.txtWin1A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Me.strMiniBoardCollectionOff("K" & Me.txtWin1A.Text.Trim))
    '    Dim WIN1b As String = If(Me.txtWin1B.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Me.strMiniBoardCollectionOff("K" & Me.txtWin1B.Text.Trim))
    '    Dim WIN1c As String = If(Me.txtWin1C.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Me.strMiniBoardCollectionOff("K" & Me.txtWin1C.Text.Trim))
    '    Dim WIN1d As String = If(Me.txtWin1D.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtWin1D.Text.Trim)) + &H20))
    '    Dim WIN1e As String = If(Me.txtWin1E.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Me.strMiniBoardCollectionOff("K" & Me.txtWin1E.Text.Trim))
    '    Dim WIN1f As String = If(Me.txtWin1F.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Me.strMiniBoardCollectionOff("K" & Me.txtWin1F.Text.Trim))
    '    'second win
    '    Dim WIN2a As String = If(Me.txtWin2A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Me.strMiniBoardCollectionOff("K" & Me.txtWin2A.Text.Trim))
    '    Dim WIN2b As String = If(Me.txtWin2B.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Me.strMiniBoardCollectionOff("K" & Me.txtWin2B.Text.Trim))
    '    Dim WIN2c As String = If(Me.txtWin2C.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Me.strMiniBoardCollectionOff("K" & Me.txtWin2C.Text.Trim))
    '    Dim WIN2d As String = If(Me.txtWin2D.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtWin2D.Text.Trim)) + &H20))
    '    Dim WIN2e As String = If(Me.txtWin2E.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Me.strMiniBoardCollectionOff("K" & Me.txtWin2E.Text.Trim))
    '    Dim WIN2f As String = If(Me.txtWin2F.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Me.strMiniBoardCollectionOff("K" & Me.txtWin2F.Text.Trim))
    '    'general
    '    Dim general As String = If(Me.chkOfficial.Checked, "&H08", "&H00")
    '    general = Hex(general + If(Me.chkPhoto.Checked, &H4, &H0))
    '    general = Hex(general + If(Me.chkObj.Checked, &H2, &H0))
    '    general = Hex(general + If(Me.chkDeadHeat.Checked, &H1, &H0))
    '    If general.Length < 2 Then general = general.PadLeft(2, "0")

    '    'build the string to send
    '    Dim MessageToSend As String = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}{19}{20}",
    '                                             SOH, BoardAddress, BoardControl, BoardDimming, PayloadType,
    '                                             WIN1a, WIN1b, WIN1c, WIN1d, WIN1e, WIN1f,
    '                                             WIN2a, WIN2b, WIN2c, WIN2d, WIN2e, WIN2f,
    '                                             general, EOT, Me.CalculateCRC(1), Me.CalculateCRC(2))
    '    '
    '    Me.myMiniComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend(MessageToSend))
    '    If Not Me.myMiniComPort.ErrorMessage = "" Then Me.ShowError(True)
    '    '
    '    'If sendWinDataOnly Then Return
    '    '
    '    'PLACE
    '    PayloadType = "14"
    '    '
    '    'first place
    '    Dim PLACE1a As String = If(Me.txtPlace1A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtPlace1A.Text.Trim))
    '    Dim PLACE1b As String = If(Me.txtPlace1B.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtPlace1B.Text.Trim))
    '    Dim PLACE1c As String = If(Me.txtPlace1C.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtPlace1C.Text.Trim)) + &H20))
    '    Dim PLACE1d As String = If(Me.txtPlace1D.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtPlace1D.Text.Trim))
    '    Dim PLACE1e As String = If(Me.txtPlace1E.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtPlace1E.Text.Trim))
    '    'second place
    '    Dim PLACE2a As String = If(Me.txtPlace2A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtPlace2A.Text.Trim))
    '    Dim PLACE2b As String = If(Me.txtPlace2B.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtPlace2B.Text.Trim))
    '    Dim PLACE2c As String = If(Me.txtPlace2C.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtPlace2C.Text.Trim)) + &H20))
    '    Dim PLACE2d As String = If(Me.txtPlace2D.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtPlace2D.Text.Trim))
    '    Dim PLACE2e As String = If(Me.txtPlace2E.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtPlace2E.Text.Trim))
    '    '3rd place
    '    Dim PLACE3a As String = If(Me.txtPlace3A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtPlace3A.Text.Trim))
    '    Dim PLACE3b As String = If(Me.txtPlace3B.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtPlace3B.Text.Trim))
    '    Dim PLACE3c As String = If(Me.txtPlace3C.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtPlace3C.Text.Trim)) + &H20))
    '    Dim PLACE3d As String = If(Me.txtPlace3D.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtPlace3D.Text.Trim))
    '    Dim PLACE3e As String = If(Me.txtPlace3E.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtPlace3E.Text.Trim))
    '    '
    '    'build the string to send
    '    MessageToSend = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}{19}{20}{21}{22}",
    '                                             SOH, BoardAddress, BoardControl, BoardDimming, PayloadType,
    '                                             PLACE1a, PLACE1b, PLACE1c, PLACE1d, PLACE1e,
    '                                             PLACE2a, PLACE2b, PLACE2c, PLACE2d, PLACE2e,
    '                                             PLACE3a, PLACE3b, PLACE3c, PLACE3d, PLACE3e,
    '                                             EOT, Me.CalculateCRC(1), Me.CalculateCRC(2))
    '    Me.myMiniComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend(MessageToSend))
    '    If Not Me.myMiniComPort.ErrorMessage = "" Then Me.ShowError(True)

    '    'SHOW
    '    PayloadType = "15"
    '    '
    '    Dim SHOW1a As String = If(Me.txtShow1A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtShow1A.Text.Trim))
    '    Dim SHOW1b As String = If(Me.txtShow1B.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtShow1B.Text.Trim))
    '    Dim SHOW1c As String = If(Me.txtShow1C.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtShow1C.Text.Trim)) + &H20))
    '    Dim SHOW1d As String = If(Me.txtShow1D.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtShow1D.Text.Trim))
    '    Dim SHOW1e As String = If(Me.txtShow1E.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtShow1E.Text.Trim))
    '    'second show
    '    Dim SHOW2a As String = If(Me.txtShow2A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtShow2A.Text.Trim))
    '    Dim SHOW2b As String = If(Me.txtShow2B.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtShow2B.Text.Trim))
    '    Dim SHOW2c As String = If(Me.txtShow2C.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtShow2C.Text.Trim)) + &H20))
    '    Dim SHOW2d As String = If(Me.txtShow2D.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtShow2D.Text.Trim))
    '    Dim SHOW2e As String = If(Me.txtShow2E.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtShow2E.Text.Trim))
    '    '3rd SHOW
    '    Dim SHOW3a As String = If(Me.txtShow3A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtShow3A.Text.Trim))
    '    Dim SHOW3b As String = If(Me.txtShow3B.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtShow3B.Text.Trim))
    '    Dim SHOW3c As String = If(Me.txtShow3C.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtShow3C.Text.Trim)) + &H20))
    '    Dim SHOW3d As String = If(Me.txtShow3D.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtShow3D.Text.Trim))
    '    Dim SHOW3e As String = If(Me.txtShow3E.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtShow3E.Text.Trim))
    '    '4th show
    '    Dim SHOW4a As String = If(Me.txtShow4A.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtShow4A.Text.Trim))
    '    Dim SHOW4b As String = If(Me.txtShow4B.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtShow4B.Text.Trim))
    '    Dim SHOW4c As String = If(Me.txtShow4C.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtShow4C.Text.Trim)) + &H20))
    '    Dim SHOW4d As String = If(Me.txtShow4D.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtShow4D.Text.Trim))
    '    Dim SHOW4e As String = If(Me.txtShow4E.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"), Me.strMiniBoardCollectionOff("K" & Me.txtShow4E.Text.Trim))
    '    'build the string to send
    '    MessageToSend = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}{19}{20}{21}{22}{23}{24}{25}{26}{27}",
    '                                             SOH, BoardAddress, BoardControl, BoardDimming, PayloadType,
    '                                             SHOW1a, SHOW1b, SHOW1c, SHOW1d, SHOW1e,
    '                                             SHOW2a, SHOW2b, SHOW2c, SHOW2d, SHOW2e,
    '                                             SHOW3a, SHOW3b, SHOW3c, SHOW3d, SHOW3e,
    '                                             SHOW4a, SHOW4b, SHOW4c, SHOW4d, SHOW4e,
    '                                             EOT, Me.CalculateCRC(1), Me.CalculateCRC(2))

    '    Me.myMiniComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend(MessageToSend))
    '    If Not Me.myMiniComPort.ErrorMessage = "" Then Me.ShowError(True)
    '    ''
    'End Sub

    'Private Sub PrepareMiniTimingDataToSend()
    '    'set constants

    '    'start of header
    '    Dim SOH As String = "01"
    '    'address, lets get it fom the settings, tania will later decide how to approach this
    '    Dim BoardAddress As String = "00" 'My.Settings.MiniBoardAddress
    '    'control, will send it always "on" from here
    '    Dim BoardControl As String = "01"
    '    'dimming, it should be done through the settings as well
    '    Dim BoardDimming As String = "0" & Hex(My.Settings.MiniBoardDimming)
    '    'pay load type for timing 22 (hex(16))
    '    Dim PayloadType As String = "16"
    '    'end of transmission
    '    Dim EOT As String = "04"
    '    '
    '    Dim FINISHa As String = If(Me.txtFinisha.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                            Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtFinisha.Text.Trim)) + &H20))
    '    '
    '    Dim FINISHb As String = If(Me.txtFinishb.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Me.strMiniBoardCollectionOff("K" & Me.txtFinishb.Text.Trim))
    '    '
    '    Dim FINISHc As String = If(Me.txtFinishc.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtFinishc.Text.Trim)) + &H20))
    '    '
    '    Dim FINISHd As String = If(Me.txtFinishd.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Me.strMiniBoardCollectionOff("K" & Me.txtFinishd.Text.Trim))
    '    '
    '    Dim FINISHe As String = If(Me.txtFinishe.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Me.strMiniBoardCollectionOff("K" & Me.txtFinishe.Text.Trim))
    '    '
    '    Dim MILEa As String = If(Me.txtMilea.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                              Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtMilea.Text.Trim)) + &H20))
    '    '
    '    Dim MILEb As String = If(Me.txtMileb.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Me.strMiniBoardCollectionOff("K" & Me.txtMileb.Text.Trim))
    '    '
    '    Dim MILEc As String = If(Me.txtMilec.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                              Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txtMilec.Text.Trim)) + &H20))
    '    '
    '    Dim MILEd As String = If(Me.txtMiled.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Me.strMiniBoardCollectionOff("K" & Me.txtMiled.Text.Trim))
    '    '
    '    Dim MILEe As String = If(Me.txtMilee.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Me.strMiniBoardCollectionOff("K" & Me.txtMilee.Text.Trim))
    '    '
    '    Dim S34a As String = If(Me.txt34a.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                              Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txt34a.Text.Trim)) + &H20))
    '    '
    '    Dim S34b As String = If(Me.txt34b.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Me.strMiniBoardCollectionOff("K" & Me.txt34b.Text.Trim))
    '    '
    '    Dim S34c As String = If(Me.txt34c.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                              Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txt34c.Text.Trim)) + &H20))
    '    '
    '    Dim S34d As String = If(Me.txt34d.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Me.strMiniBoardCollectionOff("K" & Me.txt34d.Text.Trim))
    '    '
    '    Dim S34e As String = If(Me.txt34e.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Me.strMiniBoardCollectionOff("K" & Me.txt34e.Text.Trim))
    '    '
    '    Dim S12a As String = If(Me.txt12a.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txt12a.Text.Trim)) + &H20))
    '    '
    '    Dim S12b As String = If(Me.txt12b.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Me.strMiniBoardCollectionOff("K" & Me.txt12b.Text.Trim))
    '    '
    '    Dim S12c As String = If(Me.txt12c.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txt12c.Text.Trim)) + &H20))
    '    '
    '    Dim S12d As String = If(Me.txt12d.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Me.strMiniBoardCollectionOff("K" & Me.txt12d.Text.Trim))
    '    '
    '    Dim S12e As String = If(Me.txt12e.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Me.strMiniBoardCollectionOff("K" & Me.txt12e.Text.Trim))
    '    '
    '    Dim S14a As String = If(Me.txt14a.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txt14a.Text.Trim)) + &H20))
    '    '
    '    Dim S14b As String = If(Me.txt14b.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Me.strMiniBoardCollectionOff("K" & Me.txt14b.Text.Trim))
    '    '
    '    Dim S14c As String = If(Me.txt14c.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Hex(("&H" & Me.strMiniBoardCollectionOff("K" & Me.txt14c.Text.Trim)) + &H20))
    '    '
    '    Dim S14d As String = If(Me.txt14d.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Me.strMiniBoardCollectionOff("K" & Me.txt14d.Text.Trim))
    '    '
    '    Dim S14e As String = If(Me.txt14e.Text.Trim = "", Me.strMiniBoardCollectionOff("KS"),
    '                             Me.strMiniBoardCollectionOff("K" & Me.txt14e.Text.Trim))
    '    'build the string to send
    '    Dim MessageToSend As String = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}{19}{20}{21}{22}{23}{24}{25}{26}{27}{28}{29}{30}{31}{32}",
    '                                             SOH, BoardAddress, BoardControl, BoardDimming, PayloadType,
    '                                             FINISHa, FINISHb, FINISHc, FINISHd, FINISHe,
    '                                             MILEa, MILEb, MILEc, MILEd, MILEe,
    '                                             S34a, S34b, S34c, S34d, S34e,
    '                                             S12a, S12b, S12c, S12d, S12e,
    '                                             S14a, S14b, S14c, S14d, S14e,
    '                                             EOT, Me.CalculateCRC(1), Me.CalculateCRC(2))
    '    '
    '    Me.myMiniComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend(MessageToSend))
    '    If Not Me.myMiniComPort.ErrorMessage = "" Then Me.ShowError(True)
    '    ''
    'End Sub

    Private Sub PrepareExactaDataToSend()
        'clear trifecta before we do anything
        Me.ClearTrifecta()
        '
        Dim ConstantDataToSend As String = ""
        Dim myDataToSend As String = ""
        Dim MessageToSend As String = ""
        Dim strCheckSum As String = ""
        Dim strPort As String
        '
        'EXACTA
        strPort = ("0" & Hex(19))
        strPort = strPort.Substring(strPort.Length - 2)
        '
        ConstantDataToSend = String.Format("55AA{0}11", strPort)
        'exacta winners
        Dim Perfecta1a As String = IIf(Me.txtPerfecta1A.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPerfecta1A.Text.Trim = "", "S", Me.txtPerfecta1A.Text.Trim)))
        Dim Perfecta1b As String = IIf(Me.txtPerfecta1B.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtPerfecta1B.Text.Trim = "", "S", Me.txtPerfecta1B.Text.Trim)))
        Dim Perfecta2a As String = IIf(Me.txtPerfecta2A.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtPerfecta2A.Text.Trim = "", "S", Me.txtPerfecta2A.Text.Trim)))
        Dim Perfecta2b As String = IIf(Me.txtPerfecta2B.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtPerfecta2B.Text.Trim = "", "S", Me.txtPerfecta2B.Text.Trim)))
        'exacta amount
        Dim PerfectaAmta As String = IIf(Me.txtPerfectaAmountA.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPerfectaAmountA.Text.Trim = "", "S", Me.txtPerfectaAmountA.Text.Trim)))
        Dim PerfectaAmtb As String = IIf(Me.txtPerfectaAmountB.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPerfectaAmountB.Text.Trim = "", "S", Me.txtPerfectaAmountB.Text.Trim)))
        Dim PerfectaAmtc As String = IIf(Me.txtPerfectaAmountC.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPerfectaAmountC.Text.Trim = "", "S", Me.txtPerfectaAmountC.Text.Trim)))
        Dim PerfectaAmtd As String = IIf(Me.txtPerfectaAmountD.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPerfectaAmountD.Text.Trim = "", "S", Me.txtPerfectaAmountD.Text.Trim)))
        Dim PerfectaAmte As String = IIf(Me.txtPerfectaAmountE.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPerfectaAmountE.Text.Trim = "", "S", Me.txtPerfectaAmountE.Text.Trim)))
        Dim PerfectaAmtf As String = IIf(Me.txtPerfectaAmountF.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPerfectaAmountF.Text.Trim = "", "S", Me.txtPerfectaAmountF.Text.Trim)))
        Dim PerfectaAmtg As String = IIf(Me.txtPerfectaAmountG.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtPerfectaAmountG.Text.Trim = "", "S", Me.txtPerfectaAmountG.Text.Trim)))
        '
        strCheckSum = String.Format("{0}{1}{2}{3}0000{4}{5}{6}{7}{8}{9}{10}000000",
                                                  Perfecta1a, Perfecta1b, Perfecta2a, Perfecta2b,
                                                  PerfectaAmta, PerfectaAmtb, PerfectaAmtc, PerfectaAmtd, PerfectaAmte, PerfectaAmtf, PerfectaAmtg)
        '
        Dim strMsgCount As String = Me.Getm_Ctr
        '
        myDataToSend = String.Format("55AA{0}11{1}{2}{3}{4}{5}{6}{7}0000{8}{9}{10}{11}{12}{13}{14}000000{15}" _
                              , strPort, strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myOutsideIntensity _
                              , Perfecta1a, Perfecta1b, Perfecta2a, Perfecta2b _
                              , PerfectaAmta, PerfectaAmtb, PerfectaAmtc, PerfectaAmtd, PerfectaAmte, PerfectaAmtf, PerfectaAmtg _
                              , Me.CalcChecksumToAsciiHexString(Me.myOutsideIntensity & strCheckSum))
        'convert it to char
        MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
        'turn on exacta
        'Me.rbExacta.Checked = True
        Me.myComPort.Output(MessageToSend)
        If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()
        '
        'Me.timerOfficial.Enabled = True

    End Sub

    Private Sub PrepareTrifectaDataToSend()
        '
        Me.ClearExacta()
        '
        Dim ConstantDataToSend As String = ""
        Dim myDataToSend As String = ""
        Dim MessageToSend As String = ""
        Dim strCheckSum As String = ""
        Dim strPort As String
        '
        strPort = ("0" & Hex(19))
        strPort = strPort.Substring(strPort.Length - 2)
        '
        ConstantDataToSend = String.Format("55AA{0}11", strPort)
        'trifecta winners
        Dim Trifecta1a As String = IIf(Me.txtTrifecta1A.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtTrifecta1A.Text.Trim = "", "S", Me.txtTrifecta1A.Text.Trim)))
        Dim Trifecta1b As String = IIf(Me.txtTrifecta1B.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtTrifecta1B.Text.Trim = "", "S", Me.txtTrifecta1B.Text.Trim)))
        Dim Trifecta2a As String = IIf(Me.txtTrifecta2A.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtTrifecta2A.Text.Trim = "", "S", Me.txtTrifecta2A.Text.Trim)))
        Dim Trifecta2b As String = IIf(Me.txtTrifecta2B.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtTrifecta2B.Text.Trim = "", "S", Me.txtTrifecta2B.Text.Trim)))
        Dim Trifecta3a As String = IIf(Me.txtTrifecta3A.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtTrifecta3A.Text.Trim = "", "S", Me.txtTrifecta3A.Text.Trim)))
        Dim Trifecta3b As String = IIf(Me.txtTrifecta3B.Text.Trim = "", Me.strCollectionOff("KS"), Me.strCollectionOff("K" & IIf(Me.txtTrifecta3B.Text.Trim = "", "S", Me.txtTrifecta3B.Text.Trim)))
        'trifecta amount
        Dim TrifectaAmta As String = IIf(Me.txtTrifectaAmountA.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtTrifectaAmountA.Text.Trim = "", "S", Me.txtTrifectaAmountA.Text.Trim)))
        Dim TrifectaAmtb As String = IIf(Me.txtTrifectaAmountB.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtTrifectaAmountB.Text.Trim = "", "S", Me.txtTrifectaAmountB.Text.Trim)))
        Dim TrifectaAmtc As String = IIf(Me.txtTrifectaAmountC.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtTrifectaAmountC.Text.Trim = "", "S", Me.txtTrifectaAmountC.Text.Trim)))
        Dim TrifectaAmtd As String = IIf(Me.txtTrifectaAmountD.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtTrifectaAmountD.Text.Trim = "", "S", Me.txtTrifectaAmountD.Text.Trim)))
        Dim TrifectaAmte As String = IIf(Me.txtTrifectaAmountE.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtTrifectaAmountE.Text.Trim = "", "S", Me.txtTrifectaAmountE.Text.Trim)))
        Dim TrifectaAmtf As String = IIf(Me.txtTrifectaAmountF.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtTrifectaAmountF.Text.Trim = "", "S", Me.txtTrifectaAmountF.Text.Trim)))
        Dim TrifectaAmtg As String = IIf(Me.txtTrifectaAmountG.Text.Trim = "", Me.strCollectionOn("KS"), Me.strCollectionOn("K" & IIf(Me.txtTrifectaAmountG.Text.Trim = "", "S", Me.txtTrifectaAmountG.Text.Trim)))
        '
        strCheckSum = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}000000",
                                                  Trifecta1a, Trifecta1b, Trifecta2a, Trifecta2b, Trifecta3a, Trifecta3b,
                                                  TrifectaAmta, TrifectaAmtb, TrifectaAmtc, TrifectaAmtd, TrifectaAmte, TrifectaAmtf, TrifectaAmtg)
        '
        Dim strMsgCount As String = Me.Getm_Ctr
        '
        myDataToSend = String.Format("55AA{0}11{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}000000{17}" _
                              , strPort, strMsgCount, Me.CalcChecksumToAsciiHexString(ConstantDataToSend & strMsgCount), myOutsideIntensity _
                              , Trifecta1a, Trifecta1b, Trifecta2a, Trifecta2b, Trifecta3a, Trifecta3b _
                              , TrifectaAmta, TrifectaAmtb, TrifectaAmtc, TrifectaAmtd, TrifectaAmte, TrifectaAmtf, TrifectaAmtg _
                              , Me.CalcChecksumToAsciiHexString(Me.myOutsideIntensity & strCheckSum))
        'convert it to char
        MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
        'turn on trifecta
        'Me.rbTrifecta.Checked = True
        Me.myComPort.Output(MessageToSend)
        If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()

        'Me.timerOfficial.Enabled = True

    End Sub

    Private Function CalculateCRC(ByVal place As Integer) As String
        Return "00"
        'If place = 1 Then Return "76" Else Return "40"
    End Function

    Private Function Getm_Ctr() As String
        m_Ctr = ("&H" & m_Ctr) + &H10
        m_Ctr = IIf(m_Ctr > 256, m_Ctr - 256, m_Ctr)
        m_Ctr = (Hex(m_Ctr).PadLeft(2, "0"))
        Return m_Ctr
    End Function

    Private Function CalcChecksumToAsciiHexString(UserString As String) As String
        Dim total As Integer
        'Add the values of each Ascii Hex pair:
        For i As Integer = 0 To (UserString.Length - 1) Step 2
            total += Val("&H" & UserString.Substring(i, 2))
        Next
        '
        total = Decimal.Negate(total)
        '
        Dim strChecksumAsAsciiHex As String = Hex(total)
        'need to pull the right most two characters from strChecksumAsAsciiHex at this point
        Return strChecksumAsAsciiHex.Substring(strChecksumAsAsciiHex.Length - 2, 2)
        '
    End Function

    Private Function CalculateCheckSum(UserString As String) As String
        Try
            Dim total As Integer
            'Add the values of each Ascii Hex pair:
            For i As Integer = 0 To (UserString.Length - 1) Step 2
                If UserString.Substring(i, 2) = "20" Then
                    Dim s As String = String.Empty
                End If
                total += Val("&H" & UserString.Substring(i, 2))
            Next
            '
            Dim chkSum As String = Hex(total).PadLeft(4, "0")
            'return last 2 digits + first 2 separated by a space
            Return $"{chkSum.Substring(2, 2)} {chkSum.Substring(0, 2)}"
            '
        Catch ex As Exception
            MessageBox.Show($"Error Calculating Check Sum.
{ex.Message}")
            Return "00 00"
        End Try

    End Function

    Private Function GetString(ByVal objValue As Object) As String
        Try
            If IsDBNull(objValue) Then Return ""
            Return CType(objValue, String)
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Private Sub KillComSvr()
        For Each p As Process In Process.GetProcesses
            If p.ProcessName.ToLower = "rsiport" Then
                p.Kill()
                Exit For
            End If
        Next
    End Sub

    Private Sub KillDataSvr()
        For Each p As Process In Process.GetProcesses
            If p.ProcessName.ToLower = "rsiservice" Then
                p.Kill()
                Exit For
            End If
        Next
    End Sub

    Private Sub UncheckRadioButtons()
        '
        For Each c As Control In Me.gbTrack.Controls
            If TypeOf c Is System.Windows.Forms.RadioButton Then CType(c, System.Windows.Forms.RadioButton).Checked = False
        Next
        ''
    End Sub

#End Region

#Region " Events Raised by CommSvr "

    Private Delegate Sub ResetRaceBindingSource()

    Private Sub myCommSvr_DisplayNewRace()
        If Me.InvokeRequired Then
            Me.Invoke(New ResetRaceBindingSource(AddressOf myCommSvr_DisplayNewRace))
        Else
            Me.bsRace.DataSource = Me.RaceDisplayDataset.RACE
            Me.bsRace.ResetBindings(False)
            Me.bsRace.EndEdit()
        End If
        ''
    End Sub

    Private Delegate Sub ResetMTPBindingSource()

    Private Sub myCommSvr_DisplayMTP()
        If Me.InvokeRequired Then
            Me.Invoke(New ResetMTPBindingSource(AddressOf myCommSvr_DisplayMTP))
        Else
            Me.bsMTP.DataSource = Me.RaceDisplayDataset.MTP
            Me.bsMTP.ResetBindings(False)
            Me.bsMTP.EndEdit()
        End If
        ''
    End Sub

    Private Delegate Sub ResetTrackConditionBindingSource()

    Private Sub myCommSvr_DisplayTrackCondition()
        If Me.InvokeRequired Then
            Me.Invoke(New ResetTrackConditionBindingSource(AddressOf myCommSvr_DisplayTrackCondition))
        Else
            '
        End If
        If (rbTote.Checked) Then
            myTrackCondition = Me.myCommSvr.p_strToteTrackCondition 'Me.myCommSvr.oCommServerNet.CurrentTrackCondition
            txtTrackCondition.Text = myTrackCondition
            Me.PrepareTrackCondition(False)
        End If
    End Sub

    Private Delegate Sub ResetRaceStatusBindingSource()

    Private Sub myCommSvr_DisplayRaceStatus()
        If Me.InvokeRequired Then
            Me.Invoke(New ResetRaceStatusBindingSource(AddressOf myCommSvr_DisplayRaceStatus))
        Else
            '
        End If
        Dim strStatus As String
        strStatus = Me.myCommSvr.oCommServerNet.CurrentStatus()
        Dim strRunnersFlashingStatus As String
        strRunnersFlashingStatus = Me.myCommSvr.oCommServerNet.CurrentRunnersFlashingStatus()

        Try
            p_blnUpdateManually = False
            For Each ctl As Control In Me.gbOrder.Controls
                If TypeOf ctl Is Windows.Forms.CheckBox Then
                    CType(ctl, Windows.Forms.CheckBox).Checked = False
                End If
            Next
            p_blnUpdateManually = True
        Catch ex As Exception
            p_blnUpdateManually = True
        End Try

        If (strRunnersFlashingStatus.Trim <> "") Then
            Try
                p_blnUpdateManually = False

                Dim vntInfoArray() As String = Split(strRunnersFlashingStatus, ",")
                Dim strTemp As String = ""
                For i As Integer = 0 To vntInfoArray.Length - 1
                    strTemp = vntInfoArray(i).Trim.ToUpper()
                    If (strTemp = (Me.txtRunning1.Text & Me.txtRunning1A.Text).Trim.ToUpper()) Then
                        Me.chkResults1.Checked = True
                    ElseIf (strTemp = (Me.txtRunning2.Text & Me.txtRunning2A.Text).Trim.ToUpper()) Then
                        Me.chkResults2.Checked = True
                    ElseIf (strTemp = (Me.txtRunning3.Text & Me.txtRunning3A.Text).Trim.ToUpper()) Then
                        Me.chkResults3.Checked = True
                    ElseIf (strTemp = (Me.txtRunning4.Text & Me.txtRunning4A.Text).Trim.ToUpper()) Then
                        Me.chkResults4.Checked = True
                    End If
                Next

                p_blnUpdateManually = True
            Catch ex As Exception
                p_blnUpdateManually = True
            End Try
        End If

        Me.PrepareRunningOrderDataToSend()

        'If (InStr(strStatus, "X") > 0) Then
        '    Try
        '        m_FlagChangeEvent = False
        '        ClearStatus("")
        '        m_FlagChangeEvent = True
        '    Catch ex As Exception
        '        m_FlagChangeEvent = True
        '    End Try
        '    txtOfficial.Text = ""
        '    Me.PrepareStatusDataToSend()
        '    If (txtRaceStatus.Text.Trim = "HOLD ALL TICKETS") Then
        '        txtRaceStatus.Text = ""
        '        Me.PrepareExoticsToSend(1, False, True)
        '    End If
        'm_FlagChangeEvent = True
        If (InStr(strStatus, "F") > 0) Then
            Me.chkOfficial.Checked = True
        ElseIf (InStr(strStatus, "D") > 0) Then
            Me.chkDeadHeat.Checked = True
        ElseIf (InStr(strStatus, "P") > 0) Then
            Me.chkPhoto.Checked = True
        ElseIf ((InStr(strStatus, "I") > 0) And (InStr(strStatus, "O") > 0)) Then
            Me.chkInqObj.Checked = True
        ElseIf (InStr(strStatus, "I") > 0) Then
            Me.chkInq.Checked = True
        ElseIf (InStr(strStatus, "O") > 0) Then
            Me.chkObj.Checked = True
        Else
            Try
                m_FlagChangeEvent = False
                ClearStatus("")
                m_FlagChangeEvent = True
            Catch ex As Exception
                m_FlagChangeEvent = True
            End Try
            txtOfficial.Text = ""
            Me.PrepareStatusDataToSend()
            If (txtRaceStatus.Text.Trim = "HOLD ALL TICKETS") Then
                txtRaceStatus.Text = ""
                Me.PrepareExoticsToSend(1, False, True)
            End If
        End If
        'Me.PrepareStatusDataToSend()
    End Sub

    Private Delegate Sub ResetPTBindingSource()

    Private Sub myCommSvr_DisplayPostTime()
        If Me.InvokeRequired Then
            Me.Invoke(New ResetPTBindingSource(AddressOf myCommSvr_DisplayPostTime))
        Else
            Me.bsPostTime.DataSource = Me.RaceDisplayDataset.POSTTIME
            Me.bsPostTime.ResetBindings(False)
            Me.bsPostTime.EndEdit()
        End If
        ''
    End Sub

    Private Delegate Sub ResetTODBindingSource()

    Private Sub myCommSvr_DisplayTOD()
        If Me.InvokeRequired Then
            Me.Invoke(New ResetTODBindingSource(AddressOf myCommSvr_DisplayTOD))
        Else
            Me.bsTOD.DataSource = Me.RaceDisplayDataset.TOD
            Me.bsTOD.ResetBindings(False)
            Me.bsTOD.EndEdit()
        End If
        Me.PrepareTimeOfDayToSend()
    End Sub

    Private Delegate Sub ResetTeletimerBindingSource()

    Private Sub myCommSvr_DisplayTeletimer()
        If Me.InvokeRequired Then
            Me.Invoke(New ResetTeletimerBindingSource(AddressOf myCommSvr_DisplayTeletimer))
        Else
            Me.bsTimingFinish.DataSource = Me.RaceDisplayDataset.TIMINGFINISH
            Me.bsTimingFinish.ResetBindings(False)
            Me.bsTimingFinish.EndEdit()
            '
            Me.bsTimingMile.DataSource = Me.RaceDisplayDataset.TIMINGMILE
            Me.bsTimingMile.ResetBindings(False)
            Me.bsTimingMile.EndEdit()
            '
            Me.bsTiming34.DataSource = Me.RaceDisplayDataset.TIMING34
            Me.bsTiming34.ResetBindings(False)
            Me.bsTiming34.EndEdit()
            '
            Me.bsTiming12.DataSource = Me.RaceDisplayDataset.TIMING12
            Me.bsTiming12.ResetBindings(False)
            Me.bsTiming12.EndEdit()
            '
            Me.bsTiming14.DataSource = Me.RaceDisplayDataset.TIMING14
            Me.bsTiming14.ResetBindings(False)
            Me.bsTiming14.EndEdit()
            '
        End If
        Me.PrepareTeletimerToSend()
    End Sub

    Private Delegate Sub ResetOddsBindingSource()

    Private Sub myCommSvr_DisplayNewOdds()
        If Me.InvokeRequired Then
            Me.Invoke(New ResetOddsBindingSource(AddressOf myCommSvr_DisplayNewOdds))
        Else
            Dim myArrayOdds() As Integer = Me.myCommSvr.ArrayOdds
            'Does a slash show up in teh GUI or not.  The array holds 0 for no slash, 1 for a slash.
            Me.lbl1.Visible = myArrayOdds(0) = 1
            Me.lbl2.Visible = myArrayOdds(1) = 1
            Me.lbl3.Visible = myArrayOdds(2) = 1
            Me.lbl4.Visible = myArrayOdds(3) = 1
            Me.lbl5.Visible = myArrayOdds(4) = 1
            Me.lbl6.Visible = myArrayOdds(5) = 1
            Me.lbl7.Visible = myArrayOdds(6) = 1
            Me.lbl8.Visible = myArrayOdds(7) = 1
            Me.lbl9.Visible = myArrayOdds(8) = 1
            Me.lbl10.Visible = myArrayOdds(9) = 1
            Me.lbl11.Visible = myArrayOdds(10) = 1
            Me.lbl12.Visible = myArrayOdds(11) = 1
            Me.lbl13.Visible = myArrayOdds(12) = 1
            Me.lbl14.Visible = myArrayOdds(13) = 1
            Me.lbl15.Visible = myArrayOdds(14) = 1
            Me.lbl16.Visible = myArrayOdds(15) = 1
            '
            Me.bsODDS.DataSource = Me.RaceDisplayDataset.ODDS
            Me.bsODDS.ResetBindings(False)
            Me.bsODDS.EndEdit()
            '
            Me.PrepareOddsDataToSend(Me.myCommSvr.ArrayOdds)
            'moment of truth
            For i As Integer = 1 To Me.myColMessages.Count
                Me.myComPort.Output(Me.myColMessages(i))
            Next
            If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError(False)
            '
        End If
        ''
    End Sub

    Private Sub DisplayWPSPools()
        'Here can I seach the ini for the information, depending on a count to cycle?  Otherwise, should
        'I use what is below?  Was there a reason for breaking these up?
        Me.bsWPSPools.DataSource = Me.RaceDisplayDataset.EXOTICS 'can't change this anymore.  must use "EXOTICS" until I can change it.
        Me.bsWPSPools.ResetBindings(False)
        Me.bsWPSPools.EndEdit()
        Me.PreparePoolsToSend()
    End Sub

    Private Delegate Sub ResetRunningOderAmtoteBindingSource()

    Private Sub myCommSvr_DisplayRunningOrderAmtote()
        If Me.InvokeRequired Then
            Me.Invoke(New ResetRunningOderAmtoteBindingSource(AddressOf myCommSvr_DisplayRunningOrderAmtote))
        Else
            p_blnUpdateManually = False

            Me.bsRunningOrder.DataSource = Me.RaceDisplayDataset.RUNNINGORDER
            Me.bsRunningOrder.ResetBindings(False)
            Me.bsRunningOrder.EndEdit()
            'statuses
            Me.bsStatus.DataSource = Me.RaceDisplayDataset.STATUS
            Me.bsStatus.ResetBindings(False)
            Me.bsStatus.EndEdit()
            '
            Me.chkDeadHeat.Refresh()
            Me.chkPhoto.Refresh()
            Me.chkInq.Refresh()
            Me.chkObj.Refresh()
            Me.chkInqObj.Refresh()
            Me.chkOfficial.Refresh()
            '
            Me.chkResults1.Refresh()
            Me.chkResults2.Refresh()
            Me.chkResults3.Refresh()
            Me.chkResults4.Refresh()
            '
            p_blnUpdateManually = True
            '
            Me.chkResults1.Checked = Me.RaceDisplayDataset.RUNNINGORDER.Rows(0)!WINNERDH1
            Me.chkResults2.Checked = Me.RaceDisplayDataset.RUNNINGORDER.Rows(0)!WINNERDH2
            Me.chkResults3.Checked = Me.RaceDisplayDataset.RUNNINGORDER.Rows(0)!WINNERDH3
            Me.chkResults4.Checked = Me.RaceDisplayDataset.RUNNINGORDER.Rows(0)!WINNERDH4

            Me.chkDeadHeat.Checked = Me.RaceDisplayDataset.STATUS.Rows(0)!DEADHEAT
            Me.chkPhoto.Checked = Me.RaceDisplayDataset.STATUS.Rows(0)!PHOTO
            Me.chkInq.Checked = Me.RaceDisplayDataset.STATUS.Rows(0)!INQUIRY
            Me.chkObj.Checked = Me.RaceDisplayDataset.STATUS.Rows(0)!OBJECTION
            Me.chkInqObj.Checked = Me.RaceDisplayDataset.STATUS.Rows(0)!INQOBJ
            Me.chkOfficial.Checked = Me.RaceDisplayDataset.STATUS.Rows(0)!OFFICIAL

            PrepareRunningOrderDataToSend()

        End If
        ''
    End Sub

    Private Delegate Sub ResetRunningOderBindingSource()

    Private Sub myCommSvr_DisplayRunningOrder()
        If Me.InvokeRequired Then
            Me.Invoke(New ResetRunningOderBindingSource(AddressOf myCommSvr_DisplayRunningOrder))
        Else

            Me.bsRunningOrder.DataSource = Me.RaceDisplayDataset.RUNNINGORDER
            Me.bsRunningOrder.ResetBindings(False)
            Me.bsRunningOrder.EndEdit()
            '
            PrepareRunningOrderDataToSend()
        End If
        ''
    End Sub


    Private Delegate Sub ResetNewResultsBindingSource()

    Private Sub myCommSvr_DisplayNewResults()
        If Me.InvokeRequired Then
            Me.Invoke(New ResetNewResultsBindingSource(AddressOf myCommSvr_DisplayNewResults))
        Else

            Me.bsRunningOrder.DataSource = Me.RaceDisplayDataset.RUNNINGORDER
            Me.bsRunningOrder.ResetBindings(False)
            Me.bsRunningOrder.EndEdit()
            '
            Dim blnFlag As Boolean = myCommSvr.m_blnClearExotics
            PrepareRunningOrderDataToSend()
            If blnFlag = False Then
                Me.timerOfficial.Enabled = True
            End If

        End If
        ''
    End Sub

    Private Delegate Sub JMDisplayBindingSource(ByVal jmData As JMData)
    Private Sub DisplayJudgesMessage(ByVal jmData As JMData)
        Debug.WriteLine(jmData.ToString())
        If Me.InvokeRequired Then
            Me.Invoke(New JMDisplayBindingSource(AddressOf DisplayJudgesMessage), jmData)
        Else
            '
        End If

        Dim strRunnersFlashingStatus As String = ""

        Try
            p_blnUpdateManually = False
            For Each ctl As Control In Me.gbOrder.Controls
                If TypeOf ctl Is Windows.Forms.CheckBox Then
                    CType(ctl, Windows.Forms.CheckBox).Checked = False
                End If
            Next
            p_blnUpdateManually = True
        Catch ex As Exception
            p_blnUpdateManually = True
        End Try

        Try
            If (jmData.FlashingRunnersPresent) Then
                Try
                    Dim flashingRunners() As String = jmData.GetStatusRuners()

                    If Not flashingRunners Is Nothing Then
                        Dim strTemp As String = ""
                        For i As Integer = 0 To flashingRunners.Length - 1
                            If Not flashingRunners(i) Is Nothing Then
                                strTemp = flashingRunners(i).Trim.ToUpper()
                                If (strTemp = (Me.txtRunning1.Text & Me.txtRunning1A.Text).Trim.ToUpper()) Then
                                    Me.chkResults1.Checked = True
                                ElseIf (strTemp = (Me.txtRunning2.Text & Me.txtRunning2A.Text).Trim.ToUpper()) Then
                                    Me.chkResults2.Checked = True
                                ElseIf (strTemp = (Me.txtRunning3.Text & Me.txtRunning3A.Text).Trim.ToUpper()) Then
                                    Me.chkResults3.Checked = True
                                ElseIf (strTemp = (Me.txtRunning4.Text & Me.txtRunning4A.Text).Trim.ToUpper()) Then
                                    Me.chkResults4.Checked = True
                                End If
                            End If
                        Next
                    End If
                Catch ex As Exception
                    p_blnUpdateManually = True
                End Try
            End If

            Me.PrepareRunningOrderDataToSend()

            ClearStatusCheckboxes()

            Me.chkOfficial.Checked = jmData.Official

            'If we just set status to official, do not let
            'other status change it to something else
            If Me.chkOfficial.Checked Then
                m_FlagChangeEvent = False
            End If

            Me.chkDeadHeat.Checked = jmData.Deaheat
            Me.chkPhoto.Checked = jmData.Photo
            Me.chkInqObj.Checked = jmData.Inquiry And jmData.Objection
            Me.chkInq.Checked = jmData.Inquiry
            Me.chkObj.Checked = jmData.Objection

            m_FlagChangeEvent = True

            If (Not (jmData.Official Or jmData.Deaheat Or jmData.Photo Or jmData.Objection Or jmData.Inquiry)) Then
                Try
                    ClearStatusCheckboxes()
                Catch ex As Exception
                    m_FlagChangeEvent = True
                End Try
                txtOfficial.Text = ""
                Me.PrepareStatusDataToSend()
                If (txtRaceStatus.Text.Trim = "HOLD ALL TICKETS") Then
                    txtRaceStatus.Text = ""
                    Me.PrepareExoticsToSend(1, False, True)
                End If
            End If

        Catch ex As Exception
            '
        End Try
    End Sub

    Private Sub ClearStatusCheckboxes()
        m_FlagChangeEvent = False
        ClearStatus("")
        m_FlagChangeEvent = True
    End Sub

    Private Delegate Sub ResetRaceStatusJMBindingSource()
    Private Sub myCommSvr_DisplayRaceStatusJM()
        If Me.InvokeRequired Then
            Me.Invoke(New ResetRaceStatusJMBindingSource(AddressOf myCommSvr_DisplayRaceStatusJM))
        Else
            '
        End If

        Dim strStatus As String
        strStatus = Me.myCommSvr.p_strCurrentStatusJM
        Dim strRunnersFlashingStatus As String = ""

        Try
            p_blnUpdateManually = False
            For Each ctl As Control In Me.gbOrder.Controls
                If TypeOf ctl Is Windows.Forms.CheckBox Then
                    CType(ctl, Windows.Forms.CheckBox).Checked = False
                End If
            Next
            p_blnUpdateManually = True
        Catch ex As Exception
            p_blnUpdateManually = True
        End Try

        Try
            strRunnersFlashingStatus = Me.myCommSvr.p_CurrentRunnersStatusList

            If (strRunnersFlashingStatus.Trim <> "") Then
                Try
                    p_blnUpdateManually = False

                    Dim vntInfoArray() As String = Split(strRunnersFlashingStatus, ",")
                    Dim strTemp As String = ""
                    For i As Integer = 0 To vntInfoArray.Length - 1
                        strTemp = vntInfoArray(i).Trim.ToUpper()
                        If (strTemp = (Me.txtRunning1.Text & Me.txtRunning1A.Text).Trim.ToUpper()) Then
                            Me.chkResults1.Checked = True
                        ElseIf (strTemp = (Me.txtRunning2.Text & Me.txtRunning2A.Text).Trim.ToUpper()) Then
                            Me.chkResults2.Checked = True
                        ElseIf (strTemp = (Me.txtRunning3.Text & Me.txtRunning3A.Text).Trim.ToUpper()) Then
                            Me.chkResults3.Checked = True
                        ElseIf (strTemp = (Me.txtRunning4.Text & Me.txtRunning4A.Text).Trim.ToUpper()) Then
                            Me.chkResults4.Checked = True
                        End If
                    Next

                    p_blnUpdateManually = True
                Catch ex As Exception
                    p_blnUpdateManually = True
                End Try
            End If

            Me.PrepareRunningOrderDataToSend()

            If (InStr(strStatus, "F") > 0) Then
                Me.chkOfficial.Checked = True
            ElseIf (InStr(strStatus, "D") > 0) Then
                Me.chkDeadHeat.Checked = True
            ElseIf (InStr(strStatus, "P") > 0) Then
                Me.chkPhoto.Checked = True
            ElseIf ((InStr(strStatus, "I") > 0) And (InStr(strStatus, "O") > 0)) Then
                Me.chkInqObj.Checked = True
            ElseIf (InStr(strStatus, "I") > 0) Then
                Me.chkInq.Checked = True
            ElseIf (InStr(strStatus, "O") > 0) Then
                Me.chkObj.Checked = True
            Else
                Try
                    m_FlagChangeEvent = False
                    ClearStatus("")
                    m_FlagChangeEvent = True
                Catch ex As Exception
                    m_FlagChangeEvent = True
                End Try
                txtOfficial.Text = ""
                Me.PrepareStatusDataToSend()
                If (txtRaceStatus.Text.Trim = "HOLD ALL TICKETS") Then
                    txtRaceStatus.Text = ""
                    Me.PrepareExoticsToSend(1, False, True)
                End If
            End If

        Catch ex As Exception
            '
        End Try
    End Sub

    Private Sub DisplayWPS()
        Me.bsWin.DataSource = Me.RaceDisplayDataset.WIN
        Me.bsPlace.DataSource = Me.RaceDisplayDataset.PLACE
        Me.bsShow.DataSource = Me.RaceDisplayDataset.SHOW
        '
        Me.bsWin.ResetBindings(False)
        Me.bsPlace.ResetBindings(False)
        Me.bsShow.ResetBindings(False)
        '
        Me.bsWin.EndEdit()
        Me.bsPlace.EndEdit()
        Me.bsPlace.EndEdit()
        '
        Me.PrepareWPSDataToSend()
    End Sub

    Private Sub DisplayExotics(intRace As Integer)
        Me.PrepareExoticsToSend(intRace, True, False)
    End Sub

    'Private Delegate Sub ResetExotics()

    'Private Sub myCommSvr_DisplayExotics()
    '    If Me.InvokeRequired Then
    '        Me.Invoke(New ResetExotics(AddressOf myCommSvr_DisplayExotics))
    '    Else

    '        ''Me.bsDD.DataSource = Me.RaceDisplayDataset.DAILYDOUBLE
    '        ''Me.bsDD.ResetBindings(False)
    '        ''Me.bsDD.EndEdit()
    '        ''
    '        'Me.bsPerfecta.DataSource = Me.RaceDisplayDataset.PERFECTA
    '        'Me.bsPerfecta.ResetBindings(False)
    '        'Me.bsPerfecta.EndEdit()
    '        ''
    '        ''Me.bsBet3.DataSource = Me.RaceDisplayDataset.BET3
    '        ''Me.bsBet3.ResetBindings(False)
    '        ''Me.bsBet3.EndEdit()
    '        ''
    '        'Me.bsTrifecta.DataSource = Me.RaceDisplayDataset.TRIFECTA
    '        'Me.bsTrifecta.ResetBindings(False)
    '        'Me.bsTrifecta.EndEdit()

    '        'Static IndexPerfecta As Integer
    '        'Static IndexTrifecta As Integer
    '        'Dim strCurrentExotic As String = Me.myCurrentExotic

    '        'If (strCurrentExotic = CurrentExotic.Exacta) Then
    '        '    If ((Me.myCommSvr.myNumberExacta > 1) AndAlso _
    '        '        (IndexPerfecta < Me.myCommSvr.myNumberExacta)) Then
    '        '        Me.myCurrentExotic = CurrentExotic.Trifecta
    '        '        IndexPerfecta += 1
    '        '        IndexTrifecta = 0
    '        '        If (IndexPerfecta > Me.myCommSvr.myNumberExacta) Then
    '        '            IndexTrifecta = 0
    '        '        End If
    '        '    End If
    '        'ElseIf (strCurrentExotic = CurrentExotic.Trifecta) Then
    '        '    If ((Me.myCommSvr.myNumberTrifecta > 1) AndAlso _
    '        '     (IndexTrifecta < Me.myCommSvr.myNumberTrifecta)) Then
    '        '        Me.myCurrentExotic = CurrentExotic.Exacta
    '        '        IndexTrifecta += 1
    '        '        IndexPerfecta = 0
    '        '        If (IndexTrifecta > Me.myCommSvr.myNumberTrifecta) Then
    '        '            IndexPerfecta = 0
    '        '        End If
    '        '    End If
    '        'ElseIf (strCurrentExotic = CurrentExotic.None) Then
    '        '    IndexPerfecta = 0
    '        '    IndexTrifecta = 0
    '        'End If

    '        ''now check which one to show
    '        'Me.myCurrentExotic = Me.GetCurrentExotic
    '        'Select Case Me.myCurrentExotic
    '        '    Case CurrentExotic.Exacta
    '        '        'turn on exacta
    '        '        Me.rbExacta.Checked = True
    '        '        Me.PrepareExactaDataToSend()
    '        '    Case CurrentExotic.Trifecta
    '        '        'turn on trifecta
    '        '        Me.rbTrifecta.Checked = True
    '        '        Me.PrepareTrifectaDataToSend()
    '        '    Case CurrentExotic.None
    '        '        Me.ClearExacta()
    '        '        Me.ClearTrifecta()
    '        '        Me.ClearItems(19)
    '        'End Select
    '        'here will start timerofficial to clear official results when mtp = 5
    '        'Me.timerOfficial.Enabled = True
    '    End If
    '    Me.timerOfficial.Enabled = True
    'End Sub

    Private Delegate Sub ResetTiming()

    Private Sub myCommSvr_DisplayTiming()
        If Me.InvokeRequired Then
            Me.Invoke(New ResetTiming(AddressOf myCommSvr_DisplayTiming))
        Else
            Me.bsTimingFinish.DataSource = Me.RaceDisplayDataset.TIMINGFINISH
            Me.bsTimingFinish.ResetBindings(False)
            Me.bsTimingFinish.EndEdit()
            '
            Me.bsTimingMile.DataSource = Me.RaceDisplayDataset.TIMINGMILE
            Me.bsTimingMile.ResetBindings(False)
            Me.bsTimingMile.EndEdit()
            '
            Me.bsTiming34.DataSource = Me.RaceDisplayDataset.TIMING34
            Me.bsTiming34.ResetBindings(False)
            Me.bsTiming34.EndEdit()
            '
            Me.bsTiming12.DataSource = Me.RaceDisplayDataset.TIMING12
            Me.bsTiming12.ResetBindings(False)
            Me.bsTiming12.EndEdit()
            '
            Me.bsTiming14.DataSource = Me.RaceDisplayDataset.TIMING14
            Me.bsTiming14.ResetBindings(False)
            Me.bsTiming14.EndEdit()
            '
        End If
        Me.PrepareTeletimerToSend()
    End Sub

    Private Function GetCurrentExotic() As CurrentExotic

        'If ((Not Me.myCurrentExotic = CurrentExotic.Exacta) AndAlso _
        '       (Not Me.RaceDisplayDataset.PERFECTA.Select("").Count = 0)) OrElse _
        '       ((Not Me.RaceDisplayDataset.TRIFECTA.Select("").Count > 0) AndAlso _
        '       (Not Me.RaceDisplayDataset.PERFECTA.Select("").Count = 0)) Then
        '    Return CurrentExotic.Exacta
        'ElseIf ((Not Me.myCurrentExotic = CurrentExotic.Trifecta) AndAlso _
        '    (Not Me.RaceDisplayDataset.TRIFECTA.Select("").Count = 0)) OrElse _
        '    ((Not Me.RaceDisplayDataset.PERFECTA.Select("").Count > 0) AndAlso _
        '    (Not Me.RaceDisplayDataset.TRIFECTA.Select("").Count = 0)) Then
        '    Return CurrentExotic.Trifecta
        'ElseIf Me.RaceDisplayDataset.PERFECTA.Select("").Count = 0 AndAlso _
        '       Me.RaceDisplayDataset.TRIFECTA.Select("").Count = 0 Then
        '    Return CurrentExotic.None
        'End If
        ''just in case
        'Return CurrentExotic.None
    End Function

    Private Sub RefreshROResult(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkResults2.CheckedChanged, chkResults1.CheckedChanged, chkResults4.CheckedChanged, chkResults3.CheckedChanged
        If (p_blnUpdateManually) Then
            'If (Not myCommSvr.m_ApplicationBussy) Then
            If myCommSvr.m_ApplicationBussy Then
                Application.DoEvents()
            End If
            Try
                myCommSvr.m_ApplicationBussy = True
                Me.PrepareRunningOrderDataToSend()
                myCommSvr.m_ApplicationBussy = False
            Catch
                myCommSvr.m_ApplicationBussy = False
            End Try
            'Else
            '    Dim cb As CheckBox = DirectCast(sender, CheckBox)
            '    RemoveHandler cb.CheckedChanged, AddressOf RefreshROResult
            '    cb.Checked = Not cb.Checked
            '    AddHandler cb.CheckedChanged, AddressOf RefreshROResult
            'End If
        End If
    End Sub

    Private Sub ClearExacta()
        Me.rbExacta.Checked = False
        For Each c As Control In Me.gbPerfecta.Controls
            If TypeOf c Is Windows.Forms.TextBox Then c.Text = ""
        Next
        '
        Me.bsPerfecta.EndEdit()
    End Sub

    Private Sub ClearTrifecta()
        Me.rbTrifecta.Checked = False
        For Each c As Control In Me.gbTrifecta.Controls
            If TypeOf c Is Windows.Forms.TextBox Then c.Text = ""
        Next
        '
        Me.bsTrifecta.EndEdit()
    End Sub

#End Region

#Region " Tania's Additional Methods "

    Private Sub btnClearAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearAll.Click
        'tania todo
        ClearResultsToolStripMenuItem_Click(sender, e)
        ClearAllToolStripMenuItem_Click(sender, e)
    End Sub

    Private Sub btnClearOdds_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearOdds.Click
        ClearOddsToolStripMenuItem_Click(sender, e)
    End Sub

    Private Sub btnClearResults_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearResults.Click
        ClearResultsToolStripMenuItem_Click(sender, e)
    End Sub

    Private Sub btnClearRO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearRO.Click
        ClearRunningOrderToolStripMenuItem_Click(sender, e)
    End Sub

    Private Sub btnClearExotics_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearExotics.Click
        If (Not myCommSvr.m_ApplicationBussy) Then
            Try
                myCommSvr.m_ApplicationBussy = True
                myCommSvr.p_blnResultsOut = False
                myCommSvr.p_intRaceResultsOut = 0
                Me.timerOfficial.Enabled = False
                txtRaceStatus.Text = ""
                PrepareExoticsToSend(1, False, True)
                'ClearExoticsToolStripMenuItem_Click(sender, e)
                myCommSvr.m_ApplicationBussy = False
            Catch
                myCommSvr.m_ApplicationBussy = False
            End Try
        End If
    End Sub

    Private Sub btnRefreshPools_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefreshPools.Click
        tmrWPSPools.Enabled = True
    End Sub

    Private Sub btnClearPools_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearPools.Click
        If (Not myCommSvr.m_ApplicationBussy) Then
            Try
                myCommSvr.m_ApplicationBussy = True
                tmrWPSPools.Enabled = False
                p_blnUpdateManually = False

                For Each ctl As Control In Me.gbWPSPools.Controls
                    If TypeOf ctl Is Windows.Forms.TextBox Then
                        ctl.Text = ""
                    End If
                Next

                Me.bsWPSPools.EndEdit()

                Me.ClearItems(36)
                Me.ClearItems(38)
                Me.ClearItems(40)
                Me.ClearItems(42)
                '
                txtPoolType.Text = ""
                PrepareWPSPoolHeaderToSend(True)
                '
                p_blnUpdateManually = True
                myCommSvr.m_ApplicationBussy = False
            Catch
                p_blnUpdateManually = True
                myCommSvr.m_ApplicationBussy = False
            End Try
        End If
    End Sub

    Private Sub btnClearTimming_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearTimming.Click
        ClearTimingToolStripMenuItem_Click(sender, e)
    End Sub

    Private Sub btnTestDisplays_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTestDisplays.Click
        If (Not myCommSvr.m_ApplicationBussy) Then
            Try
                myCommSvr.m_ApplicationBussy = True
                myCommSvr.p_blnResultsOut = False
                myCommSvr.p_intRaceResultsOut = 0
                Me.timerOfficial.Enabled = False
                ClearAll()
                If (btnTestDisplays.Text = "Test Displays") Then
                    btnTestDisplays.Text = "Stop Test"
                    m_intCtr = 1
                    'm_FlagExacta = True
                    Me.tmrTest_Tick(sender, e)
                Else
                    tmrTest.Enabled = False
                    btnTestDisplays.Text = "Test Displays"
                    myCommSvr.m_ApplicationBussy = False
                End If
            Catch ex As Exception
                tmrTest.Enabled = False
                btnTestDisplays.Text = "Test Displays"
                ClearAll()
                myCommSvr.m_ApplicationBussy = False
            End Try
        Else
            If (btnTestDisplays.Text = "Stop Test") Then
                tmrTest.Enabled = False
                btnTestDisplays.Text = "Test Displays"
                ClearAll()
                myCommSvr.m_ApplicationBussy = False
            End If
        End If
    End Sub

    Private Sub tmrTest_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrTest.Tick
        tmrTest.Enabled = False
        SendRO()
        SendWPS()
        'SendExotics()
        SendOdds()
        Me.PrepareRunningOrderDataToSend()
        Me.PrepareWPSDataToSend()
        'If (m_FlagExacta) Then
        'PrepareExactaDataToSend()
        'Else
        'PrepareTrifectaDataToSend()
        'End If
        m_intCtr = m_intCtr + 1
        If m_intCtr > 16 Then
            'If m_intCtr > 12 Then
            m_intCtr = 0
            m_FlagExacta = Not m_FlagExacta
        End If
        tmrTest.Enabled = True
    End Sub

    Private Sub SendOdds()
        Dim myArrayOdds(11) As Integer
        Select Case m_intCtr
            Case 0, 1, 4, 7, 10, 13, 16
                'rbFast.Checked = True
            Case 2, 5, 8, 11, 14
                'rbSlow.Checked = True
            Case 3, 6, 9, 12, 15
                'rbSloppy.Checked = True
        End Select
        If m_intCtr >= 1 Then
            Me.txtRaceA.Text = "8"
            Me.txtOdds5.Text = "8"
        Else
            Me.txtRaceA.Text = ""
            Me.txtOdds5.Text = ""
        End If
        If m_intCtr >= 2 Then
            Me.txtRaceB.Text = "8"
            Me.txtOdds5A.Text = "8"
            myArrayOdds(4) = 1
        Else
            Me.txtRaceB.Text = ""
            Me.txtOdds5A.Text = ""
            myArrayOdds(4) = 0
        End If
        If m_intCtr >= 3 Then
            Me.txtMTPA.Text = "8"
            Me.txtOdds6.Text = "8"
        Else
            Me.txtMTPA.Text = ""
            Me.txtOdds6.Text = ""
        End If
        If m_intCtr >= 4 Then
            Me.txtMTPB.Text = "8"
            Me.txtOdds6A.Text = "8"
            myArrayOdds(5) = 1
        Else
            Me.txtMTPB.Text = ""
            Me.txtOdds6A.Text = ""
            myArrayOdds(5) = 0
        End If
        If m_intCtr >= 5 Then
            Me.txtTODA.Text = "8"
            Me.txtOdds7.Text = "8"
        Else
            Me.txtTODA.Text = ""
            Me.txtOdds7.Text = ""
        End If
        If m_intCtr >= 6 Then
            Me.txtTODB.Text = "8"
            Me.txtOdds7A.Text = "8"
            myArrayOdds(6) = 1
        Else
            Me.txtTODB.Text = ""
            Me.txtOdds7A.Text = ""
            myArrayOdds(6) = 0
        End If
        If m_intCtr >= 7 Then
            Me.txtTODC.Text = "8"
            Me.txtOdds8.Text = "8"
        Else
            Me.txtTODC.Text = ""
            Me.txtOdds8.Text = ""
        End If
        If m_intCtr >= 8 Then
            Me.txtTODD.Text = "8"
            Me.txtOdds8A.Text = "8"
            myArrayOdds(7) = 1
        Else
            Me.txtTODD.Text = ""
            Me.txtOdds8A.Text = ""
            myArrayOdds(7) = 0
        End If
        If m_intCtr >= 9 Then
            Me.txtOdds1.Text = "8"
            Me.txtOdds9.Text = "8"
        Else
            Me.txtOdds1.Text = ""
            Me.txtOdds9.Text = ""
        End If
        If m_intCtr >= 10 Then
            Me.txtOdds1A.Text = "8"
            myArrayOdds(0) = 1
            Me.txtOdds9A.Text = "8"
            myArrayOdds(8) = 1
        Else
            Me.txtOdds1A.Text = ""
            myArrayOdds(0) = 0
            Me.txtOdds9A.Text = ""
            myArrayOdds(8) = 0
        End If
        If m_intCtr >= 11 Then
            Me.txtOdds2.Text = "8"
            Me.txtOdds10.Text = "8"
        Else
            Me.txtOdds2.Text = ""
            Me.txtOdds10.Text = ""
        End If
        If m_intCtr >= 12 Then
            Me.txtOdds2A.Text = "8"
            myArrayOdds(1) = 1
            Me.txtOdds10A.Text = "8"
            myArrayOdds(9) = 1
        Else
            Me.txtOdds2A.Text = ""
            myArrayOdds(1) = 0
            Me.txtOdds10A.Text = ""
            myArrayOdds(9) = 0
        End If
        If m_intCtr >= 13 Then
            Me.txtOdds3.Text = "8"
            Me.txtOdds11.Text = "8"
        Else
            Me.txtOdds3.Text = ""
            Me.txtOdds11.Text = ""
        End If
        If m_intCtr >= 14 Then
            Me.txtOdds3A.Text = "8"
            myArrayOdds(2) = 1
            Me.txtOdds11A.Text = "8"
            myArrayOdds(10) = 1
        Else
            Me.txtOdds3A.Text = ""
            myArrayOdds(2) = 0
            Me.txtOdds11A.Text = ""
            myArrayOdds(10) = 0
        End If
        If m_intCtr >= 15 Then
            Me.txtOdds4.Text = "8"
            Me.txtOdds12.Text = "8"
        Else
            Me.txtOdds4.Text = ""
            Me.txtOdds12.Text = ""
        End If
        If m_intCtr >= 16 Then
            Me.txtOdds4A.Text = "8"
            myArrayOdds(3) = 1
            Me.txtOdds12A.Text = "8"
            myArrayOdds(11) = 1
        Else
            Me.txtOdds4A.Text = ""
            myArrayOdds(3) = 0
            Me.txtOdds12A.Text = ""
            myArrayOdds(11) = 0
        End If
        Me.lbl1.Visible = myArrayOdds(0) = 1
        Me.lbl2.Visible = myArrayOdds(1) = 1
        Me.lbl3.Visible = myArrayOdds(2) = 1
        Me.lbl4.Visible = myArrayOdds(3) = 1
        Me.lbl5.Visible = myArrayOdds(4) = 1
        Me.lbl6.Visible = myArrayOdds(5) = 1
        Me.lbl7.Visible = myArrayOdds(6) = 1
        Me.lbl8.Visible = myArrayOdds(7) = 1
        Me.lbl9.Visible = myArrayOdds(8) = 1
        Me.lbl10.Visible = myArrayOdds(9) = 1
        Me.lbl11.Visible = myArrayOdds(10) = 1
        Me.lbl12.Visible = myArrayOdds(11) = 1
        Me.bsRace.EndEdit()
        Me.bsMTP.EndEdit()
        Me.bsTOD.EndEdit()
        Me.bsODDS.EndEdit()
        Me.PrepareOddsDataToSend(myArrayOdds)
        'moment of truth
        For i As Integer = 1 To Me.myColMessages.Count
            Me.myComPort.Output(Me.myColMessages(i))
        Next
        Me.PrepareTimeOfDayToSend()
    End Sub

    Private Sub SendRO()
        Try
            p_blnUpdateManually = False
            If m_intCtr >= 1 Then
                Me.txtRunning1.Text = "8"
            Else
                Me.txtRunning1.Text = ""
            End If
            If m_intCtr >= 2 Then
                Me.txtRunning1A.Text = "8"
            Else
                Me.txtRunning1A.Text = ""
            End If
            If m_intCtr >= 3 Then
                Me.chkResults1.Checked = True
            Else
                Me.chkResults1.Checked = False
            End If
            If m_intCtr >= 4 Then
                Me.chkOfficial.Checked = True
            Else
                Me.chkOfficial.Checked = False
            End If
            If m_intCtr >= 5 Then
                Me.txtRunning2.Text = "8"
            Else
                Me.txtRunning2.Text = ""
            End If
            If m_intCtr >= 6 Then
                Me.txtRunning2A.Text = "8"
            Else
                Me.txtRunning2A.Text = ""
            End If
            If m_intCtr >= 7 Then
                Me.chkResults2.Checked = True
            Else
                Me.chkResults2.Checked = False
            End If
            If m_intCtr >= 8 Then
                Me.chkDeadHeat.Checked = True
            Else
                Me.chkDeadHeat.Checked = False
            End If
            If m_intCtr >= 9 Then
                Me.txtRunning3.Text = "8"
            Else
                Me.txtRunning3.Text = ""
            End If
            If m_intCtr >= 10 Then
                Me.txtRunning3A.Text = "8"
            Else
                Me.txtRunning3A.Text = ""
            End If
            If m_intCtr >= 11 Then
                Me.chkResults3.Checked = True
            Else
                Me.chkResults3.Checked = False
            End If
            If m_intCtr >= 12 Then
                Me.chkPhoto.Checked = True
            Else
                Me.chkPhoto.Checked = False
            End If
            If m_intCtr >= 13 Then
                Me.txtRunning4.Text = "8"
            Else
                Me.txtRunning4.Text = ""
            End If
            If m_intCtr >= 14 Then
                Me.txtRunning4A.Text = "8"
            Else
                Me.txtRunning4A.Text = ""
            End If
            If m_intCtr >= 15 Then
                Me.chkResults4.Checked = True
            Else
                Me.chkResults4.Checked = False
            End If
            If m_intCtr >= 16 Then
                Me.chkInqObj.Checked = True
            Else
                Me.chkInqObj.Checked = False
            End If
            Me.bsRunningOrder.EndEdit()
            Me.bsStatus.EndEdit()
            p_blnUpdateManually = True
        Catch ex As Exception
            p_blnUpdateManually = True
        End Try
    End Sub

    Private Sub SendWPS()
        If m_intCtr >= 1 Then
            Me.txtWin1A.Text = "8"
            Me.txtPlace1A.Text = "8"
            Me.txtShow1A.Text = "8"
            Me.txtShow4A.Text = "8"
        Else
            Me.txtWin1A.Text = ""
            Me.txtPlace1A.Text = ""
            Me.txtShow1A.Text = ""
            Me.txtShow4A.Text = ""
        End If
        If m_intCtr >= 2 Then
            Me.txtWin1B.Text = "8"
            Me.txtPlace1B.Text = "8"
            Me.txtShow1B.Text = "8"
            Me.txtShow4B.Text = "8"
        Else
            Me.txtWin1B.Text = ""
            Me.txtPlace1B.Text = ""
            Me.txtShow1B.Text = ""
            Me.txtShow4B.Text = ""
        End If
        If m_intCtr >= 3 Then
            Me.txtWin1C.Text = "8"
            Me.txtPlace1C.Text = "8"
            Me.txtShow1C.Text = "8"
            Me.txtShow4C.Text = "8"
        Else
            Me.txtWin1C.Text = ""
            Me.txtPlace1C.Text = ""
            Me.txtShow1C.Text = ""
            Me.txtShow4C.Text = ""
        End If
        If m_intCtr >= 4 Then
            Me.txtWin1D.Text = "8"
            Me.txtPlace1D.Text = "8"
            Me.txtShow1D.Text = "8"
            Me.txtShow4D.Text = "8"
        Else
            Me.txtWin1D.Text = ""
            Me.txtPlace1D.Text = ""
            Me.txtShow1D.Text = ""
            Me.txtShow4D.Text = ""
        End If
        If m_intCtr >= 5 Then
            Me.txtWin1E.Text = "8"
            Me.txtPlace1E.Text = "8"
            Me.txtShow1E.Text = "8"
            Me.txtShow4E.Text = "8"
        Else
            Me.txtWin1E.Text = ""
            Me.txtPlace1E.Text = ""
            Me.txtShow1E.Text = ""
            Me.txtShow4E.Text = ""
        End If
        If m_intCtr >= 6 Then
            Me.txtWin1F.Text = "8"
            Me.txtPlace2A.Text = "8"
            Me.txtShow2A.Text = "8"
        Else
            Me.txtWin1F.Text = ""
            Me.txtPlace2A.Text = ""
            Me.txtShow2A.Text = ""
        End If
        If m_intCtr >= 7 Then
            Me.txtWin2A.Text = "8"
            Me.txtPlace2B.Text = "8"
            Me.txtShow2B.Text = "8"
        Else
            Me.txtWin2A.Text = ""
            Me.txtPlace2B.Text = ""
            Me.txtShow2B.Text = ""
        End If
        If m_intCtr >= 8 Then
            Me.txtWin2B.Text = "8"
            Me.txtPlace2C.Text = "8"
            Me.txtShow2C.Text = "8"
        Else
            Me.txtWin2B.Text = ""
            Me.txtPlace2C.Text = ""
            Me.txtShow2C.Text = ""
        End If
        If m_intCtr >= 9 Then
            Me.txtWin2C.Text = "8"
            Me.txtPlace2D.Text = "8"
            Me.txtShow2D.Text = "8"
        Else
            Me.txtWin2C.Text = ""
            Me.txtPlace2D.Text = ""
            Me.txtShow2D.Text = ""
        End If
        If m_intCtr >= 10 Then
            Me.txtWin2D.Text = "8"
            Me.txtPlace2E.Text = "8"
            Me.txtShow2E.Text = "8"
        Else
            Me.txtWin2D.Text = ""
            Me.txtPlace2E.Text = ""
            Me.txtShow2E.Text = ""
        End If
        If m_intCtr >= 11 Then
            Me.txtWin2E.Text = "8"
            Me.txtPlace3A.Text = "8"
            Me.txtShow3A.Text = "8"
        Else
            Me.txtWin2E.Text = ""
            Me.txtPlace3A.Text = ""
            Me.txtShow3A.Text = ""
        End If
        If m_intCtr >= 12 Then
            Me.txtWin2F.Text = "8"
            Me.txtPlace3B.Text = "8"
            Me.txtShow3B.Text = "8"
        Else
            Me.txtWin2F.Text = ""
            Me.txtPlace3B.Text = ""
            Me.txtShow3B.Text = ""
        End If
        If m_intCtr >= 13 Then
            Me.txtPlace3C.Text = "8"
            Me.txtShow3C.Text = "8"
        Else
            Me.txtPlace3C.Text = ""
            Me.txtShow3C.Text = ""
        End If
        If m_intCtr >= 14 Then
            Me.txtPlace3D.Text = "8"
            Me.txtShow3D.Text = "8"
        Else
            Me.txtPlace3D.Text = ""
            Me.txtShow3D.Text = ""
        End If
        If m_intCtr >= 15 Then
            Me.txtPlace3E.Text = "8"
            Me.txtShow3E.Text = "8"
        Else
            Me.txtPlace3E.Text = ""
            Me.txtShow3E.Text = ""
        End If
        Me.bsWin.EndEdit()
        Me.bsPlace.EndEdit()
        Me.bsShow.EndEdit()
    End Sub

    Private Sub SendExotics()
        If (m_FlagExacta) Then
            If m_intCtr >= 1 Then
                '
                Me.rbExacta.Checked = True
                Me.rbTrifecta.Checked = False
                '
                Me.txtPerfecta1A.Text = "8"
                '
                Me.txtTrifecta1A.Text = ""
                Me.txtTrifecta1B.Text = ""
                Me.txtTrifecta2A.Text = ""
                Me.txtTrifecta2B.Text = ""
                Me.txtTrifecta3A.Text = ""
                Me.txtTrifecta3B.Text = ""
                Me.txtTrifectaAmountA.Text = ""
                Me.txtTrifectaAmountB.Text = ""
                Me.txtTrifectaAmountC.Text = ""
                Me.txtTrifectaAmountD.Text = ""
                Me.txtTrifectaAmountE.Text = ""
                Me.txtTrifectaAmountF.Text = ""
                Me.txtTrifectaAmountG.Text = ""
                '
            Else
                '
                Me.rbExacta.Checked = False
                Me.rbTrifecta.Checked = False
                '
                Me.txtPerfecta1A.Text = ""
            End If
            If m_intCtr >= 2 Then
                Me.txtPerfecta1B.Text = "8"
            Else
                Me.txtPerfecta1B.Text = ""
            End If
            If m_intCtr >= 3 Then
                Me.txtPerfecta2A.Text = "8"
            Else
                Me.txtPerfecta2A.Text = ""
            End If
            If m_intCtr >= 4 Then
                Me.txtPerfecta2B.Text = "8"
            Else
                Me.txtPerfecta2B.Text = ""
            End If
            If m_intCtr >= 5 Then
                Me.txtPerfectaAmountA.Text = "8"
            Else
                Me.txtPerfectaAmountA.Text = ""
            End If
            If m_intCtr >= 6 Then
                Me.txtPerfectaAmountB.Text = "8"
            Else
                Me.txtPerfectaAmountB.Text = ""
            End If
            If m_intCtr >= 7 Then
                Me.txtPerfectaAmountC.Text = "8"
            Else
                Me.txtPerfectaAmountC.Text = ""
            End If
            If m_intCtr >= 8 Then
                Me.txtPerfectaAmountD.Text = "8"
            Else
                Me.txtPerfectaAmountD.Text = ""
            End If
            If m_intCtr >= 9 Then
                Me.txtPerfectaAmountE.Text = "8"
            Else
                Me.txtPerfectaAmountE.Text = ""
            End If
            If m_intCtr >= 10 Then
                Me.txtPerfectaAmountF.Text = "8"
            Else
                Me.txtPerfectaAmountF.Text = ""
            End If
            If m_intCtr >= 11 Then
                Me.txtPerfectaAmountG.Text = "8"
            Else
                Me.txtPerfectaAmountG.Text = ""
            End If
        Else
            If m_intCtr >= 1 Then
                '
                Me.rbExacta.Checked = False
                Me.rbTrifecta.Checked = True
                '
                Me.txtTrifecta1A.Text = "8"
                '
                Me.txtPerfecta1A.Text = ""
                Me.txtPerfecta1B.Text = ""
                Me.txtPerfecta2A.Text = ""
                Me.txtPerfecta2B.Text = ""
                Me.txtPerfectaAmountA.Text = ""
                Me.txtPerfectaAmountB.Text = ""
                Me.txtPerfectaAmountC.Text = ""
                Me.txtPerfectaAmountD.Text = ""
                Me.txtPerfectaAmountE.Text = ""
                Me.txtPerfectaAmountF.Text = ""
                Me.txtPerfectaAmountG.Text = ""
                '
            Else
                '
                Me.rbExacta.Checked = False
                Me.rbTrifecta.Checked = False
                '
                Me.txtTrifecta1A.Text = ""
            End If
            If m_intCtr >= 2 Then
                Me.txtTrifecta1B.Text = "8"
            Else
                Me.txtTrifecta1B.Text = ""
            End If
            If m_intCtr >= 3 Then
                Me.txtTrifecta2A.Text = "8"
            Else
                Me.txtTrifecta2A.Text = ""
            End If
            If m_intCtr >= 4 Then
                Me.txtTrifecta2B.Text = "8"
            Else
                Me.txtTrifecta2B.Text = ""
            End If
            If m_intCtr >= 5 Then
                Me.txtTrifecta3A.Text = "8"
            Else
                Me.txtTrifecta3A.Text = ""
            End If
            If m_intCtr >= 6 Then
                Me.txtTrifecta3B.Text = "8"
            Else
                Me.txtTrifecta3B.Text = ""
            End If
            If m_intCtr >= 7 Then
                Me.txtTrifectaAmountA.Text = "8"
            Else
                Me.txtTrifectaAmountA.Text = ""
            End If
            If m_intCtr >= 8 Then
                Me.txtTrifectaAmountB.Text = "8"
            Else
                Me.txtTrifectaAmountB.Text = ""
            End If
            If m_intCtr >= 9 Then
                Me.txtTrifectaAmountC.Text = "8"
            Else
                Me.txtTrifectaAmountC.Text = ""
            End If
            If m_intCtr >= 10 Then
                Me.txtTrifectaAmountD.Text = "8"
            Else
                Me.txtTrifectaAmountD.Text = ""
            End If
            If m_intCtr >= 11 Then
                Me.txtTrifectaAmountE.Text = "8"
            Else
                Me.txtTrifectaAmountE.Text = ""
            End If
            If m_intCtr >= 12 Then
                Me.txtTrifectaAmountF.Text = "8"
            Else
                Me.txtTrifectaAmountF.Text = ""
            End If
            If m_intCtr >= 13 Then
                Me.txtTrifectaAmountG.Text = "8"
            Else
                Me.txtTrifectaAmountG.Text = ""
            End If
        End If
        Me.bsPerfecta.EndEdit()
        Me.bsTrifecta.EndEdit()
    End Sub

    Private Sub cboMTPToClearRS_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.p_intMTPToClear = CInt(Me.cboMTPToClearRS.SelectedText)
    End Sub

#End Region

    Private Sub tmrWPSPools_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrWPSPools.Tick
        Dim intRace As Integer
        Dim intCycle As Integer = CInt(lblCycleNum.Text)
        Dim strType As String

        If (Not myCommSvr.m_ApplicationBussy) Then
            Try
                myCommSvr.m_ApplicationBussy = True
                intRace = Me.myCommSvr.p_intCurrentRace
                '
                Select Case intCycle
                    Case 1
                        strType = "WIN"
                    Case 2
                        strType = "PLC"
                    Case Else
                        strType = "SHW"
                End Select
                myCommSvr.UpdateWPSPools(intRace, strType)
                '
                DisplayWPSPools()
                '
                If intCycle = 3 Then intCycle = 1 Else intCycle += 1
                lblCycleNum.Text = CStr(intCycle)
                '
                myCommSvr.m_ApplicationBussy = False
            Catch
                myCommSvr.m_ApplicationBussy = False
            End Try
        End If
    End Sub

    Private Sub cboMTPToClearRS_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboMTPToClearRS.SelectedIndexChanged
        Me.p_intMTPToClear = CInt(cboMTPToClearRS.Text)
    End Sub

    Private Sub frmMain_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
        'For Each ctrl In Me.Controls
        '    ' center control on form here!
        '    If ctrl.Type = "textbox" Then
        '        gbMTP
        '    End If
        'Next
    End Sub

    Private Sub chkStatus_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkInq.CheckedChanged, chkPhoto.CheckedChanged, chkOfficial.CheckedChanged, chkObj.CheckedChanged, chkInqObj.CheckedChanged, chkDeadHeat.CheckedChanged
        If (Me.Visible) Then
            Try
                If (m_FlagChangeEvent) Then
                    If sender.Checked Then
                        Dim holdAllTickes As Boolean = (InStr(sender.Tag, "*") > 0)
                        Dim officialText = UCase(Replace(sender.Tag, "*", ""))

                        m_FlagChangeEvent = False
                        ClearStatus(sender.Name)
                        m_FlagChangeEvent = True
                        If holdAllTickes Then
                            txtRaceStatus.Text = "HOLD ALL TICKETS"
                            Me.PrepareExoticsToSend(1, False, False)
                        Else
                            If (txtRaceStatus.Text.Trim = "HOLD ALL TICKETS") Then
                                txtRaceStatus.Text = ""
                                Me.PrepareExoticsToSend(1, False, True)
                            End If
                        End If
                        txtOfficial.Text = officialText
                        Me.PrepareStatusDataToSend()
                    Else
                        If (txtRaceStatus.Text.Trim = "HOLD ALL TICKETS") Then
                            txtRaceStatus.Text = ""
                            Me.PrepareExoticsToSend(1, False, True)
                        End If
                        txtOfficial.Text = ""
                        Me.PrepareStatusDataToSend()
                    End If
                End If
            Catch ex As Exception
                m_FlagChangeEvent = True
            End Try
        End If
    End Sub

    Private Sub ClearStatus(ByVal strName As String)
        For Each ctl As Control In Me.gbOptions.Controls
            If TypeOf ctl Is Windows.Forms.CheckBox Then
                If (ctl.Name <> strName) Then
                    CType(ctl, Windows.Forms.CheckBox).Checked = False
                End If
            End If
        Next
    End Sub

    'Private Sub txtOfficial_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOfficial.TextChanged
    '    If (Me.Visible) Then
    '        If (Not myCommSvr.m_ApplicationBussy) Then
    '            Try
    '                myCommSvr.m_ApplicationBussy = True

    '                PrepareStatusDataToSend()
    '                '
    '                myCommSvr.m_ApplicationBussy = False
    '            Catch
    '                myCommSvr.m_ApplicationBussy = False
    '            End Try
    '        End If
    '    End If
    'End Sub

    Private Sub PrepareStatusDataToSend()
        Dim myDataToSend As String = ""
        Dim MessageToSend As String = ""
        'start of header
        Dim SOH As String = "01"
        'address, lets get it fom the settings, tania will later decide how to approach this
        Dim BoardAddress As String = "02"
        'control, will send it always "on" from here
        Dim BoardControl As String = "00"
        'dimming, it should be done through the settings as well
        Dim BoardDimming As String = "00"
        'pay load type, this value will be always 24 from here (hex(18))
        Dim PayloadType As String = "18"
        'end of transmission
        Dim EOT As String = "04"
        '
        Dim StringNumber As String = "01" 'string number for Official

        Dim strTemp As String = txtOfficial.Text
        strTemp = strTemp.ToString.PadRight(8, " ")

        Dim fontSize As Integer = If(strTemp.Length <= 8, 24, 16)
        Dim strDatatosend = PrepareDataThreeColorSignStatus(strTemp, Color.Red, fontSize) '24
        strDatatosend = strDatatosend.Replace(" ", "")

        Try
            If (strDatatosend <> "-1") Then
                '(between 1 and 240 bytes of Data
                Dim intNumbDataBytes As Integer
                intNumbDataBytes = (strDatatosend.Length / 2)
                Dim NumbDataBytes As String
                NumbDataBytes = ("0" & Hex(intNumbDataBytes))
                NumbDataBytes = NumbDataBytes.Substring(NumbDataBytes.Length - 2)

                myDataToSend = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}0000",
                                                     SOH, BoardAddress, BoardControl, BoardDimming, PayloadType,
                                                     StringNumber, NumbDataBytes, strDatatosend,
                                                     EOT)

                'convert it to char
                MessageToSend = Me.myCommSvr.oCommServerNet.GetMessageToSend(myDataToSend)
                '
                Me.myComPort.Output(MessageToSend)

                If Not Me.myComPort.ErrorMessage = "" Then Me.ShowError()
            End If

        Catch ex As Exception
            '
        End Try
        ''
    End Sub

    Private Sub btnDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDisplay.Click
        If (Not myCommSvr.m_ApplicationBussy) Then
            Try
                myCommSvr.m_ApplicationBussy = True
                myCommSvr.p_blnResultsOut = False
                myCommSvr.p_intRaceResultsOut = 0
                Me.timerOfficial.Enabled = False
                PrepareExoticsToSend(1, False, False)
                myCommSvr.m_ApplicationBussy = False
            Catch
                myCommSvr.m_ApplicationBussy = False
            End Try
        End If
    End Sub

    Private Sub btnDefaulFinalTimes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDefaulFinalTimes.Click
        'lblFIN.Text = "FIN"
        'lbl78th.Text = "MIL"
        'lbl58th.Text = "3/4"
        'lblHalf.Text = "1/2"
        'lblQtr.Text = "1/4"
    End Sub

    Private Sub btnSetFinalTimes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSetFinalTimes.Click
        'PrepareFinalsToSend(False)
    End Sub

    'Private Sub PrepareFinalsToSend(ByVal blnFlagClear As Boolean)
    '    'start of header
    '    Dim SOH As String = "01"
    '    'address, lets get it fom the settings, tania will later decide how to approach this
    '    Dim BoardAddress As String = "00" 'My.Settings.MiniBoardAddress
    '    'control, will send it always "on" from here
    '    Dim BoardControl As String = "00"
    '    'dimming, it should be done through the settings as well
    '    Dim BoardDimming As String = "00" '"0" & Hex(My.Settings.MiniBoardDimming)
    '    'pay load type, this value will be always 16 from here (hex(10))
    '    Dim PayloadType As String = "17"
    '    Dim NumbOfString As String = "03"
    '    Dim NumbOfBytes As String = "5A" '30 Bytes (6 Columns * 15 Characters)
    '    Dim StartColumn As String = "00"
    '    'end of transmission
    '    Dim EOT As String = "04"
    '    '
    '    Dim strTemp As String = " "
    '    If (blnFlagClear = False) Then
    '        strTemp = UCase(lblFIN.Text).PadRight(3, " ") & UCase(lbl78th.Text).PadRight(3, " ") & UCase(lbl58th.Text).PadRight(3, " ") & UCase(lblHalf.Text).PadRight(3, " ") & UCase(lblQtr.Text).PadRight(3, " ")
    '    End If
    '    strTemp = strTemp.ToString.PadRight(15, " ")

    '    Dim displaySegments() As Emac.DisplayMatrixUtils.DisplaySegment
    '    Dim strCol As String
    '    Dim strColTemp As String
    '    Dim col As String = ""
    '    Dim col1 As String = ""
    '    Dim col2 As String = ""
    '    Dim col3 As String = ""
    '    Dim col4 As String = ""
    '    Dim col5 As String = ""
    '    Dim col6 As String = ""
    '    Dim col7 As String = ""
    '    Dim col8 As String = ""
    '    Dim col9 As String = ""
    '    Dim col10 As String = ""
    '    Dim col11 As String = ""
    '    Dim col12 As String = ""
    '    Dim col13 As String = ""
    '    Dim col14 As String = ""
    '    Dim col15 As String = ""

    '    For intCtr As Integer = 1 To Len(strTemp)
    '        Try
    '            col = ""
    '            strColTemp = ""
    '            strCol = Mid(strTemp, intCtr, 1)
    '            If (strCol = "/") Then
    '                col = "000408102040"
    '            Else
    '                displaySegments = Emac.DisplayMatrixUtils.DisplaySegmentDictionary.GetDisplaySegments(strCol)
    '                For colIdx As Integer = 0 To 5
    '                    strColTemp = displaySegments(colIdx).HexRowValue
    '                    col = col & strColTemp.PadLeft(2, "0")
    '                Next colIdx
    '            End If
    '        Catch ex As Exception
    '            col = "000000000000"
    '        End Try
    '        If intCtr = 1 Then
    '            col1 = col
    '        ElseIf intCtr = 2 Then
    '            col2 = col
    '        ElseIf intCtr = 3 Then
    '            col3 = col
    '        ElseIf intCtr = 4 Then
    '            col4 = col
    '        ElseIf intCtr = 5 Then
    '            col5 = col
    '        ElseIf intCtr = 6 Then
    '            col6 = col
    '        ElseIf intCtr = 7 Then
    '            col7 = col
    '        ElseIf intCtr = 8 Then
    '            col8 = col
    '        ElseIf intCtr = 9 Then
    '            col9 = col
    '        ElseIf intCtr = 10 Then
    '            col10 = col
    '        ElseIf intCtr = 11 Then
    '            col11 = col
    '        ElseIf intCtr = 12 Then
    '            col12 = col
    '        ElseIf intCtr = 13 Then
    '            col13 = col
    '        ElseIf intCtr = 14 Then
    '            col14 = col
    '        ElseIf intCtr = 15 Then
    '            col15 = col
    '        End If
    '    Next intCtr
    '    '
    '    'build the string to send
    '    Dim MessageToSend As String = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}{19}{20}{21}{22}{23}{24}{25}", _
    '                                             SOH, BoardAddress, BoardControl, BoardDimming, PayloadType, _
    '                                             NumbOfString, StartColumn, NumbOfBytes, _
    '                                             col1, col2, col3, col4, col5, col6, _
    '                                             col7, col8, col9, col10, col11, col12, _
    '                                             col13, col14, col15, _
    '                                             EOT, Me.CalculateCRC(1), Me.CalculateCRC(2))

    '    'Me.myMiniComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend("010080001700001E008080FE8080008080FE8080008080FE8080008080FE8080008080FE8080040000"))

    '    Me.myMiniComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend(MessageToSend))
    '    If Not Me.myMiniComPort.ErrorMessage = "" Then Me.ShowError(True)

    '    ''

    'End Sub

    Private Sub btnRefreshResults_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefreshResults.Click
        If (Not myCommSvr.m_ApplicationBussy) Then
            Try
                myCommSvr.m_ApplicationBussy = True

                Dim Race As Integer = Val(txtRaceNumber.Text)
                If ((Race > 0) And (Race <= myCommSvr.p_intMaxNumbOfRaces)) Then
                    myCommSvr.UpdateOfficialResults(myCommSvr.objResultFN(Race), Race)
                End If
                myCommSvr.m_ApplicationBussy = False
            Catch
                myCommSvr.m_ApplicationBussy = False
            End Try
        End If
    End Sub

    Private Sub txtOfficial_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOfficial.TextChanged
        Me.PrepareStatusDataToSend()
    End Sub

    Private Sub btnRefreshApp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefreshApp.Click
        If myCommSvr.m_ApplicationBussy Then
            Application.DoEvents()
        End If
        myCommSvr.m_ApplicationBussy = False
        tmrWPSPools.Enabled = True
    End Sub

    Private Sub ClearManuallyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearManuallyToolStripMenuItem.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        PrepareTimeOfDayToSend()
        PrepareWPSPoolHeaderToSend(True)
        PrepareTrackCondition(True)
        PrepareStatusDataToSend()


        'OUTPUT DATA
        'Me.myComPort.Output(Me.myMiniColMessages(i))
        ''Me.myComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend("01008000120120002600240121040000"))
        'Me.myComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend("011A010701090000000030370035040000"))
        'Me.myComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend("011A010701090001000030370035040000"))
        'Me.myComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend("011A010701090000000000000000040000"))

        Try


            'Me.myComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend("A56832FF7B00060000000900000100012502AE"))
            'Me.myComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend("A568320174111D77B9036414231457454C434F4D4532AE"))
            'Me.myComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend("A568320174111D77B9036414231457454C434F4D4532AE"))

            'Me.myComPort.Output("¥h2twÝ ¬@hello ¬®")
            'Me.myComPort.Output("A56832FF7B00060000000900000100012502AE")
            'Me.myComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend("a56832017411000000001b007700dd0300000000008000401200000000000003000568656c6c6f8005ae"))

            'Me.myComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend("a56832017411000000001b007700fb0300000000008000403000000000000003000548454c4c4f1c05ae"))
            'PrepareWPSPoolHeaderToSend(True)
            Dim stringtosend As String = Me.PrepareDataThreeColorSign("WIN", Color.Yellow, 24)
            Me.myComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend(stringtosend.Replace(" ", "")))

            '"a5 68 32 01 74 11 00 00 00 00 27 00 77 00 ee 03 00 00 00 00 00 80 00 40 17 00 00 00 00 00 00 03 00 11 6c 61 20 63 61 73 61 20 65 73 20 62 6f 6e 69 74 61 b4 09 ae"
            'Dim a As String = "a5 68 32 01 74 11 00 00 00 00 1b 00 77 00 db 03 00 00 00 00 00 80 00 40 10 00 00 00 00 00 00 03 00 05 68 65 6c 6c 6f 7c 05 ae"
            'a = "a5 68 32 01 74 11 00 00 00 00 1b 00 77 00 db 03 00 00 00 00 00 80 00 40 10 00 00 00 00 00 00 03 00 05 68 65 6c 6c 6f 84 ae"
            '6832017411000000001b007700db0300000000008000401000000000000003000568656c6c6f
            'Crc16.ComputeChecksum(Encoding.UTF8.GetBytes("Some string"))
            'Dim s As Byte() = System.Text.Encoding.UTF8.GetBytes("6832017411000000001b007700db0300000000008000401000000000000003000568656c6c6f")
            'a = a.Replace(" ", "")
            'a = "a5 68 32 01 74 11 00 00 00 00 1b 00 77 00 db 03 00 00 00 00 00 80 00 40 10 00 00 00 00 00 00 03 00 05 68 65 6c 6c 6f 7c 05 ae"
            'a = "a5 68 32 01 74 11 00 00 00 00 1b 00 77 00 93 03 00 00 00 00 00 64 00 14 10 00 00 00 00 00 00 03 00 05 68 65 6c 6c 6f ec 04 ae"
            'Me.myComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend(a.Replace(" ", "")))
        Catch ex As Exception

        End Try
        Me.myComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend("011A010001150000000A37353735373537353735273537353735040000"))




        Me.myComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend("011A010001150000000A37353735373537353735273537353735040000"))
        Me.myComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend("011A010001150000000040404040404040404040404040404040040000"))
        Me.myComPort.Output(Me.myCommSvr.oCommServerNet.GetMessageToSend("010080001700001E008080FE8080008080FE8080008080FE8080008080FE8080008080FE8080040000"))


    End Sub
End Class