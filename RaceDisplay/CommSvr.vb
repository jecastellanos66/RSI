Option Explicit On

Imports System.Runtime.InteropServices
Imports RSIData

'purpose of this class is to allow the user to update the information on the text
'boxes from the from the commserver object.  This class will take care of the destruction of
'all objects created by the commserver by calling the closeall procedure in the commserver object

Public Class CommSvr

    Private Const STX As Char = Chr(2)
    Private Const ETX As Char = Chr(3)

#Region " Events "
    '
    Public Delegate Sub CommEventMessage(ByVal strMessage As String)
    Public Event CommEventMsg As CommEventMessage
    Public Event CommEventTimerMsg As CommEventMessage
    '
    Public m_blnClearExotics As Boolean
    '
    Public p_strToteTrackCondition As String = ""
    Public p_strCurrentStatusJM As String
    Public p_CurrentRunnersStatusList As String
    Public p_intCurrentRace As Integer
    Public p_strCurrentMTP As String
    Public p_intMaxNumbOfRaces As Integer = 30
    Public p_intMaxFinisherPositions As Integer = 4

    Public Structure typFinishRunner
        Public strRunner As String
        Public intRunner As Integer
        Public blnDH As Boolean
        Public intFinishPosition As Integer
    End Structure

    Public p_udtWinners(p_intMaxNumbOfRaces + 1, p_intMaxFinisherPositions + 1) As typFinishRunner

    Private m_objWINPool(p_intMaxNumbOfRaces + 1) As clsRunnerTotals
    Private m_objPlacePool(p_intMaxNumbOfRaces + 1) As clsRunnerTotals
    Private m_objShowPool(p_intMaxNumbOfRaces + 1) As clsRunnerTotals
    Private p_blnWINPool As Boolean = True
    Private p_blnPlacePool As Boolean = True
    Private p_blnShowPool As Boolean = True
    Private p_blnSkipNewOfficialStatusEvent As Boolean = True

    Public objResultFN(p_intMaxNumbOfRaces + 1) As clsFinisherData
    Private objResultWin(p_intMaxNumbOfRaces + 1) As clsResultWPS
    Private objResultPlc(p_intMaxNumbOfRaces + 1) As clsResultWPS
    Private objResultShow(p_intMaxNumbOfRaces + 1) As clsResultWPS
    Public p_blnTmrResultsBusy As Boolean = False
    Public p_intRaceResultsOut As Integer = 0
    Public p_blnResultsOut As Boolean = False
    Public p_intClearResults As Integer = 1
    '
    Public Delegate Sub CommEventObjects(ByVal shrRaceNumber As Short)
    Public Event CommEventNewOdds As CommEventObjects
    Public Event CommEventRaceChange As CommEventObjects
    Public Event CommEventNewRO As CommEventObjects
    Public Event CommEventNewMTP As CommEventObjects
    Public Event CommEventNewTOD As CommEventObjects
    Public Event CommEventTrackConditionChange As CommEventObjects
    Public Event CommEventRaceStatusChange As CommEventObjects
    '
    Public Event DisplayNewOdds()
    Public Event DisplayNewRace()
    Public Event DisplayMTP()
    Public Event DisplayPostTime()
    Public Event DisplayTOD()
    Public Event DisplayTrackCondition()
    Public Event DisplayRaceStatus()
    Public Event DisplayRaceStatusJM()
    Public Event DisplayTeletimer()
    Public Event DisplayRunningOrder()
    Public Event DisplayRunningOrderAmtote()
    Public Event DisplayNewResults()
    Public Event DisplayMessages()
    Public Event DisplayTiming()
    'Public Event DisplayExotics()
    Public Event DisplayJudgesMessage(ByVal jmData As JMData)

    Public myNumberExacta As Integer
    Public myNumberTrifecta As Integer

    Public m_ApplicationBussy As Boolean = False
    '
#End Region

#Region " Declarations "

    Public WithEvents oCommServerNet As RSIPortNet.clsCommSvr

    'Public WithEvents oCommServer As RSIPort.clsCommSvr

    Public tmrWPS As Timer
    Private myDataset As RaceDisplayDataset
    Private myCurrentMessage As String
    Private myArrOdds(15) As Integer

#End Region

#Region " Enums "

    Private Enum WPSType
        WIN = 1
        PLC = 2
        SHW = 3
    End Enum

    Private Enum ExoticType
        DD = 1
        Perfecta = 2
        BET3 = 3
        Trifecta = 4
        Max = 5
    End Enum

    Private intCurrPage(ExoticType.Max) As Integer

#End Region

#Region " Properties "

    Public ReadOnly Property CurrentMessage() As String
        Get
            Return Me.myCurrentMessage
        End Get
    End Property

    Public ReadOnly Property Dataset() As RaceDisplayDataset
        Get
            Return Me.myDataset
        End Get
    End Property

    Public ReadOnly Property ArrayOdds() As System.Array
        Get
            Return Me.myArrOdds
        End Get
    End Property

#End Region

#Region " Constructor "

    Public Sub New(skipNewOfficialStatusEvent As Boolean)
        p_blnSkipNewOfficialStatusEvent = skipNewOfficialStatusEvent
        Try
            For intRaceNum As Integer = 1 To p_intMaxNumbOfRaces
                For intCtrFor As Integer = 1 To p_intMaxFinisherPositions
                    p_udtWinners(intRaceNum, intCtrFor).strRunner = ""
                    p_udtWinners(intRaceNum, intCtrFor).intRunner = 0
                    p_udtWinners(intRaceNum, intCtrFor).blnDH = False
                    p_udtWinners(intRaceNum, intCtrFor).intFinishPosition = 0
                Next
            Next

            For intRaceNum As Integer = 0 To p_intMaxNumbOfRaces
                m_objWINPool(intRaceNum) = New clsRunnerTotals()
                m_objPlacePool(intRaceNum) = New clsRunnerTotals()
                m_objShowPool(intRaceNum) = New clsRunnerTotals()

                objResultWin(intRaceNum) = New clsResultWPS()
                objResultPlc(intRaceNum) = New clsResultWPS()
                objResultShow(intRaceNum) = New clsResultWPS()
            Next

            For i As Integer = 1 To ExoticType.Max
                intCurrPage(i) = 1
            Next
            Me.myDataset = New RaceDisplayDataset

            oCommServerNet = New RSIPortNet.clsCommSvr()
            If oCommServerNet Is Nothing Then Return

            'oCommServer = New RSIPort.clsCommSvr()
            'If oCommServer Is Nothing Then Return
            '
            'AddHandler oCommServer.Message, AddressOf oCommServer_Message
            'AddHandler oCommServer.TimerMsg, AddressOf oCommServer_TimerMsg
            'AddHandler oCommServer.NewOdds, AddressOf oCommServer_NewOdds
            'AddHandler oCommServer.RaceChange, AddressOf oCommServer_RaceChange
            'AddHandler oCommServer.NewRO, AddressOf oCommServer_NewRO
        Catch ex As Exception
            Throw New Exception("Error creating CommSvr class" & vbCrLf & ex.Message)

        End Try
        ''
    End Sub

#End Region

#Region " Destructor "

    Protected Overrides Sub Finalize()
        Try
            'Do While Marshal.ReleaseComObject(oCommServer) > 0
            'Loop
            'oCommServer.CloseAll()
            'oCommServer = Nothing

            'Do While Marshal.ReleaseComObject(oCommServerNet) > 0
            'Loop

            oCommServerNet.CloseAll()
            oCommServerNet = Nothing

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error finalizing CommSvr " + ex.ToString())
            'nothing
        End Try
    End Sub


#End Region

    Public Sub CancelProcess()
        If Not oCommServerNet Is Nothing Then
            oCommServerNet.CancelProcess = True
        End If
    End Sub

#Region " Events raised by comm server "

    'Private Sub oCommServer_Message(ByRef strMessage As String)
    '    'when the event for the tote comm object is raised then
    '    'we write to the text boxes
    '    RaiseEvent CommEventMsg(strMessage)
    'End Sub

    Private Sub oCommServerNet_Message(ByRef strMessage As String) Handles oCommServerNet.Message
        Try
            Me.myCurrentMessage = strMessage
            If IsJudgeMessage(strMessage) Then
                ProcessJudgesMessage(strMessage)
            End If
            'RaiseEvent DisplayMessages()
        Catch
            'nothing
        End Try
    End Sub

    'Private Sub oCommServer_Message(ByRef strMessage As String) Handles oCommServer.Message
    '    Try
    '        Me.myCurrentMessage = strMessage
    '        If IsJudgeMessage(strMessage) Then
    '            ProcessJudgesMessage(strMessage)
    '        End If
    '        'RaiseEvent DisplayMessages()
    '    Catch
    '        'nothing
    '    End Try
    'End Sub

    Private Sub ProcessJudgesMessage(message As String)
        Dim data As JMData = ProcessJM(message)
        RaiseEvent DisplayJudgesMessage(data)
    End Sub

    Private Function ProcessJM(ByVal message As String) As JMData
        Dim strTemp As String
        Dim intInfoPointer As Integer
        Dim intCounter As Integer
        Dim data As JMData

        On Error GoTo ErrHndlr

        '"JMS00S SARATOGA SUMMER          0159RACE090090OFFCOBCINCPHCDE"

        message = message.TrimStart(STX).TrimEnd(ETX)
        data = New JMData()

        data.TrackMeetID = Mid(message, 47, 3)

        data.RaceNumber = Mid(message, 58, 2)

        intInfoPointer = 66

        strTemp = Mid(message, intInfoPointer, 3)
        data.Official = UCase(strTemp) = "OFF"
        intInfoPointer = intInfoPointer + 3

        strTemp = Mid(message, intInfoPointer, 3)
        data.Objection = UCase(strTemp) = "OBJ"
        intInfoPointer = intInfoPointer + 3

        data.SetObjectionRunners(Mid(message, intInfoPointer, 15))
        intInfoPointer = intInfoPointer + 15


        strTemp = Mid(message, intInfoPointer, 3)
        data.Inquiry = UCase(strTemp) = "INQ"
        intInfoPointer = intInfoPointer + 3

        data.SetInquiryRunners(Mid(message, intInfoPointer, 15))
        intInfoPointer = intInfoPointer + 15

        strTemp = Mid(message, intInfoPointer, 3)
        data.Photo = UCase(strTemp) = "PHT"
        intInfoPointer = intInfoPointer + 3

        strTemp = Mid(message, intInfoPointer, 3)
        data.Deaheat = UCase(strTemp) = "DEA"
        intInfoPointer = intInfoPointer + 3

        'From protocol document...
        '   <',’> indicates dead heat with next.
        '   <'.’> indicates next position.
        '   <' ‘> indicates nothing follows.

        Dim NumbDH As Integer = CInt(Mid(message, intInfoPointer, 1))
        intInfoPointer = intInfoPointer + 1

        Dim numOfDhRunners As Integer = 4
        data.DeaheatRunners = New String(numOfDhRunners) {}
        If (NumbDH > 0) Then
            Dim runnerIdx As Integer = 0
            For intCounter = 1 To NumbDH
                Dim runnerData As String = Mid(message, intInfoPointer, 4)

                If (runnerData.EndsWith(",")) Then
                    'Found dead hit with next.. get runner
                    AddDeadHeatRunner(numOfDhRunners, Mid(message, intInfoPointer, 3).Trim(), runnerIdx, data.DeaheatRunners)

                    'Get next runner 
                    intInfoPointer = intInfoPointer + 4
                    AddDeadHeatRunner(numOfDhRunners, Mid(message, intInfoPointer, 3).Trim(), runnerIdx, data.DeaheatRunners)
                ElseIf (runnerData.EndsWith(" ")) Then
                    Exit For
                End If

                intInfoPointer = intInfoPointer + 4
            Next intCounter
        End If

        ProcessJM = data

        Exit Function

ErrHndlr:
        'Implement error logging
    End Function

    Private Sub AddDeadHeatRunner(numOfDhRunners As Integer, dhRunner As String, ByRef runnerIdx As Integer, ByRef dhRunners As String())
        If runnerIdx < numOfDhRunners Then
            dhRunners(runnerIdx) = dhRunner
            runnerIdx = runnerIdx + 1
        End If
    End Sub

    Private Function IsJudgeMessage(message As String) As Boolean

        Dim msgType As String = UCase(Mid(message, 11, 2))
        Return msgType = "JM"

    End Function


    'Private Sub oCommServer_RaceHeaderChange(ByRef CurrentMTP As String, ByRef CurrentTime As String, ByRef CurrentPostTime As String, ByRef CurrentRace As Short) Handles oCommServer.RaceHeaderChange
    '    Dim Race As Integer = Convert.ToInt32(CurrentRace)
    '    Dim strMtp As String = ""
    '    'Ignore strMtp --> They are behind
    '    Try
    '        If m_ApplicationBussy Then
    '            Application.DoEvents()
    '        End If
    '        m_ApplicationBussy = True
    '        p_intCurrentRace = Race
    '        Me.UpdateRace(Race)
    '        'set mtp
    '        If ((IsDate(CurrentTime)) And (IsDate(CurrentPostTime))) Then
    '            Dim startTime As DateTime = DateTime.Parse(CurrentTime)
    '            Dim endTime As DateTime = DateTime.Parse(CurrentPostTime)
    '            Dim varTime As TimeSpan

    '            varTime = endTime.Subtract(startTime)

    '            Dim fractionalMinutes As Double = varTime.TotalMinutes
    '            Dim wholeMinutes As Integer = CInt(fractionalMinutes)

    '            strMtp = wholeMinutes.ToString()

    '            If (wholeMinutes < 0) Then
    '                strMtp = 0
    '            End If
    '        Else
    '            strMtp = "   "
    '        End If
    '        p_strCurrentMTP = strMtp
    '        Me.UpdateMTP(CurrentPostTime, strMtp, Race)

    '        'p_strCurrentMTP = CurrentMTP
    '        'Me.UpdateMTP(CurrentPostTime, CurrentMTP, Race)
    '        'time of day tod
    '        Me.UpdateTOD(CurrentTime, Race)
    '        m_ApplicationBussy = False
    '    Catch
    '        m_ApplicationBussy = False
    '    End Try
    'End Sub

    Private Sub oCommServerNet_RaceHeaderChange(ByRef CurrentMTP As String, ByRef CurrentTime As String, ByRef CurrentPostTime As String, ByRef CurrentRace As Short) Handles oCommServerNet.RaceHeaderChange
        Dim Race As Integer = Convert.ToInt32(CurrentRace)
        Dim strMtp As String = ""
        'Ignore strMtp --> They are behind
        Try
            If m_ApplicationBussy Then
                Application.DoEvents()
            End If
            m_ApplicationBussy = True
            p_intCurrentRace = Race
            Me.UpdateRace(Race)
            'set mtp
            If ((IsDate(CurrentTime)) And (IsDate(CurrentPostTime))) Then
                Dim startTime As DateTime = DateTime.Parse(CurrentTime)
                Dim endTime As DateTime = DateTime.Parse(CurrentPostTime)
                Dim varTime As TimeSpan

                varTime = endTime.Subtract(startTime)

                Dim fractionalMinutes As Double = varTime.TotalMinutes
                Dim wholeMinutes As Integer = CInt(fractionalMinutes)

                strMtp = wholeMinutes.ToString()

                If (wholeMinutes < 0) Then
                    strMtp = 0
                End If
            Else
                strMtp = "   "
            End If
            p_strCurrentMTP = strMtp
            Me.UpdateMTP(CurrentPostTime, strMtp, Race)

            'p_strCurrentMTP = CurrentMTP
            'Me.UpdateMTP(CurrentPostTime, CurrentMTP, Race)
            'time of day tod
            Me.UpdateTOD(CurrentTime, Race)
            m_ApplicationBussy = False
        Catch
            m_ApplicationBussy = False
        End Try
    End Sub

    'Private Sub oCommServer_TrackConditionChange(ByRef CurrentTrackCondition As String, ByRef intRace As Short) Handles oCommServer.TrackConditionChange
    '    Try
    '        If m_ApplicationBussy Then
    '            Application.DoEvents()
    '        End If
    '        m_ApplicationBussy = True
    '        Me.UpdateTrackCondition(CurrentTrackCondition)
    '        m_ApplicationBussy = False
    '    Catch
    '        m_ApplicationBussy = False
    '    End Try
    'End Sub

    Private Sub oCommServerNet_TrackConditionChange(ByRef CurrentTrackCondition As String, ByRef intRace As Short) Handles oCommServerNet.TrackConditionChange
        Try
            If m_ApplicationBussy Then
                Application.DoEvents()
            End If
            m_ApplicationBussy = True
            Me.UpdateTrackCondition(CurrentTrackCondition)
            m_ApplicationBussy = False
        Catch
            m_ApplicationBussy = False
        End Try
    End Sub

    'United
    'Private Sub oCommServer_RaceStatusChange(ByRef intRace As Short) Handles oCommServer.RaceStatusChange
    '    Me.UpdateRaceStatus()
    'End Sub

    'Private Sub oCommServer_NewOdds(ByRef m_objOdds As clsOdds, ByRef intRace As Short) Handles oCommServer.NewOdds
    '    Dim Race As Integer = Convert.ToInt32(intRace)
    '    Try
    '        If m_ApplicationBussy Then
    '            Application.DoEvents()
    '        End If
    '        m_ApplicationBussy = True
    '        'always update odds
    '        Me.UpdateOdds(m_objOdds, Race)
    '        m_ApplicationBussy = False
    '    Catch
    '        m_ApplicationBussy = False
    '    End Try
    'End Sub

    Private Sub oCommServerNet_NewOdds(ByRef m_objOdds As clsOdds, ByRef intRace As Short) Handles oCommServerNet.NewOdds
        Dim Race As Integer = Convert.ToInt32(intRace)
        Try
            If m_ApplicationBussy Then
                Application.DoEvents()
            End If
            m_ApplicationBussy = True
            'always update odds
            Me.UpdateOdds(m_objOdds, Race)
            m_ApplicationBussy = False
        Catch
            m_ApplicationBussy = False
        End Try
    End Sub

    'Private Sub oCommServer_NewRO(ByRef m_objRunningOrder As clsRunningOrder, ByRef intRace As Short) Handles oCommServer.NewRO
    '    'Private Sub oCommServer_NewRO(ByRef m_objRunningOrder As clsRunningOrder, ByRef strToteCompany as string, ByRef intRace As Short) Handles oCommServer.NewRO
    '    Dim Race As Integer = Convert.ToInt32(intRace)
    '    Try
    '        If m_ApplicationBussy Then
    '            Application.DoEvents()
    '        End If
    '        m_ApplicationBussy = True
    '        Dim strToteCompany As String = ""
    '        If (strToteCompany.ToUpper() = "AMTOTE") Then
    '            Me.UpdateRunningOrderAmtote(m_objRunningOrder, Race)
    '        Else
    '            'If (Not p_blnResultsOut) Then
    '            '    If (Val(p_strCurrentMTP) <= 0) Then
    '            Me.UpdateRunningOrder(m_objRunningOrder, Race)
    '            '    End If
    '            'End If
    '        End If
    '        m_ApplicationBussy = False
    '    Catch
    '        m_ApplicationBussy = False
    '    End Try
    'End Sub

    Private Sub oCommServer_NewRO(ByRef m_objRunningOrder As clsRunningOrder, ByRef intRace As Short) Handles oCommServerNet.NewRO
        'Private Sub oCommServer_NewRO(ByRef m_objRunningOrder As clsRunningOrder, ByRef strToteCompany as string, ByRef intRace As Short) Handles oCommServer.NewRO
        Dim Race As Integer = Convert.ToInt32(intRace)
        Try
            If m_ApplicationBussy Then
                Application.DoEvents()
            End If
            m_ApplicationBussy = True
            Dim strToteCompany As String = ""
            If (strToteCompany.ToUpper() = "AMTOTE") Then
                Me.UpdateRunningOrderAmtote(m_objRunningOrder, Race)
            Else
                'If (Not p_blnResultsOut) Then
                '    If (Val(p_strCurrentMTP) <= 0) Then
                Me.UpdateRunningOrder(m_objRunningOrder, Race)
                '    End If
                'End If
            End If
            m_ApplicationBussy = False
        Catch
            m_ApplicationBussy = False
        End Try
    End Sub

    'Private Sub oCommServer_NewOfficialStatus(ByRef blnOfficialStatus As Boolean, ByRef blnObjStatus As Boolean, ByRef blnInqStatus As Boolean, ByRef blnPhotoStatus As Boolean, ByRef blnDHStatus As Boolean, ByRef intRace As Short) Handles oCommServer.NewOfficialStatus

    '    If p_blnSkipNewOfficialStatusEvent Then
    '        Exit Sub
    '    End If

    '    Dim Race As Integer = Convert.ToInt32(intRace)
    '    Try
    '        If m_ApplicationBussy Then
    '            Application.DoEvents()
    '        End If
    '        m_ApplicationBussy = True
    '        Me.UpdateStatusJM(blnOfficialStatus, blnObjStatus, blnInqStatus, blnPhotoStatus, blnDHStatus, Race)
    '        m_ApplicationBussy = False
    '    Catch
    '        m_ApplicationBussy = False
    '    End Try
    'End Sub

    Public Event NewOfficialStatus(ByRef m_udtJudgesInfo As RSIData.typJudgesInfo, ByRef CurrentOfficialStatus As String, ByRef intRace As Short)
    Private Sub oCommServerNet_NewOfficialStatus(ByRef judgesInfo As RSIData.typJudgesInfo, ByRef CurrentOfficialStatus As String, ByRef intRace As Short) Handles oCommServerNet.NewOfficialStatus

        If p_blnSkipNewOfficialStatusEvent Then
            Exit Sub
        End If

        Dim Race As Integer = Convert.ToInt32(intRace)
        Try
            If m_ApplicationBussy Then
                Application.DoEvents()
            End If
            m_ApplicationBussy = True
            Me.UpdateStatusJM(judgesInfo.blnOfficialStatus, judgesInfo.blnObjStatus, judgesInfo.blnInqStatus, judgesInfo.blnPhotoStatus, judgesInfo.blnDHStatus, Race)
            m_ApplicationBussy = False
        Catch
            m_ApplicationBussy = False
        End Try
    End Sub

    'Private Sub oCommServer_NewTT(ByRef m_objTeleTimer As clsTeleTimer, ByRef intRace As Short) Handles oCommServer.NewTT
    '    Dim Race As Integer = Convert.ToInt32(intRace)
    '    Try
    '        If m_ApplicationBussy Then
    '            Application.DoEvents()
    '        End If
    '        m_ApplicationBussy = True
    '        Me.UpdateTeletimer(m_objTeleTimer, Race)
    '        m_ApplicationBussy = False
    '    Catch
    '        m_ApplicationBussy = False
    '    End Try
    'End Sub

    Private Sub oCommServerNet_NewTT(ByRef m_objTeleTimer As clsTeleTimer, ByRef intRace As Short) Handles oCommServerNet.NewTT
        Dim Race As Integer = Convert.ToInt32(intRace)
        Try
            If m_ApplicationBussy Then
                Application.DoEvents()
            End If
            m_ApplicationBussy = True
            Me.UpdateTeletimer(m_objTeleTimer, Race)
            m_ApplicationBussy = False
        Catch
            m_ApplicationBussy = False
        End Try
    End Sub

    'Private Sub oCommServer_NewResults(ByRef m_objFinisherData As clsFinisherData, ByRef intRace As Short) Handles oCommServer.NewResults
    '    Dim Race As Integer = Convert.ToInt32(intRace)
    '    Try
    '        If m_ApplicationBussy Then
    '            Application.DoEvents()
    '        End If
    '        m_ApplicationBussy = True
    '        p_intRaceResultsOut = Race
    '        p_blnResultsOut = True
    '        p_intClearResults = 1
    '        If ((Race > 0) And (Race <= p_intMaxNumbOfRaces)) Then
    '            objResultFN(Race) = m_objFinisherData
    '        End If
    '        Me.UpdateOfficialResults(m_objFinisherData, intRace)
    '        m_ApplicationBussy = False
    '    Catch
    '        m_ApplicationBussy = False
    '    End Try
    'End Sub

    Private Sub oCommServerNet_NewResults(ByRef m_objFinisherData As clsFinisherData, ByRef intRace As Short) Handles oCommServerNet.NewResults
        Dim Race As Integer = Convert.ToInt32(intRace)
        Try
            If m_ApplicationBussy Then
                Application.DoEvents()
            End If
            m_ApplicationBussy = True
            p_intRaceResultsOut = Race
            p_blnResultsOut = True
            p_intClearResults = 1
            If ((Race > 0) And (Race <= p_intMaxNumbOfRaces)) Then
                objResultFN(Race) = m_objFinisherData
            End If
            Me.UpdateOfficialResults(m_objFinisherData, intRace)
            m_ApplicationBussy = False
        Catch
            m_ApplicationBussy = False
        End Try
    End Sub

    'Private Sub oCommServer_NewResultsWIN(ByRef m_objResultWPS As clsResultWPS, ByRef intRace As Short) Handles oCommServer.NewResultsWIN
    '    Dim Race As Integer = Convert.ToInt32(intRace)
    '    Dim blnFlag As Boolean
    '    Try
    '        'blnFlag = frmMain.timerOfficial.Enabled
    '        'frmMain.timerOfficial.Enabled = False
    '        If (p_blnTmrResultsBusy) Then
    '            Application.DoEvents()
    '        End If
    '        If ((Race > 0) And (Race <= p_intMaxNumbOfRaces)) Then
    '            If (Char.IsUpper(m_objResultWPS.Status, 0)) Then
    '                objResultWin(Race) = m_objResultWPS
    '            End If
    '        End If
    '        'frmMain.timerOfficial.Enabled = blnFlag
    '    Catch ex As Exception
    '        'frmMain.timerOfficial.Enabled = blnFlag
    '    End Try
    'End Sub

    Private Sub oCommServerNet_NewResultsWIN(ByRef m_objResultWPS As clsResultWPS, ByRef intRace As Short) Handles oCommServerNet.NewResultsWIN
        Dim Race As Integer = Convert.ToInt32(intRace)
        Dim blnFlag As Boolean
        Try
            'blnFlag = frmMain.timerOfficial.Enabled
            'frmMain.timerOfficial.Enabled = False
            If (p_blnTmrResultsBusy) Then
                Application.DoEvents()
            End If
            If ((Race > 0) And (Race <= p_intMaxNumbOfRaces)) Then
                If (Char.IsUpper(m_objResultWPS.Status, 0)) Then
                    objResultWin(Race) = m_objResultWPS
                End If
            End If
            'frmMain.timerOfficial.Enabled = blnFlag
        Catch ex As Exception
            'frmMain.timerOfficial.Enabled = blnFlag
        End Try
    End Sub

    'Private Sub oCommServer_NewResultsPLC(ByRef m_objResultWPS As clsResultWPS, ByRef intRace As Short) Handles oCommServer.NewResultsPLC
    '    Dim Race As Integer = Convert.ToInt32(intRace)
    '    Dim blnFlag As Boolean
    '    Try
    '        'blnFlag = frmMain.timerOfficial.Enabled
    '        'frmMain.timerOfficial.Enabled = False
    '        If (p_blnTmrResultsBusy) Then
    '            Application.DoEvents()
    '        End If
    '        If ((Race > 0) And (Race <= p_intMaxNumbOfRaces)) Then
    '            If (Char.IsUpper(m_objResultWPS.Status, 0)) Then
    '                objResultPlc(Race) = m_objResultWPS
    '            End If
    '        End If
    '        'frmMain.timerOfficial.Enabled = blnFlag
    '    Catch ex As Exception
    '        'frmMain.timerOfficial.Enabled = blnFlag
    '    End Try
    'End Sub

    Private Sub oCommServerNet_NewResultsPLC(ByRef m_objResultWPS As clsResultWPS, ByRef intRace As Short) Handles oCommServerNet.NewResultsPLC
        Dim Race As Integer = Convert.ToInt32(intRace)
        Dim blnFlag As Boolean
        Try
            'blnFlag = frmMain.timerOfficial.Enabled
            'frmMain.timerOfficial.Enabled = False
            If (p_blnTmrResultsBusy) Then
                Application.DoEvents()
            End If
            If ((Race > 0) And (Race <= p_intMaxNumbOfRaces)) Then
                If (Char.IsUpper(m_objResultWPS.Status, 0)) Then
                    objResultPlc(Race) = m_objResultWPS
                End If
            End If
            'frmMain.timerOfficial.Enabled = blnFlag
        Catch ex As Exception
            'frmMain.timerOfficial.Enabled = blnFlag
        End Try
    End Sub

    'Private Sub oCommServer_NewResultsSHW(ByRef m_objResultWPS As clsResultWPS, ByRef intRace As Short) Handles oCommServer.NewResultsSHW
    '    Dim Race As Integer = Convert.ToInt32(intRace)
    '    Dim blnFlag As Boolean
    '    Try
    '        'blnFlag = frmMain.timerOfficial.Enabled
    '        'frmMain.timerOfficial.Enabled = False
    '        If (p_blnTmrResultsBusy) Then
    '            Application.DoEvents()
    '        End If
    '        If ((Race > 0) And (Race <= p_intMaxNumbOfRaces)) Then
    '            If (Char.IsUpper(m_objResultWPS.Status, 0)) Then
    '                objResultShow(Race) = m_objResultWPS
    '            End If
    '        End If
    '        'frmMain.timerOfficial.Enabled = blnFlag
    '    Catch ex As Exception
    '        'frmMain.timerOfficial.Enabled = blnFlag
    '    End Try
    'End Sub

    Private Sub oCommServerNet_NewResultsSHW(ByRef m_objResultWPS As clsResultWPS, ByRef intRace As Short) Handles oCommServerNet.NewResultsSHW
        Dim Race As Integer = Convert.ToInt32(intRace)
        Dim blnFlag As Boolean
        Try
            'blnFlag = frmMain.timerOfficial.Enabled
            'frmMain.timerOfficial.Enabled = False
            If (p_blnTmrResultsBusy) Then
                Application.DoEvents()
            End If
            If ((Race > 0) And (Race <= p_intMaxNumbOfRaces)) Then
                If (Char.IsUpper(m_objResultWPS.Status, 0)) Then
                    objResultShow(Race) = m_objResultWPS
                End If
            End If
            'frmMain.timerOfficial.Enabled = blnFlag
        Catch ex As Exception
            'frmMain.timerOfficial.Enabled = blnFlag
        End Try
    End Sub

    'No trabaja porque no tiene timming en el CommSvr, CommPort 4
    'Private Sub oCommServer_TimerMsg(ByRef strTimer_Msg As String) Handles oCommServer.TimerMsg
    '    If (Not m_ApplicationBussy) Then
    '        Try
    '            m_ApplicationBussy = True
    '            Me.UpdateTiming()
    '            m_ApplicationBussy = False
    '        Catch
    '            m_ApplicationBussy = False
    '        End Try
    '    End If
    'End Sub

    'Private Sub oCommServer_NewWINPool(ByRef m_objRunnerTotals As clsRunnerTotals, ByRef intRace As Short) Handles oCommServer.NewWINPool
    '    Dim Race As Integer = Convert.ToInt32(intRace)
    '    Try
    '        p_blnWINPool = False
    '        Application.DoEvents()
    '        If ((Race > 0) And (Race <= p_intMaxNumbOfRaces)) Then
    '            m_objWINPool(Race) = m_objRunnerTotals
    '        End If
    '        p_blnWINPool = True
    '        '
    '    Catch ex As Exception
    '        p_blnWINPool = True
    '    End Try
    'End Sub

    Private Sub oCommServerNet_NewWINPool(ByRef m_objRunnerTotals As clsRunnerTotals, ByRef intRace As Short) Handles oCommServerNet.NewWINPool
        Dim Race As Integer = Convert.ToInt32(intRace)
        Try
            p_blnWINPool = False
            Application.DoEvents()
            If ((Race > 0) And (Race <= p_intMaxNumbOfRaces)) Then
                m_objWINPool(Race) = m_objRunnerTotals
            End If
            p_blnWINPool = True
            '
        Catch ex As Exception
            p_blnWINPool = True
        End Try
    End Sub

    'Private Sub oCommServer_NewPlacePool(ByRef m_objRunnerTotals As clsRunnerTotals, ByRef intRace As Short) Handles oCommServer.NewPlacePool
    '    Dim Race As Integer = Convert.ToInt32(intRace)

    '    Try
    '        p_blnPlacePool = False
    '        Application.DoEvents()
    '        If ((Race > 0) And (Race <= p_intMaxNumbOfRaces)) Then
    '            m_objPlacePool(Race) = m_objRunnerTotals
    '        End If
    '        p_blnPlacePool = True
    '        '
    '    Catch ex As Exception
    '        p_blnPlacePool = True
    '    End Try
    'End Sub

    Private Sub oCommServerNet_NewPlacePool(ByRef m_objRunnerTotals As clsRunnerTotals, ByRef intRace As Short) Handles oCommServerNet.NewPlacePool
        Dim Race As Integer = Convert.ToInt32(intRace)

        Try
            p_blnPlacePool = False
            Application.DoEvents()
            If ((Race > 0) And (Race <= p_intMaxNumbOfRaces)) Then
                m_objPlacePool(Race) = m_objRunnerTotals
            End If
            p_blnPlacePool = True
            '
        Catch ex As Exception
            p_blnPlacePool = True
        End Try
    End Sub

    'Private Sub oCommServer_NewShowPool(ByRef m_objRunnerTotals As clsRunnerTotals, ByRef intRace As Short) Handles oCommServer.NewShowPool
    '    Dim Race As Integer = Convert.ToInt32(intRace)

    '    Try
    '        p_blnShowPool = False
    '        Application.DoEvents()
    '        If ((Race > 0) And (Race <= p_intMaxNumbOfRaces)) Then
    '            m_objShowPool(Race) = m_objRunnerTotals
    '        End If
    '        p_blnShowPool = True
    '        '
    '    Catch ex As Exception
    '        p_blnShowPool = True
    '    End Try
    'End Sub

    Private Sub oCommServerNet_NewShowPool(ByRef m_objRunnerTotals As clsRunnerTotals, ByRef intRace As Short) Handles oCommServerNet.NewShowPool
        Dim Race As Integer = Convert.ToInt32(intRace)

        Try
            p_blnShowPool = False
            Application.DoEvents()
            If ((Race > 0) And (Race <= p_intMaxNumbOfRaces)) Then
                m_objShowPool(Race) = m_objRunnerTotals
            End If
            p_blnShowPool = True
            '
        Catch ex As Exception
            p_blnShowPool = True
        End Try
    End Sub

#End Region

#Region " Update Methods "

    Private Sub UpdateMTP(ByVal strPT As String, ByVal strMTP As String, ByVal intRace As Integer)
        Me.myDataset.MTP.Clear()
        Me.myDataset.POSTTIME.Clear()

        If strMTP Is Nothing Then Return

        strMTP = Me.FixString(strMTP)
        '
        Dim drMTP As RaceDisplayDataset.MTPRow
        drMTP = Me.myDataset.MTP.NewMTPRow
        '
        drMTP.BeginEdit()
        drMTP.RACEKEY = intRace
        drMTP.MINUTESA = strMTP.Substring(0, 1)
        drMTP.MINUTESB = strMTP.Substring(1, 1)
        drMTP.EndEdit()
        '
        Me.myDataset.MTP.AddMTPRow(drMTP)
        '
        RaiseEvent DisplayMTP()
        '
        If strPT Is Nothing Then Return

        strPT = Me.FixString(strPT, 4)
        '
        Dim drPT As RaceDisplayDataset.POSTTIMERow
        drPT = Me.myDataset.POSTTIME.NewPOSTTIMERow
        '
        drPT.BeginEdit()
        drPT.RACEKEY = intRace
        drPT.PTA = strPT.Substring(0, 1).Replace("0", "")
        drPT.PTB = strPT.Substring(1, 1)
        drPT.PTC = strPT.Substring(2, 1)
        drPT.PTD = strPT.Substring(3, 1)
        drPT.EndEdit()
        '
        Me.myDataset.POSTTIME.AddPOSTTIMERow(drPT)
        '
        RaiseEvent DisplayPostTime()
    End Sub

    Private Sub UpdateTOD(ByVal strTOD As String, ByVal intRace As Integer)
        Me.myDataset.TOD.Clear()

        If strTOD Is Nothing Then Return
        '
        Try
            strTOD = Format(CDate(strTOD), "hh:mm")
        Catch ex As Exception
            '
        End Try

        '
        strTOD = Me.FixString(strTOD, 4)
        Dim drTOD As RaceDisplayDataset.TODRow
        drTOD = Me.myDataset.TOD.NewTODRow
        '
        drTOD.BeginEdit()
        drTOD.RACEKEY = intRace

        drTOD.TODA = strTOD.Substring(0, 1).Replace("0", "")
        drTOD.TODB = strTOD.Substring(1, 1)
        drTOD.TODC = strTOD.Substring(2, 1)
        drTOD.TODD = strTOD.Substring(3, 1)
        drTOD.EndEdit()
        '
        Me.myDataset.TOD.AddTODRow(drTOD)
        '
        RaiseEvent DisplayTOD()
        ''
    End Sub

    Private Sub UpdateTrackCondition(ByVal CurrentTrackCondition As String)
        p_strToteTrackCondition = CurrentTrackCondition
        RaiseEvent DisplayTrackCondition()
    End Sub

    Private Sub UpdateRaceStatus() 'United
        RaiseEvent DisplayRaceStatus()
    End Sub

    Private Sub UpdateStatusJM(ByVal blnOfficialStatus As Boolean, ByVal blnObjStatus As Boolean, ByVal blnInqStatus As Boolean, ByVal blnPhotoStatus As Boolean, ByVal blnDHStatus As Boolean, ByVal intRace As Integer) 'Sportech
        Try
            p_strCurrentStatusJM = ""
            p_CurrentRunnersStatusList = ""
            If (blnOfficialStatus) Then
                p_strCurrentStatusJM = p_strCurrentStatusJM + "F"
            End If
            If (blnInqStatus) Then
                p_strCurrentStatusJM = p_strCurrentStatusJM + "I"
                'p_CurrentRunnersStatusList = GetRunnerListWithIndicator(",", m_udtJudgesInfo.strInqRunnerList, 5, 3)
            End If
            If (blnObjStatus) Then
                p_strCurrentStatusJM = p_strCurrentStatusJM + "O"
                'p_CurrentRunnersStatusList = GetRunnerListWithIndicator(",", m_udtJudgesInfo.strObjRunnerList, 5, 3)
            End If
            If (blnDHStatus) Then
                p_strCurrentStatusJM = p_strCurrentStatusJM + "D"
                'p_CurrentRunnersStatusList = GetRunnerDHList(m_udtJudgesInfo.intNumbDHRunners, m_udtJudgesInfo.strDHRunnerList, 4)
            End If
            If (blnPhotoStatus) Then
                p_strCurrentStatusJM = p_strCurrentStatusJM + "P"
            End If
            '
            'm_blnClearExotics = False
            'If blnOfficialStatus = False Then
            'm_blnClearExotics = True
            'End If
            ''
        Catch ex As Exception
            '
        End Try
        RaiseEvent DisplayRaceStatusJM()
    End Sub

    Private Sub UpdateStatusJM(ByVal jmData As JMData)

    End Sub

    Private Function GetRunnerListWithIndicator(ByVal Indicator As String, ByVal strList As String, ByVal NumberOfRows As Integer, ByVal intLen As Integer) As String
        Dim Ctr As Integer = 1
        Dim c As String = ""
        Dim z As String = ""
        Try
            For i As Integer = 1 To (NumberOfRows)
                c = Mid(strList, Ctr, intLen)
                If c.Trim() <> "" Then
                    If (z.Trim = "") Then
                        z = c
                    Else
                        z = z + Indicator + c
                    End If
                End If
                Ctr = Ctr + intLen
            Next
            Return z
        Catch ex As Exception
            Return ""
        End Try

    End Function

    Private Function GetRunnerDHList(ByVal NumberOfRows As Integer, ByVal strList As String, ByVal intLen As Integer) As String
        Dim Ctr As Integer = 1
        Dim c As String = ""
        Dim z As String = ""
        Dim flag As Boolean = False
        Try
            For i As Integer = 1 To (NumberOfRows)
                c = Mid(strList, Ctr, intLen)
                If (Strings.Right(c, 1) = ",") Then
                    flag = True
                    If (z.Trim = "") Then
                        z = Mid(strList, Ctr, intLen - 1)
                    Else
                        z = z + "," + Mid(strList, Ctr, intLen - 1)
                    End If
                ElseIf (flag) Then
                    flag = False
                    z = z + "," + Mid(strList, Ctr, intLen - 1)
                End If
                Ctr = Ctr + intLen
            Next
            Return z
        Catch ex As Exception
            Return ""
        End Try

    End Function

    Public Sub UpdateWPSPools(ByVal intRace As Integer, ByVal strType As String)
        'this sub is public because i need to call it from frmMain's tmrWPSPools
        'Here we'll show the WIN PLC SHOW totals 
        Dim objWPSPools As New RSIData.clsRunnerTotals
        Dim s As String
        Dim x As Integer
        Dim IndexOfs As Integer
        Dim blnWPSFound As Boolean = False

        Me.myDataset.EXOTICS.Clear()
        Dim drWPSPools As RaceDisplayDataset.EXOTICSRow
        drWPSPools = Me.myDataset.EXOTICS.NewEXOTICSRow

        Dim processPoolsMsg As Boolean
        processPoolsMsg = (strType = "WIN" And m_objWINPool(intRace).NumberOfRows <= 12)
        processPoolsMsg = processPoolsMsg Or (strType = "PLC" And m_objPlacePool(intRace).NumberOfRows <= 12)
        processPoolsMsg = processPoolsMsg Or (strType = "SHW" And m_objShowPool(intRace).NumberOfRows <= 12)

        If processPoolsMsg Then
            drWPSPools.BeginEdit()
            drWPSPools.RACEKEY = intRace
            If (strType = "WIN") Then
                drWPSPools.POOLTYPEA = "WIN"
            ElseIf (strType = "PLC") Then
                drWPSPools.POOLTYPEA = "PLACE"
            ElseIf strType = "SHW" Then
                drWPSPools.POOLTYPEA = "SHOW"
            End If

            Try
                If ((intRace > 0) And (intRace <= p_intMaxNumbOfRaces)) Then
                    If (strType.Trim = "WIN") Then
                        If (p_blnWINPool) Then
                            objWPSPools = m_objWINPool(intRace)
                        End If
                    ElseIf (strType.Trim = "PLC") Then
                        If (p_blnPlacePool) Then
                            objWPSPools = m_objPlacePool(intRace)
                        End If
                    ElseIf (strType.Trim = "SHW") Then
                        If (p_blnShowPool) Then
                            objWPSPools = m_objShowPool(intRace)
                        End If
                    End If
                    If Not IsDBNull(objWPSPools.PoolCode) Then
                        blnWPSFound = True
                    End If
                End If

                'If objWPSPools Is Nothing Then Return

                If (blnWPSFound) Then
                    'grab the pool's total
                    s = Trim(objWPSPools.PoolTotal)
                    'fix amount
                    If s IsNot Nothing Then
                        s = Me.FixAmount(s, 6)
                        s = s.Replace(",", "")
                        IndexOfs = InStr(s, ".")
                        If (IndexOfs > 0) Then
                            s = s.Remove(IndexOfs - 1, Len(s) - IndexOfs + 1)
                        End If
                        s = s.ToString.PadLeft(6, " ")
                    Else
                        s = "      "
                    End If

                    'populate the pool total (6 spots (A-F))
                    drWPSPools.POOLTOTA = s.Substring(0, 1)
                    drWPSPools.POOLTOTB = s.Substring(1, 1)
                    drWPSPools.POOLTOTC = s.Substring(2, 1)
                    drWPSPools.POOLTOTD = s.Substring(3, 1)
                    drWPSPools.POOLTOTE = s.Substring(4, 1)
                    drWPSPools.POOLTOTF = s.Substring(5, 1)

                    'grab each individual horse's pool value
                    x = 8 'to set field index
                    For i As Integer = 1 To (objWPSPools.NumberOfRows)
                        If i > 12 Then Exit For 'in case more than 12 runners
                        '
                        s = Trim(objWPSPools.Amount(i).ToString())
                        '
                        If ((s IsNot Nothing) And (IsNumeric(s))) Then
                            s = Me.FixAmount(s, 5)
                            s = s.Replace(",", "")
                            IndexOfs = InStr(s, ".")
                            If (IndexOfs > 0) Then
                                s = s.Remove(IndexOfs - 1, Len(s) - IndexOfs + 1)
                            End If
                            s = s.ToString.PadLeft(5, " ")
                        Else
                            s = "     "
                        End If
                        '
                        drWPSPools(x) = s.Substring(0, 1)
                        drWPSPools(x + 1) = s.Substring(1, 1)
                        drWPSPools(x + 2) = s.Substring(2, 1)
                        drWPSPools(x + 3) = s.Substring(3, 1)
                        drWPSPools(x + 4) = s.Substring(4, 1)
                        x += 5
                        '
                    Next
                Else
                    'ClearWPS();
                End If

            Catch ex As Exception
                ClearPoolsObjectRow(drWPSPools)
            End Try

            drWPSPools.EndEdit()
            Me.myDataset.EXOTICS.AddEXOTICSRow(drWPSPools)
        Else
            ClearPoolsObjectRow(drWPSPools)
        End If

        objWPSPools = Nothing
    End Sub

    Private Sub ClearPoolsObjectRow(ByRef drWPSPools As RaceDisplayDataset.EXOTICSRow)
        Dim x As Integer
        Dim s As String

        drWPSPools.POOLTYPEA = ""
        'populate the pool total (6 spots (A-F))
        s = "      "
        drWPSPools.POOLTOTA = s.Substring(0, 1)
        drWPSPools.POOLTOTB = s.Substring(1, 1)
        drWPSPools.POOLTOTC = s.Substring(2, 1)
        drWPSPools.POOLTOTD = s.Substring(3, 1)
        drWPSPools.POOLTOTE = s.Substring(4, 1)
        drWPSPools.POOLTOTF = s.Substring(5, 1)

        x = 8  'to set field index
        s = "     "
        For i As Integer = 1 To 12 '12 runners

            drWPSPools(x) = s.Substring(0, 1)
            drWPSPools(x + 1) = s.Substring(1, 1)
            drWPSPools(x + 2) = s.Substring(2, 1)
            drWPSPools(x + 3) = s.Substring(3, 1)
            drWPSPools(x + 4) = s.Substring(4, 1)
            x += 5
        Next
    End Sub

    Private Sub UpdateOdds(ByVal ObjOdds As clsOdds, ByVal intRace As Integer)
        Dim s As String
        Dim x As Integer = 1 'to set field index
        '
        Me.myDataset.ODDS.Clear()
        Dim drOdds As RaceDisplayDataset.ODDSRow
        drOdds = Me.myDataset.ODDS.NewODDSRow
        '
        drOdds.BeginEdit()
        drOdds.RACEKEY = intRace
        'set all array elements to zero avery time we come here
        For i As Integer = 0 To Me.myArrOdds.Length - 1
            Me.myArrOdds(i) = 0
        Next
        Try
            '
            If ObjOdds Is Nothing Then Return

            'The board currently can process up to 12 horses
            'If more than 12, do not send odds
            If ObjOdds.NumberOfRows <= 12 Then
                For i As Integer = 1 To (ObjOdds.NumberOfRows)
                    'If i > 12 Then Exit For 'in case more than 16 odds ...
                    '
                    s = ObjOdds.OddsByRunner(i)
                    'fill up array to identify if odd has "/"
                    'Me.myArrOdds(i - 1) = If(s.Contains("/") OrElse s.Contains("-"), 1, 0)
                    Me.myArrOdds(i - 1) = If(s Like "#/#" OrElse s Like "#-#", 1, 0)
                    'Me.myArrOdds(i - 1) = IIf(s.Contains("/"), 1, 0)
                    '
                    If ((s IsNot Nothing) And (Val(s) > 0)) Then
                        s = Me.FixString(s)
                    Else
                        s = "  "
                    End If
                    '
                    drOdds(x) = s.Substring(0, 1)
                    drOdds(x + 1) = s.Substring(1, 1)
                    x += 2
                    '
                Next
            End If
            '
        Catch ex As Exception
            '
            x = 1
            For i As Integer = 1 To 12
                s = "  "
                drOdds(x) = s.Substring(0, 1)
                drOdds(x + 1) = s.Substring(1, 1)
                x += 2
            Next
            '
        End Try
        '
        drOdds.EndEdit()
        Me.myDataset.ODDS.AddODDSRow(drOdds)
        '
        RaiseEvent DisplayNewOdds()
        '
        ObjOdds = Nothing

    End Sub

    Private Sub UpdateRace(ByVal intRace As Integer)
        Me.myDataset.RACE.Clear()
        '
        Dim drRace As RaceDisplayDataset.RACERow
        '
        drRace = Me.myDataset.RACE.NewRACERow
        drRace.BeginEdit()
        drRace.RACEKEY = intRace
        drRace.RACENBRA = intRace.ToString.PadLeft(2, "").Substring(0, 1)
        drRace.RACENBRB = intRace.ToString.PadLeft(2, "").Substring(1, 1)
        drRace.EndEdit()
        '
        Me.myDataset.RACE.AddRACERow(drRace)
        '
        RaiseEvent DisplayNewRace()
        ''
    End Sub

    Public Sub UpdateRunningOrderAmtote(ByVal ObjRO As clsRunningOrder, ByVal intRace As Integer)
        If ObjRO Is Nothing Then Return
        Me.myDataset.RUNNINGORDER.Clear()
        '
        'get current status
        Me.myDataset.STATUS.Clear()
        Dim drStatus As RaceDisplayDataset.STATUSRow = Me.myDataset.STATUS.NewSTATUSRow
        'first things first get race status
        Dim status As String = ObjRO.FinishStatus()
        drStatus.BeginEdit()
        drStatus.RACEKEY = intRace 'row zeroByRef     
        For i As Integer = 1 To (drStatus.Table.Columns.Count - 1)
            drStatus(i) = False
        Next
        ' 'i' inquiry, 'j' objection, 'p' photo, 'h' dead heat, 'o' official, 'n' no race//'i' inquiry, 'j' objection, 'p' photo, 'h' dead heat, 'o' official, 'n' no race
        If status.Contains("i") Then drStatus.INQUIRY = True
        If status.Contains("j") Then drStatus.OBJECTION = True
        If status.Contains("h") Then drStatus.DEADHEAT = True
        If status.Contains("p") Then drStatus.PHOTO = True
        If status.Contains("o") Then drStatus.OFFICIAL = True
        If status.Contains("ij") Then
            drStatus.INQOBJ = True
            drStatus.INQUIRY = False
            drStatus.OBJECTION = False
        End If
        If status.Contains("ji") Then
            drStatus.INQOBJ = True
            drStatus.INQUIRY = False
            drStatus.OBJECTION = False
        End If
        '
        drStatus.EndEdit()
        '
        Dim s As String
        Dim c As String
        Dim z As Boolean
        Dim drRO As RaceDisplayDataset.RUNNINGORDERRow
        drRO = Me.myDataset.RUNNINGORDER.NewRUNNINGORDERRow
        '
        drRO.BeginEdit()
        drRO.RACEKEY = intRace
        '
        Dim x As Integer = 1
        Dim y As Integer = 9
        For i As Integer = 1 To (ObjRO.NumberOfRows)
            If i > 4 Then Exit For 'in case we have more than  4 places
            z = False
            c = " "
            c = ObjRO.Status(i)
            s = ObjRO.Program(i)
            'If s IsNot Nothing Then s = Me.FixString(s) Else s = "  "
            If s IsNot Nothing Then
                s = UCase(s)
                ''''
                If (InStr(c, "i") > 0) Then
                    z = True
                End If
                If (InStr(c, "j") > 0) Then
                    z = True
                End If
                If (InStr(c, "h") > 0) Then
                    z = True
                End If
                If (InStr(c, "p") > 0) Then
                    z = True
                End If
                If (InStr(c, "o") > 0) Then
                    z = True
                End If
                ''''
                If (InStr(s, "A") > 0) Then
                    s = s.Trim()
                    s = s.PadLeft(2, " ")
                ElseIf (InStr(s, "B") > 0) Then
                    s = s.Trim()
                    s = s.PadLeft(2, " ")
                ElseIf (InStr(s, "C") > 0) Then
                    s = s.Trim()
                    s = s.PadLeft(2, " ")
                Else
                    s = Me.FixString(s)
                End If

            Else
                s = "  "
            End If
            '
            drRO(x) = s.Substring(0, 1)
            drRO(x + 1) = s.Substring(1, 1)
            '
            drRO(y) = z
            '
            x += 2
            y += 1
            '
        Next
        '
        drRO.EndEdit()
        'Amtote
        Me.myDataset.STATUS.AddSTATUSRow(drStatus)
        Me.myDataset.RUNNINGORDER.AddRUNNINGORDERRow(drRO)
        '
        m_blnClearExotics = False
        If ObjRO.NumberOfRows = 0 Then
            m_blnClearExotics = True
        End If
        RaiseEvent DisplayRunningOrderAmtote()
        ''
    End Sub

    Public Sub UpdateRunningOrder(ByVal ObjRO As clsRunningOrder, ByVal intRace As Integer)
        If ObjRO Is Nothing Then Return
        Me.myDataset.RUNNINGORDER.Clear()
        '
        Dim s As String
        Dim z As Boolean
        Dim drRO As RaceDisplayDataset.RUNNINGORDERRow
        drRO = Me.myDataset.RUNNINGORDER.NewRUNNINGORDERRow
        '
        drRO.BeginEdit()
        drRO.RACEKEY = intRace
        '
        Dim x As Integer = 1
        Dim y As Integer = 9
        For i As Integer = 1 To (ObjRO.NumberOfRows)
            If i > 4 Then Exit For 'in case we have more than  4 places
            z = False
            s = ObjRO.Program(i)
            If s IsNot Nothing Then
                s = UCase(s)
                ''''
                If (InStr(s, "A") > 0) Then
                    s = s.Trim()
                    s = s.PadLeft(2, " ")
                ElseIf (InStr(s, "B") > 0) Then
                    s = s.Trim()
                    s = s.PadLeft(2, " ")
                ElseIf (InStr(s, "C") > 0) Then
                    s = s.Trim()
                    s = s.PadLeft(2, " ")
                Else
                    s = Me.FixString(s)
                End If
            Else
                s = "  "
            End If
            '
            drRO(x) = s.Substring(0, 1)
            drRO(x + 1) = s.Substring(1, 1)
            '
            drRO(y) = z
            '
            x += 2
            y += 1
            '
        Next
        '
        drRO.EndEdit()
        Me.myDataset.RUNNINGORDER.AddRUNNINGORDERRow(drRO)
        '
        m_blnClearExotics = False
        If ObjRO.NumberOfRows = 0 Then
            m_blnClearExotics = True
        End If
        RaiseEvent DisplayRunningOrder()
        ''
    End Sub

    Public Sub UpdateTeletimer(ByVal ObjTT As RSIData.clsTeleTimer, ByVal intRace As Integer)

        If ObjTT Is Nothing Then Return

        Me.myDataset.TIMINGFINISH.Clear()
        Me.myDataset.TIMINGMILE.Clear()
        Me.myDataset.TIMING34.Clear()
        Me.myDataset.TIMING12.Clear()
        Me.myDataset.TIMING14.Clear()

        Dim sTime As String
        Dim blnFinshTaken As Boolean
        blnFinshTaken = False
        For n As Integer = 1 To ObjTT.NumberOfRows
            sTime = ""
            If ObjTT.Description(n).ToString.ToLower.Contains("fin") Then
                sTime = ObjTT.Time(n).ToString
                sTime = sTime.Replace(".", ":")
                If (InStr(sTime, ":", CompareMethod.Text) = 0) Then
                    sTime = sTime & ":00"
                End If
                sTime = sTime.PadLeft(7, " ")
                If (sTime.Substring(5, 1) = ":") Then
                    sTime = sTime & "0"
                    sTime = Trim(sTime)
                End If
                sTime = sTime.PadLeft(7, " ")
                Dim drfinish As RaceDisplayDataset.TIMINGFINISHRow = Me.myDataset.TIMINGFINISH.NewTIMINGFINISHRow
                drfinish.BeginEdit()
                drfinish(1) = sTime.Substring(0, 1)
                drfinish(2) = sTime.Substring(2, 1)
                drfinish(3) = sTime.Substring(3, 1)
                drfinish(4) = sTime.Substring(5, 1)
                drfinish(5) = sTime.Substring(6, 1)
                drfinish.EndEdit()
                Me.myDataset.TIMINGFINISH.AddTIMINGFINISHRow(drfinish)
                blnFinshTaken = True
            ElseIf ((ObjTT.Description(n).ToString.ToLower.Trim() = "1 mile") Or (ObjTT.Description(n).ToString.ToLower.Trim() = "7/8") Or (ObjTT.Description(n).ToString.ToLower.Trim() = "mile")) Then
                sTime = ObjTT.Time(n).ToString
                sTime = sTime.Replace(".", ":")
                If (InStr(sTime, ":", CompareMethod.Text) = 0) Then
                    sTime = sTime & ":00"
                End If
                sTime = sTime.PadLeft(7, " ")
                If (sTime.Substring(5, 1) = ":") Then
                    sTime = sTime & "0"
                    sTime = Trim(sTime)
                End If
                sTime = sTime.PadLeft(7, " ")
                Dim drmile As RaceDisplayDataset.TIMINGMILERow = Me.myDataset.TIMINGMILE.NewTIMINGMILERow
                drmile.BeginEdit()
                drmile(1) = sTime.Substring(0, 1)
                drmile(2) = sTime.Substring(2, 1)
                drmile(3) = sTime.Substring(3, 1)
                drmile(4) = sTime.Substring(5, 1)
                drmile(5) = sTime.Substring(6, 1)
                drmile.EndEdit()
                Me.myDataset.TIMINGMILE.AddTIMINGMILERow(drmile)
            ElseIf ((ObjTT.Description(n).ToString.ToLower.Trim() = "3/4") Or (ObjTT.Description(n).ToString.ToLower.Trim() = "5/8")) Then
                sTime = ObjTT.Time(n).ToString
                sTime = sTime.Replace(".", ":")
                If (InStr(sTime, ":", CompareMethod.Text) = 0) Then
                    sTime = sTime & ":00"
                End If
                sTime = sTime.PadLeft(7, " ")
                If (sTime.Substring(5, 1) = ":") Then
                    sTime = sTime & "0"
                    sTime = Trim(sTime)
                End If
                sTime = sTime.PadLeft(7, " ")
                Dim dr34 As RaceDisplayDataset.TIMING34Row = Me.myDataset.TIMING34.NewTIMING34Row
                dr34.BeginEdit()
                dr34(1) = sTime.Substring(0, 1)
                dr34(2) = sTime.Substring(2, 1)
                dr34(3) = sTime.Substring(3, 1)
                dr34(4) = sTime.Substring(5, 1)
                dr34(5) = sTime.Substring(6, 1)
                dr34.EndEdit()
                Me.myDataset.TIMING34.AddTIMING34Row(dr34)
            ElseIf (ObjTT.Description(n).ToString.ToLower.Trim() = "1/2") Then
                sTime = ObjTT.Time(n).ToString
                sTime = sTime.Replace(".", ":")
                If (InStr(sTime, ":", CompareMethod.Text) = 0) Then
                    sTime = sTime & ":00"
                End If
                sTime = sTime.PadLeft(7, " ")
                If (sTime.Substring(5, 1) = ":") Then
                    sTime = sTime & "0"
                    sTime = Trim(sTime)
                End If
                sTime = sTime.PadLeft(7, " ")
                Dim dr12 As RaceDisplayDataset.TIMING12Row = Me.myDataset.TIMING12.NewTIMING12Row
                dr12.BeginEdit()
                dr12(1) = sTime.Substring(0, 1)
                dr12(2) = sTime.Substring(2, 1)
                dr12(3) = sTime.Substring(3, 1)
                dr12(4) = sTime.Substring(5, 1)
                dr12(5) = sTime.Substring(6, 1)
                dr12.EndEdit()
                Me.myDataset.TIMING12.AddTIMING12Row(dr12)
            ElseIf (ObjTT.Description(n).ToString.ToLower.Trim() = "1/4") Then
                sTime = ObjTT.Time(n).ToString
                sTime = sTime.Replace(".", ":")
                If (InStr(sTime, ":", CompareMethod.Text) = 0) Then
                    sTime = sTime & ":00"
                End If
                sTime = sTime.PadLeft(7, " ")
                If (sTime.Substring(5, 1) = ":") Then
                    sTime = sTime & "0"
                    sTime = Trim(sTime)
                End If
                sTime = sTime.PadLeft(7, " ")
                Dim dr14 As RaceDisplayDataset.TIMING14Row = Me.myDataset.TIMING14.NewTIMING14Row
                dr14.BeginEdit()
                dr14(1) = sTime.Substring(0, 1)
                dr14(2) = sTime.Substring(2, 1)
                dr14(3) = sTime.Substring(3, 1)
                dr14(4) = sTime.Substring(5, 1)
                dr14(5) = sTime.Substring(6, 1)
                dr14.EndEdit()
                Me.myDataset.TIMING14.AddTIMING14Row(dr14)
            Else
                If (blnFinshTaken = False) Then
                    sTime = ObjTT.Time(n).ToString
                    sTime = sTime.Replace(".", ":")
                    If (InStr(sTime, ":", CompareMethod.Text) = 0) Then
                        sTime = sTime & ":00"
                    End If
                    sTime = sTime.PadLeft(7, " ")
                    If (sTime.Substring(5, 1) = ":") Then
                        sTime = sTime & "0"
                        sTime = Trim(sTime)
                    End If
                    sTime = sTime.PadLeft(7, " ")
                    Dim drfinish As RaceDisplayDataset.TIMINGFINISHRow = Me.myDataset.TIMINGFINISH.NewTIMINGFINISHRow
                    drfinish.BeginEdit()
                    drfinish(1) = sTime.Substring(0, 1)
                    drfinish(2) = sTime.Substring(2, 1)
                    drfinish(3) = sTime.Substring(3, 1)
                    drfinish(4) = sTime.Substring(5, 1)
                    drfinish(5) = sTime.Substring(6, 1)
                    drfinish.EndEdit()
                    Me.myDataset.TIMINGFINISH.AddTIMINGFINISHRow(drfinish)
                    blnFinshTaken = True
                End If
            End If
        Next

        RaiseEvent DisplayTeletimer()

    End Sub

    Public Sub UpdateOfficialResults(ByVal objFinisher As RSIData.clsFinisherData, ByVal intRace As Integer)
        p_intRaceResultsOut = intRace
        p_blnResultsOut = True

        Try
            If objFinisher Is Nothing Then Return
            Me.myDataset.RUNNINGORDER.Clear()
            '
            Dim drRO As RaceDisplayDataset.RUNNINGORDERRow
            drRO = Me.myDataset.RUNNINGORDER.NewRUNNINGORDERRow
            '
            drRO.BeginEdit()
            drRO.RACEKEY = intRace
            '
            Dim x As Integer = 1
            Dim y As Integer = 9
            Dim intCtr As Integer = 1

            For intCtrFor As Integer = 1 To p_intMaxFinisherPositions
                p_udtWinners(intRace, intCtrFor).strRunner = ""
                p_udtWinners(intRace, intCtrFor).intRunner = 0
                p_udtWinners(intRace, intCtrFor).blnDH = False
                p_udtWinners(intRace, intCtrFor).intFinishPosition = 0
            Next

            Dim strWinner As String
            Dim intWinner As Integer
            Dim intNumOfRows As Integer
            Dim intNumOfCols As Integer
            Dim intFinishPos As Integer
            Dim blnDH As Boolean = False
            Dim shrCtrFor As Short
            Dim shrCtrFor2 As Short
            Dim intPost As Integer = 0

            intNumOfRows = objFinisher.NumberOfRows
            For i As Integer = 1 To intNumOfRows
                shrCtrFor = CShort(i)
                intFinishPos = objFinisher.FinishPosition(shrCtrFor)
                intNumOfCols = objFinisher.NumberOfColumns(shrCtrFor)
                blnDH = objFinisher.DH(shrCtrFor)

                If intCtr > 4 Then Exit For 'in case we have more than  4 places

                'if the object objFinisher doesn't have data for that row (intCtrFor), then
                'the variable 'intFinishPos' will be '0'.
                If ((intFinishPos >= 1) And (intFinishPos <= p_intMaxFinisherPositions)) Then
                    For j As Integer = 1 To (intNumOfCols)

                        If intCtr > 4 Then Exit For 'in case we have more than  4 places

                        shrCtrFor2 = CShort(j)
                        strWinner = objFinisher.Program(shrCtrFor, shrCtrFor2).Trim()
                        If (Mid(strWinner, 1, 1) = "0") Then
                            strWinner = Mid(strWinner, 2)
                        End If
                        intWinner = objFinisher.Runner(shrCtrFor, shrCtrFor2)
                        intPost = intPost + 1
                        p_udtWinners(intRace, intPost).strRunner = strWinner
                        p_udtWinners(intRace, intPost).intRunner = Convert.ToString(intWinner)
                        p_udtWinners(intRace, intPost).intFinishPosition = intFinishPos
                        p_udtWinners(intRace, intPost).blnDH = blnDH

                        If strWinner IsNot Nothing Then
                            strWinner = UCase(strWinner)
                            ''''
                            If (InStr(strWinner, "A") > 0) Then
                                strWinner = strWinner.Trim()
                                strWinner = strWinner.PadLeft(2, " ")
                            ElseIf (InStr(strWinner, "B") > 0) Then
                                strWinner = strWinner.Trim()
                                strWinner = strWinner.PadLeft(2, " ")
                            ElseIf (InStr(strWinner, "C") > 0) Then
                                strWinner = strWinner.Trim()
                                strWinner = strWinner.PadLeft(2, " ")
                            Else
                                strWinner = Me.FixString(strWinner)
                            End If
                        Else
                            strWinner = "  "
                        End If
                        '
                        drRO(x) = strWinner.Substring(0, 1)
                        drRO(x + 1) = strWinner.Substring(1, 1)
                        '
                        drRO(y) = blnDH
                        '
                        x += 2
                        y += 1
                        '            
                        intCtr += 1
                    Next
                End If
            Next
            '
            drRO.EndEdit()
            Me.myDataset.RUNNINGORDER.AddRUNNINGORDERRow(drRO)
            '
            m_blnClearExotics = False
            If objFinisher.NumberOfRows = 0 Then
                m_blnClearExotics = True
            End If
            ''
        Catch ex As Exception
            'nothing
        Finally
            RaiseEvent DisplayNewResults()
        End Try
    End Sub

    Public Sub UpdateResultsWPS(intRace As Integer)
        'RS5_WIN =^CAMA | RS | 05 | 07 / 06 / 2018 | WIN | 01 | N | 10 | 000000310^
        'RS5_PLC =^CAMA | RS | 05 | 07 / 06 / 2018 | PLC | 02 | N | 10 | 000000260 | 03 | 000000700^
        Dim intProgram As Integer
        Dim strTemp As String = ""
        Dim intTagPost As Integer

        Try

            Me.myDataset.WIN.Clear()
            Me.myDataset.PLACE.Clear()
            Me.myDataset.SHOW.Clear()

            Dim shrRace As Short = CShort(intRace)

            If (p_blnResultsOut) Then
                'If (p_intRaceResultsOut > 0) And (p_intRaceResultsOut = intRace) Then
                If (p_intRaceResultsOut > 0) Then
                    For intTagPost = 1 To 4
                        intProgram = p_udtWinners(intRace, intTagPost).intRunner

                        If intTagPost <= 2 Then
                            Me.UpdateWPS(intRace, intProgram, WPSType.WIN, intTagPost)
                        End If
                        If intTagPost <= 3 Then
                            Me.UpdateWPS(intRace, intProgram, WPSType.PLC, intTagPost)
                        End If
                        If intTagPost <= 4 Then
                            Me.UpdateWPS(intRace, intProgram, WPSType.SHW, intTagPost)
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            ex.ToString()
        End Try
    End Sub

    Private Sub UpdateExotics(ByVal intRace As Integer, ByVal myExoticType As ExoticType)
        Try 'new code all changes in this object took place in this method
            Dim ObjExotics As New RSIData.clsResultExotic
            Dim drExotic As DataRow
            Dim newRow As Boolean
            Dim intAmountDigits As Integer = 7
            '

            Dim shtTmp As Short = ObjExotics.NumberOfPrices
            Select Case myExoticType
                Case ExoticType.DD
                    'ObjExotics = Me.oCommServer.ObjectRsExo(intRace, intRace & "_" & "2DD ").Item(intRace & "_" & "2DD ")
                    ObjExotics = Me.oCommServerNet.ObjectRsExo(intRace, intRace & "_" & "2DD ").Item(intRace & "_" & "2DD ")
                    If ObjExotics Is Nothing Then Return
                    '
                    Dim dr As RaceDisplayDataset.DAILYDOUBLERow
                    If Me.myDataset.DAILYDOUBLE.Rows.Count = 0 Then
                        dr = Me.myDataset.DAILYDOUBLE.NewDAILYDOUBLERow
                        newRow = True
                    Else
                        dr = Me.myDataset.DAILYDOUBLE(0)
                    End If
                    drExotic = dr
                    '
                Case ExoticType.Perfecta
                    'ObjExotics = Me.oCommServer.ObjectRsExo(intRace, intRace & "_" & "2EX ").Item(intRace & "_" & "2EX ")
                    ObjExotics = Me.oCommServerNet.ObjectRsExo(intRace, intRace & "_" & "2EX ").Item(intRace & "_" & "2EX ")
                    If ObjExotics Is Nothing Then Return
                    '
                    myNumberExacta = ObjExotics.NumberOfPrices
                    Dim dr As RaceDisplayDataset.PERFECTARow
                    If Me.myDataset.PERFECTA.Rows.Count = 0 Then
                        dr = Me.myDataset.PERFECTA.NewPERFECTARow
                        newRow = True
                    Else
                        dr = Me.myDataset.PERFECTA(0)
                    End If
                    drExotic = dr
                    intAmountDigits = 8
                    '
                Case ExoticType.BET3
                    'ObjExotics = Me.oCommServer.ObjectRsExo(intRace, intRace & "_" & "2P03").Item(intRace & "_" & "2P03")
                    ObjExotics = Me.oCommServerNet.ObjectRsExo(intRace, intRace & "_" & "2P03").Item(intRace & "_" & "2P03")
                    If ObjExotics Is Nothing Then Return
                    '
                    Dim dr As RaceDisplayDataset.BET3Row
                    If Me.myDataset.BET3.Rows.Count = 0 Then
                        dr = Me.myDataset.BET3.NewBET3Row
                        newRow = True
                    Else
                        dr = Me.myDataset.BET3(0)
                    End If
                    drExotic = dr
                    intAmountDigits = 8 'new code, pick 3 and trifecta now have 8 places
                    '
                Case ExoticType.Trifecta
                    'ObjExotics = Me.oCommServer.ObjectRsExo(intRace, intRace & "_" & "2TRI").Item(intRace & "_" & "2TRI")
                    ObjExotics = Me.oCommServerNet.ObjectRsExo(intRace, intRace & "_" & "2TRI").Item(intRace & "_" & "2TRI")
                    If ObjExotics Is Nothing Then Return
                    '
                    myNumberTrifecta = ObjExotics.NumberOfPrices

                    Dim dr As RaceDisplayDataset.TRIFECTARow
                    If Me.myDataset.TRIFECTA.Rows.Count = 0 Then
                        dr = Me.myDataset.TRIFECTA.NewTRIFECTARow
                        newRow = True
                    Else
                        dr = Me.myDataset.TRIFECTA(0)
                    End If
                    drExotic = dr
                    intAmountDigits = 8
                    '
                Case Else
                    Return
            End Select
            '
            Dim strWinnerList As String
            Dim strAmount As String
            Dim intPos As Integer = 1
            Dim intNumberOfPrices As Integer
            intNumberOfPrices = ObjExotics.NumberOfPrices
            '
            drExotic.BeginEdit()
            drExotic("RACEKEY") = intRace
            '
            Dim i As Integer
            intCurrPage(myExoticType) += 1
            If intCurrPage(myExoticType) > intNumberOfPrices Then
                intCurrPage(myExoticType) = 1
            End If
            i = intCurrPage(myExoticType)
            'For i = i To intNumberOfPrices
            strWinnerList = ObjExotics.WinnerList(i)
            If strWinnerList IsNot Nothing Then
                Dim blnNeedToGetWinner As Boolean
                Dim blnFoundAll As Boolean
                Dim a() As String = FixArray(strWinnerList)
                For x As Integer = 0 To a.Length - 1
                    blnNeedToGetWinner = (myExoticType = ExoticType.BET3) AndAlso (a(x).Contains(","))
                    blnFoundAll = Not a(x).Trim = "" AndAlso a(x).ToUpper.Contains("ALL")
                    '
                    If blnNeedToGetWinner Then
                        Dim intRaceNbr As Integer
                        Select Case x
                            Case 0
                                intRaceNbr = intRace - 2
                            Case 1
                                intRaceNbr = intRace - 1
                            Case 2
                                intRaceNbr = intRace
                        End Select
                        a(x) = Me.GetSpecificRaceWinner(intRaceNbr)
                    ElseIf blnFoundAll Then
                        a(x) = "[]"
                    End If
                    '
                    drExotic(intPos) = a(x).PadLeft(2, " ").Substring(0, 1).Replace("0", "")
                    drExotic(intPos + 1) = a(x).PadLeft(2, " ").Substring(1, 1)
                    '
                    intPos += 2
                Next
            End If
            'done with winners, do amount

            strAmount = ObjExotics.Amount(i)

            'Try
            '    If (IsNumeric(strAmount)) Then
            '        strAmount = (Val(strAmount)) * 2
            '    End If
            'Catch ex As Exception
            '    '
            'End Try

            If strAmount IsNot Nothing Then strAmount = Me.FixAmount(strAmount, intAmountDigits)
            '
            drExotic(intPos) = strAmount.Substring(0, 1)
            drExotic(intPos + 1) = strAmount.Substring(1, 1)
            drExotic(intPos + 2) = strAmount.Substring(2, 1)
            drExotic(intPos + 3) = strAmount.Substring(3, 1)
            Select Case intAmountDigits
                Case 7
                    '4th place is the ".", jump it
                    drExotic(intPos + 4) = strAmount.Substring(5, 1)
                    drExotic(intPos + 5) = strAmount.Substring(6, 1)
                Case 8
                    '5th place is the ".", jump it
                    drExotic(intPos + 4) = strAmount.Substring(4, 1)
                    drExotic(intPos + 5) = strAmount.Substring(6, 1)
                    drExotic(intPos + 6) = strAmount.Substring(7, 1)
            End Select
            '
            drExotic.EndEdit()
            '
            If Not newRow Then Return
            'identify exotic type, update or insert row
            Select Case myExoticType
                Case ExoticType.DD
                    Me.myDataset.DAILYDOUBLE.AddDAILYDOUBLERow(drExotic)
                Case ExoticType.Perfecta
                    Me.myDataset.PERFECTA.AddPERFECTARow(drExotic)
                Case ExoticType.BET3
                    Me.myDataset.BET3.AddBET3Row(drExotic)
                Case ExoticType.Trifecta
                    Me.myDataset.TRIFECTA.AddTRIFECTARow(drExotic)
            End Select
            '
        Catch ex As Exception
            'do nothing
        End Try
        ''
    End Sub

    Private Sub UpdateWPS(ByVal intRace As Integer, ByVal HorseNbr As String, ByVal myWpsType As WPSType, ByVal Post As Integer)
        Try
            Dim ObjWPS As New RSIData.clsResultWPS
            Dim strType As String = myWpsType.ToString
            '
            If (strType.Trim.ToUpper() = "WIN") Then
                ObjWPS = objResultWin(intRace)
            ElseIf (strType.Trim.ToUpper() = "PLC") Then
                ObjWPS = objResultPlc(intRace)
            ElseIf (strType.Trim.ToUpper() = "SHW") Then
                ObjWPS = objResultShow(intRace)
            End If

            If ObjWPS Is Nothing Then Return
            '
            Dim s As String
            Dim newRow As Boolean
            Dim drWPS As DataRow
            Dim intColumns As Integer = 6
            '
            Select Case myWpsType
                Case WPSType.WIN
                    Dim dr As RaceDisplayDataset.WINRow
                    If Me.myDataset.WIN.Rows.Count = 0 Then
                        dr = Me.myDataset.WIN.NewWINRow
                        newRow = True
                    Else
                        dr = Me.myDataset.WIN(0)
                    End If
                    drWPS = dr
                    intColumns = 7 'WIN table has two more columns than PLACE and SHOW
                Case WPSType.PLC
                    Dim dr As RaceDisplayDataset.PLACERow
                    If Me.myDataset.PLACE.Rows.Count = 0 Then
                        dr = Me.myDataset.PLACE.NewPLACERow
                        newRow = True
                    Else
                        dr = Me.myDataset.PLACE(0)
                    End If
                    drWPS = dr
                Case WPSType.SHW
                    Dim dr As RaceDisplayDataset.SHOWRow
                    If Me.myDataset.SHOW.Rows.Count = 0 Then
                        dr = Me.myDataset.SHOW.NewSHOWRow
                        newRow = True
                    Else
                        dr = Me.myDataset.SHOW(0)
                    End If
                    drWPS = dr
                Case Else
                    Return
            End Select
            '
            Dim i As Integer
            'field index position to start update
            'If newRow Then
            '    i = 1
            'ElseIf drWPS(intColumns) Is System.DBNull.Value Then
            '    i = intColumns
            'ElseIf drWPS.Table.Columns.Count >= 14 AndAlso drWPS(11) Is System.DBNull.Value Then
            '    i = 11
            'ElseIf drWPS.Table.Columns.Count >= 17 AndAlso drWPS(16) Is System.DBNull.Value Then
            '    i = 16
            'Else 'cover your back
            '    i = 0 'Before it was i = 1 (Tania)
            'End If

            If Post = 1 Then
                i = 1
            ElseIf Post = 2 Then
                i = intColumns
            ElseIf drWPS.Table.Columns.Count >= 14 AndAlso Post = 3 Then
                i = 11
            ElseIf drWPS.Table.Columns.Count >= 17 AndAlso Post = 4 Then
                i = 16
            Else 'cover your back
                i = 0 'Before it was i = 1 (Tania)
            End If

            drWPS("RACEKEY") = intRace
            s = ObjWPS.AmountByRunner(HorseNbr)

            'This part should be comment (Tania)
            'get out if "" is returned, it means horse does not have this place
            'If s.Trim = "" OrElse (s Is Nothing) Then
            '    drWPS = Nothing
            '    Return
            'End If

            'This part is new (Tania)
            If (s Is Nothing) Then
                drWPS = Nothing
                Return
                's = ""
            End If

            'fix amount
            If (s.Trim() = "") Then
                s = s.ToString.PadLeft(intColumns, " ")
            Else
                s = Me.FixAmount(s, intColumns)
            End If

            '
            drWPS(i) = s.Substring(0, 1)
            drWPS(i + 1) = s.Substring(1, 1)
            drWPS(i + 2) = s.Substring(2, 1)
            Select Case intColumns
                Case 6
                    drWPS(i + 3) = s.Substring(4, 1) '3 has the . value
                    drWPS(i + 4) = s.Substring(5, 1)
                Case 7
                    drWPS(i + 3) = s.Substring(3, 1)
                    drWPS(i + 4) = s.Substring(5, 1) '4 has the . value
                    drWPS(i + 5) = s.Substring(6, 1)
            End Select
            '
            drWPS.EndEdit()
            'get out if not new row
            If Not newRow Then Return
            '
            Select Case myWpsType
                Case WPSType.WIN
                    Me.myDataset.WIN.AddWINRow(drWPS)
                Case WPSType.PLC
                    Me.myDataset.PLACE.AddPLACERow(drWPS)
                Case WPSType.SHW
                    Me.myDataset.SHOW.AddSHOWRow(drWPS)
            End Select
        Catch ex As Exception
            'nothing
        End Try
        ''
    End Sub

    'Private Sub UpdateTiming()
    '    '
    '    Me.myDataset.TIMINGFINISH.Clear()
    '    Me.myDataset.TIMINGMILE.Clear()
    '    Me.myDataset.TIMING34.Clear()
    '    Me.myDataset.TIMING12.Clear()
    '    Me.myDataset.TIMING14.Clear()
    '    '
    '    Dim sTime As String
    '    For n As Integer = 1 To oCommServer.p_objTimerData.NumDistances
    '        sTime = ""
    '        If oCommServer.p_objTimerData.DistanceData(n).Distance.ToString.ToLower.Contains("fin") Then
    '            sTime = oCommServer.p_objTimerData.DistanceData(n).Time.ToString.PadLeft(7, " ")
    '            '
    '            Dim drfinish As RaceDisplayDataset.TIMINGFINISHRow = Me.myDataset.TIMINGFINISH.NewTIMINGFINISHRow
    '            drfinish.BeginEdit()
    '            'drfinish.TIMINGFINISHA = sTime.Substring(0, 1)
    '            'drfinish.TIMINGFINISHB = sTime.Substring(2, 1)
    '            'drfinish.TIMINGFINISHC = sTime.Substring(3, 1)
    '            'drfinish.TIMINGFINISHD = sTime.Substring(5, 1)
    '            'drfinish.TIMINGFINISHE = sTime.Substring(6, 1)
    '            drfinish(1) = sTime.Substring(0, 1)
    '            drfinish(2) = sTime.Substring(2, 1)
    '            drfinish(3) = sTime.Substring(3, 1)
    '            drfinish(4) = sTime.Substring(5, 1)
    '            drfinish(5) = sTime.Substring(6, 1)
    '            drfinish.EndEdit()
    '            '
    '            Me.myDataset.TIMINGFINISH.AddTIMINGFINISHRow(drfinish)
    '            '
    '        ElseIf oCommServer.p_objTimerData.DistanceData(n).Distance.ToString.ToLower.Contains("mile") Then
    '            sTime = oCommServer.p_objTimerData.DistanceData(n).Time.ToString.PadLeft(7, " ")
    '            '
    '            Dim drmile As RaceDisplayDataset.TIMINGMILERow = Me.myDataset.TIMINGMILE.NewTIMINGMILERow
    '            drmile.BeginEdit()
    '            'drmile.TIMINGMILEA = sTime.Substring(0, 1)
    '            'drmile.TIMINGMILEB = sTime.Substring(2, 1)
    '            'drmile.TIMINGMILEC = sTime.Substring(3, 1)
    '            'drmile.TIMINGMILED = sTime.Substring(5, 1)
    '            'drmile.TIMINGMILEE = sTime.Substring(6, 1)
    '            drmile(1) = sTime.Substring(0, 1)
    '            drmile(2) = sTime.Substring(2, 1)
    '            drmile(3) = sTime.Substring(3, 1)
    '            drmile(4) = sTime.Substring(5, 1)
    '            drmile(5) = sTime.Substring(6, 1)
    '            drmile.EndEdit()
    '            '
    '            Me.myDataset.TIMINGMILE.AddTIMINGMILERow(drmile)
    '            '
    '        ElseIf oCommServer.p_objTimerData.DistanceData(n).Distance.ToString.ToLower.Contains("3/4") Then
    '            sTime = oCommServer.p_objTimerData.DistanceData(n).Time.ToString.PadLeft(7, " ")
    '            '
    '            Dim dr34 As RaceDisplayDataset.TIMING34Row = Me.myDataset.TIMING34.NewTIMING34Row
    '            dr34.BeginEdit()
    '            'dr34.TIMING34A = sTime.Substring(0, 1)
    '            'dr34.TIMING34B = sTime.Substring(2, 1)
    '            'dr34.TIMING34C = sTime.Substring(3, 1)
    '            'dr34.TIMING34D = sTime.Substring(5, 1)
    '            'dr34.TIMING34E = sTime.Substring(6, 1)
    '            dr34(1) = sTime.Substring(0, 1)
    '            dr34(2) = sTime.Substring(2, 1)
    '            dr34(3) = sTime.Substring(3, 1)
    '            dr34(4) = sTime.Substring(5, 1)
    '            dr34(5) = sTime.Substring(6, 1)
    '            dr34.EndEdit()
    '            '
    '            Me.myDataset.TIMING34.AddTIMING34Row(dr34)
    '            '
    '        ElseIf oCommServer.p_objTimerData.DistanceData(n).Distance.ToString.ToLower.Contains("1/2") Then
    '            sTime = oCommServer.p_objTimerData.DistanceData(n).Time.ToString.PadLeft(7, " ")
    '            '
    '            Dim dr12 As RaceDisplayDataset.TIMING12Row = Me.myDataset.TIMING12.NewTIMING12Row
    '            dr12.BeginEdit()
    '            'dr12.TIMING12A = sTime.Substring(0, 1)
    '            'dr12.TIMING12B = sTime.Substring(2, 1)
    '            'dr12.TIMING12C = sTime.Substring(3, 1)
    '            'dr12.TIMING12D = sTime.Substring(5, 1)
    '            'dr12.TIMING12E = sTime.Substring(6, 1)
    '            dr12(1) = sTime.Substring(0, 1)
    '            dr12(2) = sTime.Substring(2, 1)
    '            dr12(3) = sTime.Substring(3, 1)
    '            dr12(4) = sTime.Substring(5, 1)
    '            dr12(5) = sTime.Substring(6, 1)
    '            dr12.EndEdit()
    '            '
    '            Me.myDataset.TIMING12.AddTIMING12Row(dr12)
    '            '
    '        ElseIf oCommServer.p_objTimerData.DistanceData(n).Distance.ToString.ToLower.Contains("1/4") Then
    '            sTime = oCommServer.p_objTimerData.DistanceData(n).Time.ToString.PadLeft(7, " ")
    '            '
    '            Dim dr14 As RaceDisplayDataset.TIMING14Row = Me.myDataset.TIMING14.NewTIMING14Row
    '            dr14.BeginEdit()
    '            'dr14.TIMING14A = sTime.Substring(0, 1)
    '            'dr14.TIMING14B = sTime.Substring(2, 1)
    '            'dr14.TIMING14C = sTime.Substring(3, 1)
    '            'dr14.TIMING14D = sTime.Substring(5, 1)
    '            'dr14.TIMING14E = sTime.Substring(6, 1)
    '            dr14(1) = sTime.Substring(0, 1)
    '            dr14(2) = sTime.Substring(2, 1)
    '            dr14(3) = sTime.Substring(3, 1)
    '            dr14(4) = sTime.Substring(5, 1)
    '            dr14(5) = sTime.Substring(6, 1)
    '            dr14.EndEdit()
    '            '
    '            Me.myDataset.TIMING14.AddTIMING14Row(dr14)
    '        End If
    '    Next
    '    RaiseEvent DisplayTiming()
    'End Sub

    Private Function GetSpecificRaceWinner(ByVal intRace As Integer) As String
        If Not intRace > 0 Then Return ""
        '
        Try
            Dim ObjWinner As New RSIData.clsFinisherData
            'ObjWinner = Me.oCommServer.ObjectFinisher(intRace).Item("F" & intRace)
            ObjWinner = Me.oCommServerNet.ObjectFinisher(intRace).Item("F" & intRace)
            If ObjWinner Is Nothing Then Return ""
            'we have a winner for the race
            Return Me.FixString(ObjWinner.Runner(1, 1).ToString)
        Catch ex As Exception
            Return ""
        End Try
    End Function

#End Region

#Region " Get Methods "
    '
#End Region

#Region "Helper Methods "

    Private Function FixArray(ByVal str As String) As System.Array
        Dim a() As String = str.Split(New Char() {"/", "-", "*"})
        '
        Return a
    End Function

    Private Function FixAmount(ByVal str As String, ByVal intNumberOfPlaces As Integer) As String
        If str.Trim = "" Then Return str
        '
        Dim c As Char() = str
        For i As Integer = 0 To (c.Length - 1)
            If (Not Asc(c(i)) = 46) And (Not Asc(c(i)) >= 48) OrElse (Not Asc(c(i)) <= 57) Then
                str = Replace(str, c(i), "")
            End If
        Next
        'convert string to dec, pad it later
        Dim decAmount As Decimal = FormatNumber(str, 2)
        Dim strAmount As String = (decAmount / 100).ToString
        'return it as string, remember the dot
        Return strAmount.ToString.PadLeft(intNumberOfPlaces, " ")
        ''
    End Function

    Private Function FixString(ByVal str As String) As String
        If str.Trim = "" Then Return str
        '
        Dim c As Char() = str
        For i As Integer = 0 To (c.Length - 1)
            If (Not Asc(c(i)) >= 48) OrElse (Not Asc(c(i)) <= 57) Then
                str = Replace(str, c(i), "")
            End If
        Next
        '
        Return str.PadLeft(2, " ")
        ''
    End Function

    Private Function FixString(ByVal str As String, ByVal PadDigit As Integer) As String
        If str.Trim = "" Then Return str
        '
        Dim c As Char() = str
        For i As Integer = 0 To (c.Length - 1)
            If (Not Asc(c(i)) >= 48) OrElse (Not Asc(c(i)) <= 57) Then
                str = Replace(str, c(i), "")
            End If
        Next
        '
        Return str.PadLeft(PadDigit, " ")
        ''
    End Function


#End Region

End Class