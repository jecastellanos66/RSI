Option Strict Off
Option Explicit On

Imports System.Configuration
Imports VB = Microsoft.VisualBasic

Public Class clsCommSvr
    Implements RSIService.IRaceFxNotify

    Private myError As String

    Public Structure typTrackingOdds 'Define user-defined type.
        Dim strOdds As String
    End Structure

    Public Structure typOddsStatistics 'Define user-defined type.
        Dim strHi As String
        Dim strLo As String
        <VBFixedArray(5)> Dim udtOddsTracking() As typTrackingOdds 'Declare a static array.

        'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
        Public Sub Initialize()
            'UPGRADE_WARNING: Lower bound of array udtOddsTracking was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
            ReDim udtOddsTracking(5)
        End Sub
    End Structure

    'Dim MyDb As New ADODB.Connection
    'Dim MySet As New ADODB.Recordset
    'Dim myCmd As New ADODB.Command

    Dim ToteBufferToProcess As String

    Dim CancelTimer As Double
    Dim ErrorRef As String

    Dim counter As Short
    Dim time_to_skip As Short

    'Number of worker objects to process data
    Private m_intNumWorkers As Short

    'Protocol to process
    Private m_intProtocol As Short

    Private m_strTrackCode As String

    Private m_blnOddsStatistics As Boolean

    Private m_strETX As String
    Private m_strSTX As String

    Private m_strBaud_Tote As String
    Private m_strDataBits_Tote As String
    Private m_strParity_Tote As String
    Private m_strStopBits_Tote As String

    Private m_strBaud_Timer As String
    Private m_strDataBits_Timer As String
    Private m_strParity_Timer As String
    Private m_strStopBits_Timer As String

    'Private m_strConnectionString As String

    'Array of RSIServices to process data asynchronously
    Private m_objDataWorkers() As RSIService.trpDataObject

    'Implement RSIService callback interface

    Private CommTote As New clsCommTote
    Private WithEvents Timer1 As clsTimer
    Private WithEvents Timer2 As clsTimer2

    'Private Declare API's
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)

    'returning the message that was just passed
    Public Event Message(ByRef strMessage As String)
    Public Event TimerMsg(ByRef strTimer_Msg As String)
    Public Event NewOdds(ByRef m_objOdds As RSIData.clsOdds, ByRef intRace As Short)
    'returning a flag when an RO message is received
    Public Event NewRO(ByRef m_objRunningOrder As RSIData.clsRunningOrder, ByRef intRace As Short)
    'returning a flag when the MTP change
    Public Event RaceHeaderChange(ByRef CurrentMTP As String, ByRef CurrentTime As String, ByRef CurrentPostTime As String, ByRef CurrentRace As Short)
    'Public Event MTPChange(strMTP As String, strMessageTime As String, strPostTime As String, intRace As Integer)
    Public Event TrackConditionChange(ByRef CurrentTrackCondition As String, ByRef intRace As Short)
    Public Event NewWINPool(ByRef m_objRunnerTotals As RSIData.clsRunnerTotals, ByRef intRace As Short)
    Public Event NewPlacePool(ByRef m_objRunnerTotals As RSIData.clsRunnerTotals, ByRef intRace As Short)
    Public Event NewShowPool(ByRef m_objRunnerTotals As RSIData.clsRunnerTotals, ByRef intRace As Short)
    'returning a flag when an RS message is received
    Public Event NewResults(ByRef m_objFinisherData As RSIData.clsFinisherData, ByRef intRace As Short)
    Public Event NewResultsWIN(ByRef m_objResultWPS As RSIData.clsResultWPS, ByRef intRace As Short)
    Public Event NewResultsPLC(ByRef m_objResultWPS As RSIData.clsResultWPS, ByRef intRace As Short)
    Public Event NewResultsSHW(ByRef m_objResultWPS As RSIData.clsResultWPS, ByRef intRace As Short)
    Public Event NewOfficialStatus(ByRef m_udtJudgesInfo As RSIData.typJudgesInfo, ByRef CurrentOfficialStatus As String, ByRef intRace As Short)
    'returning a flag when an TT message is received
    Public Event NewTT(ByRef m_objTeleTimer As RSIData.clsTeleTimer, ByRef intRace As Short)
    Public Event TimeOfDayChange()
    Public Event RaceStatusChange(ByRef intRace As Short)

    Private m_intCurrentRace As Short
    Private m_strCurrentMTP As String
    Private m_dtmCurrentDate As Date
    Private m_strCurrentTime As String
    ''''''''''''''''
    Private m_strCurrentPostTime As String
    Private m_strCurrentTrackCondition As String
    Private m_strCurrentStatus As String
    Private m_strCurrentRunnersFlashingStatus As String
    Private m_strCurrentOfficialStatus As String
    ''''''''''''''''

    'UPGRADE_WARNING: Lower bound of array colFinisherData was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Private colFinisherData(g_intMaxNumbOfRaces) As Collection 'One Item for each Race
    'UPGRADE_WARNING: Lower bound of array colMTPInfo was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Private colMTPInfo(g_intMaxNumbOfRaces) As Collection 'One Item for each Race
    'UPGRADE_WARNING: Lower bound of array colOdds was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Private colOdds(g_intMaxNumbOfRaces) As Collection 'One Item for each Race
    'UPGRADE_WARNING: Lower bound of array colOrderKey was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Private colOrderKey(g_intMaxNumbOfRaces) As Collection 'One Item for each Race
    'UPGRADE_WARNING: Lower bound of array colProbables was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Private colProbables(g_intMaxNumbOfRaces) As Collection 'More than one Item for each Race
    'UPGRADE_WARNING: Lower bound of array colResultExotic was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Private colResultExotic(g_intMaxNumbOfRaces) As Collection 'More than one Item for each Race
    'UPGRADE_WARNING: Lower bound of array colResultExoticAutototeV6 was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Private colResultExoticAutototeV6(g_intMaxNumbOfRaces) As Collection 'More than one Item for each Race
    'UPGRADE_WARNING: Lower bound of array colResultFullExotic was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Private colResultFullExotic(g_intMaxNumbOfRaces) As Collection 'More than one Item for each Race
    'UPGRADE_WARNING: Lower bound of array colResultPrice was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Private colResultPrice(g_intMaxNumbOfRaces) As Collection 'More than one Item for each Race
    'UPGRADE_WARNING: Lower bound of array colRunningOrder was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Private colRunningOrder(g_intMaxNumbOfRaces) As Collection 'One Item for each Race
    'UPGRADE_WARNING: Lower bound of array colPoolTotals was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Private colPoolTotals(g_intMaxNumbOfRaces) As Collection 'One Item for each Race
    'UPGRADE_WARNING: Lower bound of array colRunnerTotals was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Private colRunnerTotals(g_intMaxNumbOfRaces) As Collection 'More than one Item for each Race
    'UPGRADE_WARNING: Lower bound of array colWillPays was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Private colWillPays(g_intMaxNumbOfRaces) As Collection 'More than one Item for each Race
    'UPGRADE_WARNING: Lower bound of array colWillPaysAutototeV6 was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Private colWillPaysAutototeV6(g_intMaxNumbOfRaces) As Collection 'More than one Item for each Race
    'UPGRADE_WARNING: Lower bound of array colTeletimer was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Private colTeletimer(g_intMaxNumbOfRaces) As Collection 'More than one Item for each Race
    'Private colRunningTT(1 To g_intMaxNumbOfRaces) As Collection   'One Item for each Race
    'UPGRADE_WARNING: Lower bound of array colOddsStatistics was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Private colOddsStatistics(g_intMaxNumbOfRaces) As Collection
    Private exitRequested As Boolean

    Private m_objIniFile As clsINI_RW
    'Private m_arrMsgType(1 To g_intMaxNumMsgType) As String

    Dim m_intTimingMode As Short

    'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Class_Initialize_Renamed()
        Dim blnNewIniFile As Boolean

        'Set p_objTimerData = New TimerData 'Timming

        'Initialize Current Race.
        m_intCurrentRace = 1

        'Initialize Current MTP.
        m_strCurrentMTP = " "

        m_strCurrentTrackCondition = " "

        m_strCurrentStatus = " "
        m_strCurrentRunnersFlashingStatus = " "

        m_strCurrentPostTime = " "

        m_strCurrentOfficialStatus = " "

        'Initialize Current Date.
        m_dtmCurrentDate = Today

        'Initialize Current Time.
        'm_strCurrentTime = CStr(Time)

        'Load array message type with message header
        'loadArrayMsgType

        exitRequested = False

        'Verified that the .ini File for today exists; if not we create a new .ini file for today's date.
        m_objIniFile = New clsINI_RW
        blnNewIniFile = m_objIniFile.NewFile
        If blnNewIniFile Then
            m_objIniFile.InitFile()
        End If

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''instantiating objects'''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '  Dim oMSComm_Timer As MSComm 'Timming
        '  Set oMSComm_Timer = New MSComm 'Timming
        '  Set CommTimer.CommTimer = oMSComm_Timer 'Timming


        Dim oMSComm_Tote As MSCommLib.MSComm
        oMSComm_Tote = New MSCommLib.MSComm
        CommTote.CommTote = oMSComm_Tote

        Dim oTimer1 As ccrpTimers6.ccrpTimer
        oTimer1 = New ccrpTimers6.ccrpTimer
        Timer1 = New clsTimer
        Timer1.Timer1 = oTimer1

        Dim oTimer2 As ccrpTimers6.ccrpTimer
        oTimer2 = New ccrpTimers6.ccrpTimer
        Timer2 = New clsTimer2
        Timer2.Timer2 = oTimer2
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'counterdisplay.Show
        LogEntry("Starting Comm Server " & Now)

        time_to_skip = 1

        'Intialize number of workers
        m_intNumWorkers = 10

        'Initialize tote company to 0
        '(Autote is currently 0)
        'm_intProtocol = 1 'United
        'm_intProtocol = 2 'Amtote
        m_intProtocol = 0 'Autotote

        'set default settings for the ports
        'intitialise comm port handler for Totedata

        'With CommTote.CommTote
        '    .Settings = "9600,n,8,1"
        '    .InBufferSize = 32500
        '    .RThreshold = 1
        '    .InputLen = 1
        '    .CommPort = 1
        'End With

        LoadComPortSettings(CommTote.CommTote, "toteComPortNumber", "toteComSettings")

        Timer1.Timer1.Interval = 10
        m_strTrackCode = "NONE"
        m_blnOddsStatistics = False
        m_intTimingMode = 1
        'p_objTimerData.MessageType = 1 'Timming

        'Initialize array of data objects
        InitalizeWorkersArray()

        'If the inifile for today was already created, it could have messages. If that is
        'the case we need to load the message's objects.
        If Not blnNewIniFile Then
            LoadMessageObject()
        End If

        'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Dir("C:\RSI\StopRSIPort.tmp", FileAttribute.Normal) <> "" Then
            'remove old signal file
            Kill("C:\RSI\StopRSIPort.tmp")
        End If

        On Error GoTo PortOpenError
        'ErrorRef = "OpeningCom1" 'Timming
        'CommTimer.CommTimer.PortOpen = True 'Timming
        ErrorRef = "OpeningCom2"
        CommTote.CommTote.PortOpen = True
        'Open file data storage
        CommTote.ArchFile = FreeFile()
        FileOpen(CommTote.ArchFile, "C:\RSI\data" & VB6.Format(Today, "MMDD") & ".dat", OpenMode.Append)
        'Open timer data storage
        'CommTimer.ArchTimer = FreeFile() 'Timming
        'Open "c:\RSI\Timer" & Format$(Date, "MMDD") & ".dat" For Append As CommTimer.ArchTimer 'Timming
        On Error GoTo 0
        Timer1.Timer1.Enabled = True

        'Nov - 06, 2001 - EMJ - Added calling procedure in order to
        'be able to retrieve the properties later on.
        'LoadPortSettings_Timer CommTimer.CommTimer.Settings 'Timming
        LoadPortSettings_Tote((CommTote.CommTote.Settings))

        Exit Sub

PortOpenError:

        Select Case ErrorRef
            Case "OpeningCom1"
                LogEntry("Failed to open Com1")
            Case "OpeningCom1"
                LogEntry("Failed to open Com2")
        End Select

        Resume Next

    End Sub
    Public Sub New()
        MyBase.New()
        Class_Initialize_Renamed()
    End Sub

    'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Class_Terminate_Renamed()

        Dim intCounter As Short

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Getting rid of objects

        'Set CommTimer.CommTimer = Nothing 'Timming

        'UPGRADE_NOTE: Object CommTote.CommTote may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'Do While Marshal.ReleaseComObject(CommTote.CommTote) > 0
        'Loop
        If Not CommTote Is Nothing Then
            CommTote.CommTote = Nothing
        End If

        'UPGRADE_NOTE: Object Timer1.Timer1 may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        If Not Timer1 Is Nothing Then
            Timer1.Timer1 = Nothing
        End If

        'UPGRADE_NOTE: Object Timer2.Timer2 may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        If Not Timer2 Is Nothing Then
            Timer2.Timer2 = Nothing
        End If

        'Set CommTimer = Nothing 'Timming
        'UPGRADE_NOTE: Object CommTote may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        CommTote = Nothing

        'Set p_objTimerData = Nothing 'Timming

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        For intCounter = 1 To g_intMaxNumbOfRaces
            'UPGRADE_NOTE: Object colFinisherData() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colFinisherData(intCounter) = Nothing
            'UPGRADE_NOTE: Object colMTPInfo() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colMTPInfo(intCounter) = Nothing
            'UPGRADE_NOTE: Object colOdds() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colOdds(intCounter) = Nothing
            'UPGRADE_NOTE: Object colOrderKey() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colOrderKey(intCounter) = Nothing
            'UPGRADE_NOTE: Object colProbables() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colProbables(intCounter) = Nothing
            'UPGRADE_NOTE: Object colResultExotic() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colResultExotic(intCounter) = Nothing
            'UPGRADE_NOTE: Object colResultFullExotic() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colResultFullExotic(intCounter) = Nothing
            'UPGRADE_NOTE: Object colResultExoticAutototeV6() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colResultExoticAutototeV6(intCounter) = Nothing
            'UPGRADE_NOTE: Object colResultPrice() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colResultPrice(intCounter) = Nothing
            'UPGRADE_NOTE: Object colRunningOrder() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colRunningOrder(intCounter) = Nothing
            'UPGRADE_NOTE: Object colPoolTotals() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colPoolTotals(intCounter) = Nothing
            'UPGRADE_NOTE: Object colRunnerTotals() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colRunnerTotals(intCounter) = Nothing
            'UPGRADE_NOTE: Object colWillPays() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colWillPays(intCounter) = Nothing
            'UPGRADE_NOTE: Object colWillPaysAutototeV6() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colWillPaysAutototeV6(intCounter) = Nothing
            'UPGRADE_NOTE: Object colTeletimer() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colTeletimer(intCounter) = Nothing
        Next intCounter

        'UPGRADE_NOTE: Object m_objIniFile may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objIniFile = Nothing

        '--- Destroy objects created
        For intCounter = 0 To UBound(m_objDataWorkers)
            'UPGRADE_NOTE: Object m_objDataWorkers() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            m_objDataWorkers(intCounter) = Nothing
        Next intCounter

        'Close all open files
        FileClose()

    End Sub
    Protected Overrides Sub Finalize()
        Class_Terminate_Renamed()
        MyBase.Finalize()
    End Sub

    Public Property CancelProcess() As Boolean
        Get
            Return exitRequested
        End Get

        Set(value As Boolean)
            exitRequested = value
        End Set
    End Property

    Public Sub CloseAll()
        'Nov-15, 2001 - EMJ - Work around, because the clas was not going through the terminate event
        'had to force it.
        Class_Terminate_Renamed()
    End Sub

    Public ReadOnly Property ErrorMessage() As String
        Get
            ErrorMessage = myError
        End Get
    End Property

    Public ReadOnly Property strBaud_Rate_Tote() As String
        Get
            strBaud_Rate_Tote = m_strBaud_Tote
        End Get
    End Property

    Public ReadOnly Property strDataBits_Tote() As String
        Get
            strDataBits_Tote = m_strDataBits_Tote
        End Get
    End Property

    Public ReadOnly Property strParity_Tote() As String
        Get
            strParity_Tote = m_strParity_Tote
        End Get
    End Property

    Public ReadOnly Property strStopBits_Tote() As String
        Get
            strStopBits_Tote = m_strStopBits_Tote
        End Get
    End Property

    Public ReadOnly Property strBaud_Rate_Timer() As String
        Get
            strBaud_Rate_Timer = m_strBaud_Timer
        End Get
    End Property

    Public ReadOnly Property strDataBits_Timer() As String
        Get
            strDataBits_Timer = m_strDataBits_Timer
        End Get
    End Property

    Public ReadOnly Property strParity_Timer() As String
        Get
            strParity_Timer = m_strParity_Timer
        End Get
    End Property

    Public ReadOnly Property strStopBits_Timer() As String
        Get
            strStopBits_Timer = m_strStopBits_Timer
        End Get
    End Property

    'Nov-14, 2001 - EMJ - added property in order to get the interval
    'that is set for the timer
    Public ReadOnly Property GetTimerInterval() As Short
        Get
            On Error GoTo ErrHndlr

            GetTimerInterval = Timer1.Timer1.Interval

            Exit Property

ErrHndlr:
            'Implement error logging
        End Get
    End Property

    'Read Only properties
    Public ReadOnly Property CurrentRace() As Short
        Get
            On Error GoTo ErrHndlr

            CurrentRace = m_intCurrentRace

            Exit Property

ErrHndlr:
            'Implement error logging
        End Get
    End Property

    Public ReadOnly Property CurrentMTP() As String
        Get
            On Error GoTo ErrHndlr

            CurrentMTP = m_strCurrentMTP

            Exit Property

ErrHndlr:
            'Implement error logging
        End Get
    End Property

    Public ReadOnly Property CurrentTrackCondition() As String
        Get
            On Error GoTo ErrHndlr

            CurrentTrackCondition = m_strCurrentTrackCondition

            Exit Property

ErrHndlr:
            'Implement error logging
        End Get
    End Property

    Public ReadOnly Property CurrentStatus() As String
        Get
            On Error GoTo ErrHndlr

            CurrentStatus = m_strCurrentStatus

            Exit Property

ErrHndlr:
            'Implement error logging
        End Get
    End Property

    Public ReadOnly Property CurrentRunnersFlashingStatus() As String
        Get
            On Error GoTo ErrHndlr

            CurrentRunnersFlashingStatus = m_strCurrentRunnersFlashingStatus
            Exit Property

ErrHndlr:
            'Implement error logging
        End Get
    End Property

    Public ReadOnly Property CurrentPostTime() As String
        Get
            On Error GoTo ErrHndlr

            CurrentPostTime = m_strCurrentPostTime

            Exit Property

ErrHndlr:
            'Implement error logging
        End Get
    End Property

    Public ReadOnly Property CurrentDate() As Date
        Get
            On Error GoTo ErrHndlr

            CurrentDate = m_dtmCurrentDate

            Exit Property

ErrHndlr:
            'Implement error logging
        End Get
    End Property

    Public ReadOnly Property CurrentTime() As String
        Get
            On Error GoTo ErrHndlr

            CurrentTime = m_strCurrentTime

            Exit Property

ErrHndlr:
            'Implement error logging
        End Get
    End Property

    Public ReadOnly Property MTPByRace(ByVal intRaceNum As Short) As String
        Get
            Dim objMTP As RSIData.typMTPInfo

            If intRaceNum >= 1 And intRaceNum <= 20 Then
                If Not colMTPInfo(intRaceNum) Is Nothing Then
                    If ItemExists(colMTPInfo(intRaceNum), "MTP" & intRaceNum) Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object colMTPInfo().Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object objMTP. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        objMTP = colMTPInfo(intRaceNum).Item("MTP" & intRaceNum)
                        MTPByRace = objMTP.strMTP
                    End If
                End If
            End If

        End Get
    End Property

    Public ReadOnly Property ObjectMTP(ByVal intRaceNum As Short) As Collection
        Get
            ObjectMTP = New Collection
            If intRaceNum >= 1 And intRaceNum <= 20 Then
                If Not colMTPInfo(intRaceNum) Is Nothing Then
                    If ItemExists(colMTPInfo(intRaceNum), "MTP" & intRaceNum) Then
                        ObjectMTP.Add(colMTPInfo(intRaceNum).Item("MTP" & intRaceNum), "MTP" & intRaceNum)
                    End If
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property ObjectRunningOrder(ByVal intRaceNum As Short) As Collection
        Get
            ObjectRunningOrder = New Collection
            If intRaceNum >= 1 And intRaceNum <= 20 Then
                If Not colRunningOrder(intRaceNum) Is Nothing Then
                    If ItemExists(colRunningOrder(intRaceNum), "RO" & intRaceNum) Then
                        ObjectRunningOrder.Add(colRunningOrder(intRaceNum).Item("RO" & intRaceNum), "RO" & intRaceNum)
                    End If
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property ObjectRsExoFull(ByVal intRaceNum As Short) As Collection
        Get
            ObjectRsExoFull = New Collection
            If intRaceNum >= 1 And intRaceNum <= 20 Then
                If Not colResultFullExotic(intRaceNum) Is Nothing Then
                    If ItemExists(colResultFullExotic(intRaceNum), "RF" & intRaceNum) Then
                        ObjectRsExoFull.Add(colResultFullExotic(intRaceNum).Item("RF" & intRaceNum), "RF" & intRaceNum)
                    End If
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property ObjectPoolTotal(ByVal intRaceNum As Short) As Collection
        Get
            ObjectPoolTotal = New Collection
            If intRaceNum >= 1 And intRaceNum <= 20 Then
                If Not colPoolTotals(intRaceNum) Is Nothing Then
                    If ItemExists(colPoolTotals(intRaceNum), "PT" & intRaceNum) Then
                        ObjectPoolTotal.Add(colPoolTotals(intRaceNum).Item("PT" & intRaceNum), "PT" & intRaceNum)
                    End If
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property ObjectOrderKey(ByVal intRaceNum As Short) As Collection
        Get
            ObjectOrderKey = New Collection
            If intRaceNum >= 1 And intRaceNum <= 20 Then
                If Not colOrderKey(intRaceNum) Is Nothing Then
                    If ItemExists(colOrderKey(intRaceNum), "OK" & intRaceNum) Then
                        ObjectOrderKey.Add(colOrderKey(intRaceNum).Item("OK" & intRaceNum), "OK" & intRaceNum)
                    End If
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property ObjectOdds(ByVal intRaceNum As Short) As Collection
        Get
            ObjectOdds = New Collection
            If intRaceNum >= 1 And intRaceNum <= 20 Then
                If Not colOdds(intRaceNum) Is Nothing Then
                    If ItemExists(colOdds(intRaceNum), "Odds" & intRaceNum) Then
                        ObjectOdds.Add(colOdds(intRaceNum).Item("Odds" & intRaceNum), "Odds" & intRaceNum)
                    End If
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property ObjectPB(ByVal intRaceNum As Short, ByVal strKey As String) As Collection
        Get
            'strKey = m_objProbables.BetAmount & m_objProbables.PoolCode
            'strKey = m_objProbables.BetAmount & m_objProbables.PoolCode & m_objProbables.NumberOfRows
            ObjectPB = New Collection
            If intRaceNum >= 1 And intRaceNum <= 20 Then
                If Not colProbables(intRaceNum) Is Nothing Then
                    If ItemExists(colProbables(intRaceNum), strKey) Then
                        ObjectPB.Add(colProbables(intRaceNum).Item(strKey), strKey)
                    End If
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property ObjectFinisher(ByVal intRaceNum As Short) As Collection
        Get
            Dim strKey As String

            strKey = "F" & intRaceNum
            ObjectFinisher = New Collection
            If intRaceNum >= 1 And intRaceNum <= 20 Then
                If Not colFinisherData(intRaceNum) Is Nothing Then
                    If ItemExists(colFinisherData(intRaceNum), strKey) Then
                        ObjectFinisher.Add(colFinisherData(intRaceNum).Item(strKey), strKey)
                    End If
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property ObjectRsExo(ByVal intRaceNum As Short, ByVal strKey As String) As Collection
        Get
            'strKey = m_objResultExotic.BetAmount & m_objResultExotic.PoolCode
            ObjectRsExo = New Collection
            If intRaceNum >= 1 And intRaceNum <= 20 Then
                If Not colResultExotic(intRaceNum) Is Nothing Then
                    If ItemExists(colResultExotic(intRaceNum), strKey) Then
                        ObjectRsExo.Add(colResultExotic(intRaceNum).Item(strKey), strKey)
                    End If
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property ObjectRsExoAutototeV6(ByVal intRaceNum As Short, ByVal strKey As String) As Collection
        Get
            'strKey = m_objResultExotic.BetAmount & m_objResultExotic.PoolCode
            ObjectRsExoAutototeV6 = New Collection
            If intRaceNum >= 1 And intRaceNum <= 20 Then
                If Not colResultExoticAutototeV6(intRaceNum) Is Nothing Then
                    If ItemExists(colResultExoticAutototeV6(intRaceNum), strKey) Then
                        ObjectRsExoAutototeV6.Add(colResultExoticAutototeV6(intRaceNum).Item(strKey), strKey)
                    End If
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property ObjectRsWPS(ByVal intRaceNum As Short, ByVal strKey As String) As Collection
        Get
            'strKey = m_objResultWPS.PoolCode
            ObjectRsWPS = New Collection
            If intRaceNum >= 1 And intRaceNum <= 20 Then
                If Not colResultPrice(intRaceNum) Is Nothing Then
                    If ItemExists(colResultPrice(intRaceNum), strKey) Then
                        ObjectRsWPS.Add(colResultPrice(intRaceNum).Item(strKey), strKey)
                    End If
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property ObjectRunnerTotals(ByVal intRaceNum As Short, ByVal strKey As String) As Collection
        Get
            'strKey = m_objRunnerTotals.PoolCode
            ObjectRunnerTotals = New Collection
            If intRaceNum >= 1 And intRaceNum <= 20 Then
                If Not colRunnerTotals(intRaceNum) Is Nothing Then
                    If ItemExists(colRunnerTotals(intRaceNum), strKey) Then
                        ObjectRunnerTotals.Add(colRunnerTotals(intRaceNum).Item(strKey), strKey)
                    End If
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property ObjectWP(ByVal intRaceNum As Short, ByVal strKey As String) As Collection
        Get
            'strKey = m_objProbables.BetAmount & m_objProbables.PoolCode
            'strKey = m_objProbables.BetAmount & m_objProbables.PoolCode & m_objProbables.NumberOfRows
            ObjectWP = New Collection
            If intRaceNum >= 1 And intRaceNum <= 20 Then
                If Not colWillPays(intRaceNum) Is Nothing Then
                    If ItemExists(colWillPays(intRaceNum), strKey) Then
                        ObjectWP.Add(colWillPays(intRaceNum).Item(strKey), strKey)
                    End If
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property ObjectWPAutototeV6(ByVal intRaceNum As Short, ByVal strKey As String) As Collection
        Get
            'strKey = m_objProbables.BetAmount & m_objProbables.PoolCode
            'strKey = m_objProbables.BetAmount & m_objProbables.PoolCode & m_objProbables.NumberOfRows
            ObjectWPAutototeV6 = New Collection
            If intRaceNum >= 1 And intRaceNum <= 20 Then
                If Not colWillPaysAutototeV6(intRaceNum) Is Nothing Then
                    If ItemExists(colWillPaysAutototeV6(intRaceNum), strKey) Then
                        ObjectWPAutototeV6.Add(colWillPaysAutototeV6(intRaceNum).Item(strKey), strKey)
                    End If
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property ObjectTT(ByVal intRaceNum As Short) As Collection
        Get
            ObjectTT = New Collection
            If intRaceNum >= 1 And intRaceNum <= 20 Then
                If Not colTeletimer(intRaceNum) Is Nothing Then
                    If ItemExists(colTeletimer(intRaceNum), "TT" & intRaceNum) Then
                        ObjectTT.Add(colTeletimer(intRaceNum).Item("TT" & intRaceNum), "TT" & intRaceNum)
                    End If
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property HiOddByRunner(ByVal intRaceNum As Short, ByVal intRunner As Short) As String
        Get
            'Dim strKey As String
            '
            '  strKey = "R" & intRunner
            '  If intRaceNum >= 1 And intRaceNum <= 20 Then
            '    If Not colOddsStatistics(intRaceNum) Is Nothing Then
            '      If ItemExists(colOddsStatistics(intRaceNum), strKey) Then
            '        HiOddByRunner = colOddsStatistics(intRaceNum).Item(strKey).strHi
            '      End If
            '    End If
            '  End If
        End Get
    End Property

    Public ReadOnly Property LoOddByRunner(ByVal intRaceNum As Short, ByVal intRunner As Short) As String
        Get
            'Dim strKey As String
            '
            '  strKey = "R" & intRunner
            '  If intRaceNum >= 1 And intRaceNum <= 20 Then
            '    If Not colOddsStatistics(intRaceNum) Is Nothing Then
            '      If ItemExists(colOddsStatistics(intRaceNum), strKey) Then
            '        LoOddByRunner = colOddsStatistics(intRaceNum).Item(strKey).strLo
            '      End If
            '    End If
            '  End If
        End Get
    End Property

    Public ReadOnly Property OddTrackingByRunner(ByVal intRaceNum As Short, ByVal intRunner As Short, ByVal intOddCtr As Short) As String
        Get
            'Dim strKey As String
            '
            '  strKey = "R" & intRunner
            '  If intRaceNum >= 1 And intRaceNum <= 20 Then
            '    If Not colOddsStatistics(intRaceNum) Is Nothing Then
            '      If ItemExists(colOddsStatistics(intRaceNum), strKey) Then
            '        If intOddCtr >= 1 And intOddCtr <= 5 Then
            '          OddTrackingByRunner = colOddsStatistics(intRaceNum).Item(strKey).udtOddsTracking(intOddCtr).strOdds
            '        End If
            '      End If
            '    End If
            '  End If
        End Get
    End Property

    Public WriteOnly Property PathIniFile() As String
        Set(ByVal Value As String)
            m_objIniFile.PathIniFile = Value
        End Set
    End Property

    'Public Properties

    Public Property Protocol() As Short
        Get
            Protocol = m_intProtocol
        End Get
        Set(ByVal Value As Short)
            Dim intCtr As Short

            m_intProtocol = Value
            For intCtr = 0 To m_intNumWorkers - 1
                'Reset protocol
                m_objDataWorkers(intCtr).Protocol = m_intProtocol
                'MsgBox m_intProtocol
            Next intCtr

        End Set
    End Property

    Public ReadOnly Property GetMessageToSend(ByVal strMsg As String) As String
        Get
            Dim intCtr As Short
            Dim strTemp As String

            On Error GoTo ErrHndlr

            For intCtr = 1 To Len(strMsg) Step 2
                strTemp = Mid(strMsg, intCtr, 2)
                GetMessageToSend = GetMessageToSend & Chr(CByte("&H" & strTemp))
            Next intCtr

            Exit Property

ErrHndlr:
            'Implement error logging

        End Get
    End Property

    'Initialize array of workers and initalize each object
    Private Sub InitalizeWorkersArray()
        Dim intCounter As Short

        On Error GoTo ErrHndlr

        ReDim m_objDataWorkers(m_intNumWorkers - 1)

        For intCounter = 0 To m_intNumWorkers - 1
            'Intialize object
            m_objDataWorkers(intCounter) = New RSIService.trpDataObject

            'Set protocol
            m_objDataWorkers(intCounter).Protocol = m_intProtocol

            'set the trackcode for tracks that don't send it (EM)
            m_objDataWorkers(intCounter).TrackCode = m_strTrackCode

            'Set call back object
            m_objDataWorkers(intCounter).Notify = Me
        Next intCounter

        'Get protocols STX and ETX
        If Not m_objDataWorkers(0) Is Nothing Then
            m_strSTX = m_objDataWorkers(0).STX
            m_strETX = m_objDataWorkers(0).ETX
        End If

        Exit Sub

ErrHndlr:
        LogEntry("Error initializing object array " & Err.Description)
    End Sub

    Private Sub ServiceCommTimer()
        'Timming
        'Dim STXpsn As Integer
        'Dim ETXpsn As Integer
        'Dim BuffLen As Integer
        'Dim intPost As Integer
        '
        'On Error GoTo TimerError
        '
        '  CommTimer.TimerBuffer = CommTimer.TimerBuffer & CommTimer.CommTimer.Input
        '
        '  'test for start clock
        '  intPost = InStr(CommTimer.TimerBuffer, Chr(26) & Chr(48) & Chr(28))
        '  If intPost > 0 Then
        '    RaiseEvent StartTimerClock(True)
        '    CommTimer.TimerBuffer = Mid(CommTimer.TimerBuffer, intPost + 3)
        '  End If
        '
        '  'test for possible message
        '  STXpsn = InStr(CommTimer.TimerBuffer, Chr$(1))
        '  If STXpsn > 0 Then
        '    BuffLen = Len(CommTimer.TimerBuffer)
        '    If STXpsn > 1 Then
        '      'stx present - strip preceeding data
        '      CommTimer.TimerBuffer = Right$(CommTimer.TimerBuffer, BuffLen - STXpsn + 1)
        '    End If
        '    'test for etx
        '    BuffLen = Len(CommTimer.TimerBuffer)
        '    ETXpsn = InStr(CommTimer.TimerBuffer, Chr$(4))
        '    If ETXpsn > 0 Then
        '      'message available
        '      p_objTimerData.ProcTimerData Left$(CommTimer.TimerBuffer, ETXpsn)
        '      RaiseEvent TimerMsg(Left$(CommTimer.TimerBuffer, ETXpsn))
        '      CommTimer.TimerBuffer = Right$(CommTimer.TimerBuffer, BuffLen - ETXpsn)
        '    End If
        '  End If
        '
        '  On Error GoTo 0
        '
        '  Exit Sub
        '
        'TimerError:
        '  On Error GoTo 0
    End Sub

    'New proc implementing the new parser architechture
    Private Sub serviceCommTote()
        Dim intWorker As Short
        Dim intETX As Short
        Dim strMessage As String
        Dim intNumTries As Short

        On Error GoTo ErrHndlr

        ToteBufferToProcess = ToteBufferToProcess & CommTote.ToteBuffer
        CommTote.ToteBuffer = ""

        'See if end of transmission has been received
        intETX = InStr(ToteBufferToProcess, m_strETX)

        'Process info if ETX found
        If intETX > 0 Then
            strMessage = Left(ToteBufferToProcess, intETX)
            RaiseEvent Message(strMessage)
            'Find a worker who is not busy
            'Try four times before giving up..If no worker is free
            'leave the buffer alone and wait for next timer "click"
            'to process the data
            For intNumTries = 0 To 3
                For intWorker = 0 To m_intNumWorkers - 1
                    If Not m_objDataWorkers(intWorker).Busy Then
                        'give the worker the data and keep on working
                        m_objDataWorkers(intWorker).ProcessSerialData(strMessage)
                        'Keep data received after ETX
                        ToteBufferToProcess = Mid(ToteBufferToProcess, intETX + 1)
                        Exit Sub
                    End If
                Next intWorker
                'wait a hundredth of a second before triyng again
                TimeDelay(100)
            Next intNumTries
        End If

        Exit Sub

ErrHndlr:
        LogEntry("Error processing tote message in serviceCommTote " & CStr(Err.Number) & " " & Err.Description)
    End Sub

    'Notification interface mechanism for RSIServices
    Private Sub IRaceFxNotify_DataObjectError(ByRef lngErrorNumber As Integer, ByRef strErrorDescription As String, ByRef intParsingClass As Short) Implements RSIService.IRaceFxNotify.DataObjectError

        On Error GoTo ErrHndlr

        'Log error
        LogEntry("Error processing message... " & CStr(lngErrorNumber) & " " & strErrorDescription & " " & CStr(intParsingClass))

        Exit Sub

ErrHndlr:
        'Implement error logging
    End Sub

    Private Sub IRaceFxNotify_DataObjectJudgesInfo(ByRef m_udtJudgesInfo As RSIData.typJudgesInfo, ByVal intRaceNum As Short) Implements RSIService.IRaceFxNotify.DataObjectJudgesInfo
        Dim strTemp As String

        On Error GoTo ErrHndlr

        strTemp = ""

        If (m_udtJudgesInfo.blnOfficialStatus) Then
            strTemp = "OFFICIAL"
        ElseIf (m_udtJudgesInfo.blnObjStatus) Then
            strTemp = "Objection"
        ElseIf (m_udtJudgesInfo.blnInqStatus) Then
            strTemp = "Inquiry"
        ElseIf (m_udtJudgesInfo.blnPhotoStatus) Then
            strTemp = "Photo"
        Else
            strTemp = ""
        End If

        If (m_strCurrentOfficialStatus <> strTemp) Then
            m_strCurrentOfficialStatus = strTemp
            RaiseEvent NewOfficialStatus(m_udtJudgesInfo, m_strCurrentOfficialStatus, intRaceNum)
        End If

        Exit Sub

ErrHndlr:
        'Implement error logging
    End Sub

    Private Sub IRaceFxNotify_DataObjectRaceStatus(ByRef m_udtRaceStatus As RSIData.typRaceStatus, ByVal intRaceNum As Short) Implements RSIService.IRaceFxNotify.DataObjectRaceStatus

        On Error GoTo ErrHndlr

        If (m_strCurrentTrackCondition <> m_udtRaceStatus.strTrackCondition) Then
            m_strCurrentTrackCondition = m_udtRaceStatus.strTrackCondition
            RaiseEvent TrackConditionChange(m_strCurrentTrackCondition, intRaceNum)
        End If

        m_strCurrentTrackCondition = m_udtRaceStatus.strTrackCondition

        Exit Sub

ErrHndlr:
        'Implement error logging

    End Sub

    Private Sub IRaceFxNotify_DataObjectSignStatus(ByRef m_udtSignStatus As RSIData.typSignStatus, ByVal intRaceNum As Short) Implements RSIService.IRaceFxNotify.DataObjectSignStatus

        On Error GoTo ErrHndlr

        'If (m_strCurrentStatus <> m_udtSignStatus.strStatus) Then
        m_strCurrentStatus = m_udtSignStatus.strStatus
        m_strCurrentRunnersFlashingStatus = m_udtSignStatus.strRunners
        RaiseEvent RaceStatusChange(intRaceNum)
        'End If
        'm_strCurrentStatus = m_udtSignStatus.strStatus
        'm_strCurrentRunnersFlashingStatus = m_udtSignStatus.strRunners

        Exit Sub

ErrHndlr:
        'Implement error loggingers
    End Sub

    Private Sub IRaceFxNotify_DataObjectResultFullExotic(ByVal m_objResultFullExotic As RSIData.clsFullExotics, ByVal intRaceNum As Short) Implements RSIService.IRaceFxNotify.DataObjectResultFullExotic
        Dim strKey As String
        Dim colTemp As Collection
        On Error GoTo ErrHndlr

        strKey = "RF" & intRaceNum

        colTemp = New Collection
        colTemp.Add(m_objResultFullExotic, strKey)

        colResultFullExotic(intRaceNum) = colTemp

        'UPGRADE_NOTE: Object colTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        colTemp = Nothing

        Exit Sub

ErrHndlr:
        'Implement error loggingers
    End Sub

    Private Sub IRaceFxNotify_DataObjectRunningOrder(ByVal m_objRunningOrder As RSIData.clsRunningOrder, ByVal intRaceNum As Short) Implements RSIService.IRaceFxNotify.DataObjectRunningOrder
        Dim strKey As String
        Dim colTemp As Collection

        On Error GoTo ErrHndlr

        'Code needed for Amtote
        '  If intRaceNum = g_intMaxNumbOfRaces Then
        '    intRaceNum = m_intCurrentRace
        '    m_objRunningOrder.Race = intRaceNum
        '  End If

        '  strKey = "RO" & intRaceNum
        '
        '  Set colTemp = New Collection
        '
        '  colTemp.Add m_objRunningOrder, strKey
        '  Set colRunningOrder(intRaceNum) = colTemp
        '
        '  Set colTemp = Nothing

        'Set colRunningOrder(intRaceNum) = New Collection
        'colRunningOrder(intRaceNum).Add m_objRunningOrder, "RO" & intRaceNum

        RaiseEvent NewRO(m_objRunningOrder, intRaceNum)

        Exit Sub

ErrHndlr:
        'Implement error logging
    End Sub

    Private Sub IRaceFxNotify_DataObjectPoolTotals(ByVal m_objPoolTotals As RSIData.clsPoolTotals, ByVal intRaceNum As Short) Implements RSIService.IRaceFxNotify.DataObjectPoolTotals
        Dim strKey As String
        Dim colTemp As Collection

        On Error GoTo ErrHndlr

        'Code needed for Amtote
        If intRaceNum = g_intMaxNumbOfRaces Then
            intRaceNum = m_intCurrentRace
            m_objPoolTotals.Race = intRaceNum
        End If

        strKey = "PT" & intRaceNum

        colTemp = New Collection

        colTemp.Add(m_objPoolTotals, strKey)
        colPoolTotals(intRaceNum) = colTemp

        'UPGRADE_NOTE: Object colTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        colTemp = Nothing

        Exit Sub

ErrHndlr:
        'Implement error logging
    End Sub

    Private Sub IRaceFxNotify_DataObjectOrderKey(ByVal m_objOrderKey As RSIData.clsRunningOrder, ByVal intRaceNum As Short) Implements RSIService.IRaceFxNotify.DataObjectOrderKey
        Dim strKey As String
        Dim colTemp As Collection

        On Error GoTo ErrHndlr

        'Code needed for Amtote
        If intRaceNum = g_intMaxNumbOfRaces Then
            intRaceNum = m_intCurrentRace
            m_objOrderKey.Race = intRaceNum
        End If

        strKey = "OK" & intRaceNum

        colTemp = New Collection

        colTemp.Add(m_objOrderKey, strKey)
        colOrderKey(intRaceNum) = colTemp

        'UPGRADE_NOTE: Object colTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        colTemp = Nothing

        'Set colOrderKey(intRaceNum) = New Collection
        'colOrderKey(intRaceNum).Add m_objOrderKey, "OK" & intRaceNum

        Exit Sub

ErrHndlr:
        'Implement error logging
    End Sub

    Private Sub IRaceFxNotify_DataObjectMTPInfo(ByRef m_udtMTPInfo As RSIData.typMTPInfo, ByVal intRaceNum As Short) Implements RSIService.IRaceFxNotify.DataObjectMTPInfo
        Dim strKey As String
        Dim colTemp As Collection
        Static intRaceNumber As Short

        On Error GoTo ErrHndlr

        '  m_intCurrentRace = intRaceNum
        '  m_strCurrentMTP = m_udtMTPInfo.strMTP
        '  m_dtmCurrentDate = m_udtMTPInfo.dtmCurrentDate
        '  m_strCurrentTime = m_udtMTPInfo.strMessageTime
        '  m_strCurrentPostTime = m_udtMTPInfo.strPostTime

        '  strKey = "MTP" & intRaceNum
        '
        '  Set colTemp = New Collection
        '
        '  colTemp.Add m_udtMTPInfo, strKey
        '  Set colMTPInfo(intRaceNum) = colTemp
        '
        '  Set colTemp = Nothing

        'Set colMTPInfo(intRaceNum) = New Collection
        'colMTPInfo(intRaceNum).Add m_udtMTPInfo, "MTP" & intRaceNum

        'If ((intRaceNumber <> intRaceNum) Or ((IsNumeric(m_strCurrentMTP)) And (m_strCurrentMTP <= 4))) Then
        '  If intRaceNumber <> intRaceNum Then
        '    RaiseEvent RaceChange(intRaceNum)
        '    intRaceNumber = intRaceNum
        '    Exit Sub
        '  End If

        '  If IsNumeric(m_strCurrentMTP) Then
        '    If ((m_strCurrentMTP <= 4) And (m_strCurrentMTP > 0)) Then
        '      RaiseEvent RaceChange(intRaceNum)
        '    End If
        '  End If

        RaiseEvent RaceHeaderChange((m_udtMTPInfo.strMTP), (m_udtMTPInfo.strMessageTime), (m_udtMTPInfo.strPostTime), intRaceNum)

        Exit Sub

ErrHndlr:
        'Implement error logging
    End Sub

    Private Sub IRaceFxNotify_DataObjectOdds(ByVal m_objOdds As RSIData.clsOdds, ByVal intRaceNum As Short) Implements RSIService.IRaceFxNotify.DataObjectOdds
        Dim intCtr As Short
        Dim intCtr2 As Short
        Dim sngObjOdd As Single
        Dim sngudtOdd As Single
        'UPGRADE_WARNING: Arrays in structure udtOddsStatistics may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
        Dim udtOddsStatistics As typOddsStatistics
        Dim strFileName As String
        Dim strFileToOpen As String
        Dim strData As String
        Dim strKey As String
        Dim colTemp As Collection

        On Error GoTo ErrHndlr

        '  strKey = "Odds" & intRaceNum
        '
        '  Set colTemp = New Collection
        '
        '  colTemp.Add m_objOdds, strKey
        '  Set colOdds(intRaceNum) = colTemp
        '
        '  Set colTemp = Nothing

        'Set colOdds(intRaceNum) = New Collection
        'colOdds(intRaceNum).Add m_objOdds, "Odds" & intRaceNum

        '  If m_blnOddsStatistics Then
        '    strFileName = Format$(Now(), "mm_dd_YYYY")
        '    strFileToOpen = "C:\ToteBoard\INI\" & strFileName & ".ini"
        '    If colOddsStatistics(intRaceNum) Is Nothing Then
        '      Set colOddsStatistics(intRaceNum) = New Collection
        '    End If
        '    For intCtr = 1 To m_objOdds.NumberOfRows
        '      udtOddsStatistics.strHi = ""
        '      udtOddsStatistics.strLo = ""
        '      Erase udtOddsStatistics.udtOddsTracking
        '      If ItemExists(colOddsStatistics(intRaceNum), "R" & intCtr) Then
        '        udtOddsStatistics = colOddsStatistics(intRaceNum).Item("R" & intCtr)
        '        colOddsStatistics(intRaceNum).Remove "R" & intCtr
        '      End If
        '      If (Val(m_objOdds.OddsByRunner(intCtr)) <> 0) Then
        '        sngObjOdd = CFraction(m_objOdds.OddsByRunner(intCtr))
        '        sngudtOdd = CFraction(udtOddsStatistics.strHi)
        '        If InStr(1, m_objOdds.OddsByRunner(intCtr), "99") = 0 Then
        '          If sngudtOdd < sngObjOdd Then
        '            udtOddsStatistics.strHi = m_objOdds.OddsByRunner(intCtr)
        '          End If
        '        End If
        '        sngudtOdd = CFraction(udtOddsStatistics.strLo)
        '        If ((sngudtOdd = 0) Or (sngudtOdd > sngObjOdd)) Then
        '          udtOddsStatistics.strLo = m_objOdds.OddsByRunner(intCtr)
        '        End If
        '      End If
        '      For intCtr2 = 1 To 4
        '        udtOddsStatistics.udtOddsTracking(intCtr2).strOdds = udtOddsStatistics.udtOddsTracking(intCtr2 + 1).strOdds
        '      Next intCtr2
        '      udtOddsStatistics.udtOddsTracking(intCtr2).strOdds = m_objOdds.OddsByRunner(intCtr)
        '      strData = "HL|" & intRaceNum & "|" & intCtr & "|HI|" & udtOddsStatistics.strHi & _
        ''        "|LO|" & udtOddsStatistics.strLo
        '      m_objIniFile.WriteData "HL" & intRaceNum & "_" & intCtr, "message", strData, strFileToOpen
        '      colOddsStatistics(intRaceNum).Add udtOddsStatistics, "R" & intCtr
        '    Next intCtr
        '  End If

        RaiseEvent NewOdds(m_objOdds, intRaceNum)

        Exit Sub

ErrHndlr:
        'Implement error logging
    End Sub

    Private Function CFraction(ByRef strData As String) As Single
        Dim intPost As Short
        Dim intNumerator As Short
        Dim intDenominator As Short

        On Error GoTo ErrHndlr

        intPost = InStr(strData, "/")
        'Amtote send the fractions as Ex.: 2-5
        intPost = IIf((intPost = 0), InStr(strData, "-"), intPost)
        If intPost > 0 Then
            intNumerator = Val(Left(strData, intPost - 1))
            intDenominator = Val(Mid(strData, intPost + 1))
            If intDenominator <> 0 Then
                CFraction = intNumerator / intDenominator
                Exit Function
            End If
        End If

        CFraction = Val(strData)

        Exit Function

ErrHndlr:
        'Implement error logging
    End Function

    Private Sub IRaceFxNotify_DataObjectProbables(ByVal m_objProbables As RSIData.clsProbables, ByVal intRaceNum As Short) Implements RSIService.IRaceFxNotify.DataObjectProbables
        'Dim strKey As String
        '
        '  'The numberOfRows is equal to the highest horse family number, so that if we
        '  'only have one runner in the message (OneRow=true) the m_objProbables.NumberOfRows
        '  'will tell us this runner number.
        '  If m_objProbables.OneRow Then
        '    strKey = intRaceNum & "_" & m_objProbables.BetAmount & m_objProbables.PoolCode & "_" & m_objProbables.NumberOfRows
        '  Else
        '    strKey = intRaceNum & "_" & m_objProbables.BetAmount & m_objProbables.PoolCode
        '  End If
        '
        '  If colProbables(intRaceNum) Is Nothing Then
        '    Set colProbables(intRaceNum) = New Collection
        '  Else
        '    If ItemExists(colProbables(intRaceNum), strKey) Then
        '      DoEvents
        '      colProbables(intRaceNum).Remove strKey
        '    End If
        '  End If
        '
        '  colProbables(intRaceNum).Add m_objProbables, strKey
    End Sub

    Private Sub IRaceFxNotify_DataObjectFinisherData(ByVal m_objFinisherData As RSIData.clsFinisherData, ByVal intRaceNum As Short) Implements RSIService.IRaceFxNotify.DataObjectFinisherData
        Dim strKey As String
        Dim colTemp As Collection

        On Error GoTo ErrHndlr

        strKey = "F" & intRaceNum
        colTemp = New Collection

        colTemp.Add(m_objFinisherData, strKey)
        colFinisherData(intRaceNum) = colTemp

        'UPGRADE_NOTE: Object colTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        colTemp = Nothing

        RaiseEvent NewResults(m_objFinisherData, intRaceNum)

        colFinisherData(intRaceNum) = New Collection
        colFinisherData(intRaceNum).Add(m_objFinisherData, "F" & intRaceNum)

        Exit Sub

ErrHndlr:
        'Implement error logging
    End Sub

    Private Sub IRaceFxNotify_DataObjectResultExotic(ByVal m_objResultExotic As RSIData.clsResultExotic, ByVal intRaceNum As Short) Implements RSIService.IRaceFxNotify.DataObjectResultExotic
        Dim strKey As String

        On Error GoTo ErrHndlr

        strKey = intRaceNum & "_" & m_objResultExotic.PoolCode

        If colResultExotic(intRaceNum) Is Nothing Then
            colResultExotic(intRaceNum) = New Collection
        Else
            If ItemExists(colResultExotic(intRaceNum), strKey) Then
                colResultExotic(intRaceNum).Remove(strKey)
            End If
        End If

        colResultExotic(intRaceNum).Add(m_objResultExotic, strKey)

        strKey = intRaceNum & "_" & m_objResultExotic.PoolCode

        If colResultExoticAutototeV6(intRaceNum) Is Nothing Then
            colResultExoticAutototeV6(intRaceNum) = New Collection
        Else
            If ItemExists(colResultExoticAutototeV6(intRaceNum), strKey) Then
                colResultExoticAutototeV6(intRaceNum).Remove(strKey)
            End If
        End If

        colResultExoticAutototeV6(intRaceNum).Add(m_objResultExotic, strKey)

        Exit Sub

ErrHndlr:
        'Implement error logging
    End Sub

    Private Sub IRaceFxNotify_DataObjectResultPrice(ByVal m_objResultWPS As RSIData.clsResultWPS, ByVal intRaceNum As Short) Implements RSIService.IRaceFxNotify.DataObjectResultPrice
        Dim strKey As String

        On Error GoTo ErrHndlr

        strKey = intRaceNum & "_" & m_objResultWPS.PoolCode

        If colResultPrice(intRaceNum) Is Nothing Then
            colResultPrice(intRaceNum) = New Collection
        Else
            If ItemExists(colResultPrice(intRaceNum), strKey) Then
                colResultPrice(intRaceNum).Remove(strKey)
            End If
        End If

        colResultPrice(intRaceNum).Add(m_objResultWPS, strKey)

        If (UCase(m_objResultWPS.PoolCode) = "WIN") Then
            RaiseEvent NewResultsWIN(m_objResultWPS, intRaceNum)
        ElseIf (UCase(m_objResultWPS.PoolCode) = "PLC") Then
            RaiseEvent NewResultsPLC(m_objResultWPS, intRaceNum)
        ElseIf (UCase(m_objResultWPS.PoolCode) = "SHW") Then
            RaiseEvent NewResultsSHW(m_objResultWPS, intRaceNum)
        End If

        Exit Sub

ErrHndlr:
        'Implement error logging
    End Sub

    Private Sub IRaceFxNotify_DataObjectRunnerTotals(ByVal m_objRunnerTotals As RSIData.clsRunnerTotals, ByVal intRaceNum As Short) Implements RSIService.IRaceFxNotify.DataObjectRunnerTotals
        Dim strKey As String

        On Error GoTo ErrHndlr

        strKey = intRaceNum & "_" & m_objRunnerTotals.PoolCode

        If colRunnerTotals(intRaceNum) Is Nothing Then
            colRunnerTotals(intRaceNum) = New Collection
        Else
            If ItemExists(colRunnerTotals(intRaceNum), strKey) Then
                colRunnerTotals(intRaceNum).Remove(strKey)
            End If
        End If

        colRunnerTotals(intRaceNum).Add(m_objRunnerTotals, strKey)

        If (m_objRunnerTotals.PoolCode = "WIN") Then
            RaiseEvent NewWINPool(m_objRunnerTotals, intRaceNum)
        ElseIf (m_objRunnerTotals.PoolCode = "PLC") Then
            RaiseEvent NewPlacePool(m_objRunnerTotals, intRaceNum)
        ElseIf (m_objRunnerTotals.PoolCode = "SHW") Then
            RaiseEvent NewShowPool(m_objRunnerTotals, intRaceNum)
        End If

        Exit Sub

ErrHndlr:
        'Implement error logging
    End Sub

    Private Sub IRaceFxNotify_DataObjectTimeInfo(ByRef m_strTimeInfo As String) Implements RSIService.IRaceFxNotify.DataObjectTimeInfo
        On Error GoTo ErrHndlr

        m_strCurrentTime = m_strTimeInfo

        'RaiseEvent TimeOfDayChange

        Exit Sub

ErrHndlr:
        'Implement error logging
    End Sub

    Private Sub IRaceFxNotify_DataObjectTTData(ByVal m_objTeleTimer As RSIData.clsTeleTimer, ByVal intRaceNum As Short) Implements RSIService.IRaceFxNotify.DataObjectTTData
        Dim strKey As String
        Dim colTemp As Collection

        On Error GoTo ErrHndlr

        'Code needed for Amtote
        If intRaceNum = g_intMaxNumbOfRaces Then
            intRaceNum = m_intCurrentRace
            m_objTeleTimer.Race = intRaceNum
        End If

        strKey = "TT" & intRaceNum

        colTemp = New Collection

        colTemp.Add(m_objTeleTimer, strKey)

        colTeletimer(intRaceNum) = colTemp

        'UPGRADE_NOTE: Object colTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        colTemp = Nothing

        RaiseEvent NewTT(m_objTeleTimer, intRaceNum)

        Exit Sub

ErrHndlr:
        'Implement error logging
    End Sub

    Private Sub IRaceFxNotify_DataObjectWillPays(ByVal m_objWillPay As RSIData.clsWillPays, ByVal intRaceNum As Short) Implements RSIService.IRaceFxNotify.DataObjectWillPays
        'Dim strKey As String
        '
        '  strKey = intRaceNum & "_" & m_objWillPay.BetAmount & m_objWillPay.PoolCode
        '
        '  If colWillPays(intRaceNum) Is Nothing Then
        '    Set colWillPays(intRaceNum) = New Collection
        '  Else
        '    If ItemExists(colWillPays(intRaceNum), strKey) Then
        '      colWillPays(intRaceNum).Remove strKey
        '    End If
        '  End If
        '
        '  colWillPays(intRaceNum).Add m_objWillPay, strKey
        '
        '  strKey = intRaceNum & "_" & m_objWillPay.PoolCode
        '
        '  If colWillPaysAutototeV6(intRaceNum) Is Nothing Then
        '    Set colWillPaysAutototeV6(intRaceNum) = New Collection
        '  Else
        '    If ItemExists(colWillPaysAutototeV6(intRaceNum), strKey) Then
        '      colWillPaysAutototeV6(intRaceNum).Remove strKey
        '    End If
        '  End If
        '
        '  colWillPaysAutototeV6(intRaceNum).Add m_objWillPay, strKey
    End Sub

    Private Sub LoadPortSettings_Timer(ByRef strSettings As String)
        Dim Pntr, Pntr1 As Short

        On Error GoTo ErrHndlr

        'setting the properties to be viewed from RaceFX
        '"9600,n,8,1"

        Pntr1 = InStr(strSettings, ",")
        m_strBaud_Timer = Left(strSettings, Pntr1 - 1)

        Pntr = InStr(Pntr1 + 1, strSettings, ",")
        m_strParity_Timer = Mid(strSettings, Pntr1 + 1, Pntr - Pntr1 - 1)

        Pntr1 = InStr(Pntr + 1, strSettings, ",")
        m_strDataBits_Timer = Mid(strSettings, Pntr + 1, Pntr1 - Pntr - 1)

        m_strStopBits_Timer = Right(strSettings, Len(strSettings) - Pntr1)

        Exit Sub

ErrHndlr:
        'Implement error logging
    End Sub

    Private Sub LoadPortSettings_Tote(ByRef strSettings As String)
        Dim Pntr, Pntr1 As Short

        On Error GoTo ErrHndlr

        'setting the properties to be viewed from RaceFX

        Pntr1 = InStr(strSettings, ",")
        m_strBaud_Tote = Left(strSettings, Pntr1 - 1)

        Pntr = InStr(Pntr1 + 1, strSettings, ",")
        m_strParity_Tote = Mid(strSettings, Pntr1 + 1, Pntr - Pntr1 - 1)

        Pntr1 = InStr(Pntr + 1, strSettings, ",")
        m_strDataBits_Tote = Mid(strSettings, Pntr + 1, Pntr1 - Pntr - 1)

        m_strStopBits_Tote = Right(strSettings, Len(strSettings) - Pntr1)

        Exit Sub

ErrHndlr:
        'Implement error logging
    End Sub

    Private Sub LoadComPortSettings(ByRef comm As MSCommLib.MSComm, ByVal comPortConfigSection As String, ByVal comSettingConfigSection As String)

        Dim portNumber As String = ConfigurationManager.AppSettings(comPortConfigSection)
        Dim portSetting As String = ConfigurationManager.AppSettings(comSettingConfigSection)

        If (String.IsNullOrWhiteSpace(portNumber) Or String.IsNullOrWhiteSpace(portSetting)) Then
            portNumber = "1"
            portSetting = "9600,n,8,1"
        Else
            ValidatePortSettings(portNumber, portSetting)
        End If

        With comm
            .Settings = portSetting
            .InBufferSize = 32500
            .RThreshold = 1
            .InputLen = 1
            .CommPort = portNumber
        End With

    End Sub

    Private Sub ValidatePortSettings(ByVal portNumber As String, ByVal settings As String)
        Dim Pntr, Pntr1 As Short

        Pntr1 = InStr(settings, ",")
        Dim strBaud As String = Left(settings, Pntr1 - 1)

        Pntr = InStr(Pntr1 + 1, settings, ",")
        Dim strParity As String = Mid(settings, Pntr1 + 1, Pntr - Pntr1 - 1)

        Pntr1 = InStr(Pntr + 1, settings, ",")
        Dim strDataBits As String = Mid(settings, Pntr + 1, Pntr1 - Pntr - 1)

        Dim strStopBits As String = Right(settings, Len(settings) - Pntr1)

        Dim result As Integer
        If Not Int32.TryParse(portNumber, result) Then
            Throw New InvalidOperationException("Invalid port number")
        End If

        If Not Int32.TryParse(strBaud, result) Then
            Throw New InvalidOperationException("Invalid baud rate")
        End If

        If strParity.ToLower() <> "n" And strParity.ToLower() <> "o" And strParity.ToLower() <> "E" Then
            Throw New InvalidOperationException("Invalid Parity")
        End If

        If strStopBits.ToLower() <> "n" And strStopBits <> "1" And strStopBits <> "2" Then
            Throw New InvalidOperationException("Invalid stop bits")
        End If

    End Sub


    Private Sub TimeDelay(ByRef lngDelay As Integer)
        Sleep(lngDelay)
    End Sub

    'Private Sub SaveSettings(strService As String, strSettings As String)
    '
    '  MySet.Open "select * from RaceFxConfig where ObjectName = '" _
    ''  & strService & "' and ParameterName = 'Settings'", MyDb, adOpenDynamic, adLockOptimistic
    '  MySet.Fields("ParameterValue") = strSettings
    '  MySet.Update
    '  MySet.Close
    '
    'End Sub

    'Private Sub LoadDbSettings()
    '  'load custom settings
    '  MySet.Open "select * from tbl_RFXConfig where ObjectName = 'CommServer'", MyDb
    '  Do Until MySet.EOF
    '    DoEvents
    '    Select Case Trim(LCase(MySet!ParameterName))
    '      Case "messagedeliveryinterval"
    '        Timer1.Timer1.Interval = Val(MySet!ParameterValue)
    '      Case "numofworkers"
    '        'JECX2...Number of workers...if not present, defaults to 10
    '        m_intNumWorkers = Val(MySet!ParameterValue)
    '      Case "toteprotocol"
    '        'JECX2...tote company to use...defaults to 0..Autotoe
    '        m_intProtocol = Val(MySet!ParameterValue)
    '      Case "trackcode"
    '        m_strTrackCode = MySet!ParameterValue
    '      'Case "connectionstring"
    '      '  m_strConnectionString = MySet!ParameterValue
    '      Case "oddsstatistics"
    '        m_blnOddsStatistics = MySet!ParameterValue
    '    End Select
    '    MySet.MoveNext
    '  Loop
    '  MySet.Close
    '  '
    '  MySet.Open "select * from tbl_RFXConfig where ObjectName = 'CommTote'", MyDb
    '  Do Until MySet.EOF
    '    DoEvents
    '    With CommTote.CommTote
    '      Select Case Trim(LCase(MySet!ParameterName))
    '        Case "settings"
    '          .Settings = MySet!ParameterValue
    '        Case "inbuffersize"
    '          .InBufferSize = MySet!ParameterValue
    '        Case "rthreshold"
    '          .RThreshold = MySet!ParameterValue
    '        Case "inputlen"
    '          .InputLen = MySet!ParameterValue
    '        Case "commport"
    '          .CommPort = MySet!ParameterValue
    '        Case "MessageDeliveryInterval"
    '          Timer1.Timer1.Interval = MySet!ParameterValue
    '      End Select
    '    End With
    '    MySet.MoveNext
    '  Loop
    '  MySet.Close
    '  '
    '  MySet.Open "select * from tbl_RFXConfig where ObjectName = 'CommTimer'", MyDb
    '  Do Until MySet.EOF
    '    DoEvents
    '    With CommTimer.CommTimer
    '      Select Case Trim(LCase(MySet!ParameterName))
    '        Case "mode"
    '          m_intTimingMode = Val(MySet!ParameterValue) '0-->Amtote(3 or 5 in system)
    '        Case "settings"
    '          .Settings = MySet!ParameterValue
    '        Case "inbuffersize"
    '          .InBufferSize = MySet!ParameterValue
    '        Case "rthreshold"
    '          .RThreshold = MySet!ParameterValue
    '        Case "inputlen"
    '          .InputLen = MySet!ParameterValue
    '        Case "commport"
    '          .CommPort = MySet!ParameterValue
    '        Case "messagetype"
    '          p_objTimerData.MessageType = Val(MySet!ParameterValue)
    '      End Select
    '    End With
    '    MySet.MoveNext
    '  Loop
    '  MySet.Close
    'End Sub

    Private Sub LogEntry(ByRef logstring As String)
        Dim Hndl As Double

        On Error GoTo ErrHndlr

        Hndl = FreeFile()
        FileOpen(Hndl, "c:\ToteBoard\RaceFxCsvr.log", OpenMode.Append)
        PrintLine(Hndl, VB.Timer() & "::" & logstring)
        FileClose(Hndl)

        Exit Sub

ErrHndlr:
        'Implement error logging
    End Sub

    Private Sub Timer1_TimerTriggered(ByRef fTriggered As Boolean) Handles Timer1.TimerTriggered
        On Error GoTo ErrHndlr

        If exitRequested Then
            Timer1.Timer1.Enabled = False
            Exit Sub
        End If

        serviceCommTote()
        If m_intTimingMode = 0 Then
            'ServiceCommTimer_Amtote 'Timming
        ElseIf m_intTimingMode = 6 Then
            'ServiceCommTimer_Canada 'Timming
        ElseIf m_intTimingMode = 7 Then
            'ServiceCommTimer_CanadaNew 'Timming
        Else
            'ServiceCommTimer 'Timming
        End If

        Exit Sub

ErrHndlr:
        'Implement error logging
    End Sub

    Private Sub Timer2_TimerTriggered(ByRef fTriggered As Boolean) Handles Timer2.TimerTriggered
        On Error GoTo ErrHndlr

        Timer1.Timer1.Enabled = False

        If exitRequested Then
            Timer2.Timer2.Enabled = False
            Exit Sub
        End If

        'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Dir("C:\RSI\StopRSIPort.tmp", FileAttribute.Normal) <> "" Then
            Kill("C:\RSI\StopRSIPort.tmp")
            On Error GoTo jump1
            CommTote.CommTote.PortOpen = False
jump1:
            On Error GoTo jump2
            'CommTimer.CommTimer.PortOpen = False 'Timming
jump2:
            End
        End If

        Timer1.Timer1.Enabled = True

        Exit Sub

ErrHndlr:
        'Implement error logging
    End Sub

    'Nov-14, 2001 - EMJ - added procedure to set the interval for the timer
    Public Sub Set_Timer_Interval(ByRef intInterval As Short)
        Timer1.Timer1.Interval = intInterval
    End Sub

    Private Function ItemExists(ByRef col As Collection, ByRef key As String) As Boolean
        'If key doesn't exist, the collection object raises an error 5-"Invalid procedure
        'call or argument".
        Dim dummy As Object
        On Error Resume Next
        'UPGRADE_WARNING: Couldn't resolve default property of object col.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object dummy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        dummy = col.Item(key)
        ItemExists = (Err.Number <> 5)
    End Function

    'Private Sub loadArrayMsgType()
    'Dim intCtr As Integer
    '
    '  For intCtr = LBound(m_arrMsgType) To UBound(m_arrMsgType)
    '    Select Case intCtr
    '      Case 1
    '        m_arrMsgType(intCtr) = "WO"
    '      Case 2
    '        m_arrMsgType(intCtr) = "FN"
    '      Case 3
    '        m_arrMsgType(intCtr) = "WR"
    '      Case 4
    '        m_arrMsgType(intCtr) = "RO"
    '      Case 5
    '        m_arrMsgType(intCtr) = "RS"
    '      Case 6
    '        m_arrMsgType(intCtr) = "RE"
    '      Case 7
    '        m_arrMsgType(intCtr) = "PB"
    '      Case 8
    '        m_arrMsgType(intCtr) = "WP"
    '      Case 9
    '        m_arrMsgType(intCtr) = "TT"
    '      Case 10
    '        m_arrMsgType(intCtr) = "RI"
    '    End Select
    '  Next intCtr
    '
    'End Sub

    Public Sub LoadMessageObject()
        Dim p_objSystem As Scripting.FileSystemObject
        Dim p_objText As Scripting.TextStream
        Dim strTemp As String
        Dim intPos As Short
        Dim strMsgTypes As String
        Dim varMsgTypes As Object
        Dim intCtr As Short

        On Error GoTo ErrHndlr

        p_objSystem = CreateObject("Scripting.FileSystemObject")
        p_objText = p_objSystem.OpenTextFile(m_objIniFile.PathIniFile, Scripting.IOMode.ForReading, Scripting.Tristate.TristateFalse)

        With p_objText
            Do Until .AtEndOfStream
                strTemp = .ReadLine
                intPos = InStr(strTemp, "=")
                If intPos > 0 Then
                    strMsgTypes = strMsgTypes & Left(strTemp, intPos) & "*"
                End If
            Loop
        End With

        p_objText.Close()

        'UPGRADE_NOTE: Object p_objText may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        p_objText = Nothing
        'UPGRADE_NOTE: Object p_objSystem may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        p_objSystem = Nothing

        'UPGRADE_WARNING: Couldn't resolve default property of object varMsgTypes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        varMsgTypes = Split(strMsgTypes, "*")
        For intCtr = LBound(varMsgTypes) To UBound(varMsgTypes) - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object varMsgTypes(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            GetIniMessages(varMsgTypes(intCtr))
        Next intCtr

        Exit Sub

ErrHndlr:
        'Implement error logging
    End Sub

    'This sub reads the messages from the ini file and sends it to the RSIService, so that
    'we could load the message's object.
    Private Sub GetIniMessages(ByVal strMessageType As String)
        Dim p_objSystem As Scripting.FileSystemObject
        Dim p_objText As Scripting.TextStream
        Dim strMessage As String
        Dim intWorker As Short
        Dim intNumTries As Short
        Dim blnFindWorker As Boolean

        On Error GoTo ErrHndlr

        p_objSystem = CreateObject("Scripting.FileSystemObject")
        p_objText = p_objSystem.OpenTextFile(m_objIniFile.PathIniFile, Scripting.IOMode.ForReading, Scripting.Tristate.TristateFalse)

        With p_objText
            Do Until .AtEndOfStream
                strMessage = .ReadLine
                If Left(strMessage, Len(strMessageType)) = strMessageType Then
                    strMessage = Mid(strMessage, Len(strMessageType) + 1)
                    If Left(strMessage, 1) = "^" Then
                        If ((Right(strMessage, 1)) = ("^")) Then
                            strMessage = Mid(strMessage, 1, Len(strMessage) - 1)
                        End If
                        'Find a worker who is not busy
                        'Try four times before giving up.
                        blnFindWorker = False
                        'For intNumTries = 0 To 6
                        Do While Not blnFindWorker
                            For intWorker = 0 To m_intNumWorkers - 1
                                If Not m_objDataWorkers(intWorker).Busy Then
                                    'give the worker the data and keep on working
                                    m_objDataWorkers(intWorker).ProcessSerialData(strMessage)
                                    Debug.Print("worker: " & intWorker & " Msg: " & strMessage)
                                    blnFindWorker = True
                                    TimeDelay(100)
                                    Exit For
                                End If
                            Next intWorker
                        Loop
                        '  If blnFindWorker Then
                        '    Exit For
                        '  End If
                        '  'wait a hundredth of a second before triyng again
                        '  TimeDelay 100
                        'Next intNumTries
                    ElseIf Left(strMessage, 2) = "HL" Then
                        LoadHorseStatistics(strMessage)
                    End If
                    Exit Do
                End If
            Loop
        End With
        p_objText.Close()

        'UPGRADE_NOTE: Object p_objText may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        p_objText = Nothing
        'UPGRADE_NOTE: Object p_objSystem may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        p_objSystem = Nothing

        Exit Sub

ErrHndlr:
        LogEntry("Error processing tote message from iniFile " & CStr(Err.Number) & " " & Err.Description)

    End Sub

    Private Sub LoadHorseStatistics(ByRef strMessage As String)
        'Dim varMsg As Variant
        'Dim intRaceNumber As Integer
        'Dim intRunnerFamily As Integer
        'Dim udtOddsStatistics As typOddsStatistics
        '
        'On Error GoTo ErrHndlr
        '
        '  varMsg = Split(strMessage, "|")
        '  intRaceNumber = varMsg(1)
        '  intRunnerFamily = varMsg(2)
        '  If colOddsStatistics(intRaceNumber) Is Nothing Then
        '    Set colOddsStatistics(intRaceNumber) = New Collection
        '  End If
        '  If ItemExists(colOddsStatistics(intRaceNumber), "R" & intRunnerFamily) Then
        '    udtOddsStatistics = colOddsStatistics(intRaceNumber).Item("R" & intRunnerFamily)
        '    colOddsStatistics(intRaceNumber).Remove "R" & intRunnerFamily
        '  End If
        '  udtOddsStatistics.strHi = varMsg(4)
        '  udtOddsStatistics.strLo = varMsg(6)
        '  colOddsStatistics(intRaceNumber).Add udtOddsStatistics, "R" & intRunnerFamily
        '
        'Exit Sub
        '
        'ErrHndlr:
        '  'Implement error logging
    End Sub

    Public Sub ResetMessageObjects()
        Dim intCounter As Short

        On Error GoTo ErrHndlr

        m_intCurrentRace = 1
        m_strCurrentMTP = " "
        m_strCurrentTrackCondition = " "
        m_strCurrentPostTime = " "
        m_strCurrentStatus = " "
        m_strCurrentRunnersFlashingStatus = " "
        m_strCurrentOfficialStatus = " "
        'm_strCurrentTime = CStr(Time)
        m_dtmCurrentDate = Today

        For intCounter = 1 To g_intMaxNumbOfRaces
            'UPGRADE_NOTE: Object colFinisherData() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colFinisherData(intCounter) = Nothing
            'UPGRADE_NOTE: Object colMTPInfo() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colMTPInfo(intCounter) = Nothing
            'UPGRADE_NOTE: Object colOdds() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colOdds(intCounter) = Nothing
            'UPGRADE_NOTE: Object colOrderKey() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colOrderKey(intCounter) = Nothing
            'UPGRADE_NOTE: Object colProbables() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colProbables(intCounter) = Nothing
            'UPGRADE_NOTE: Object colResultExotic() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colResultExotic(intCounter) = Nothing
            'UPGRADE_NOTE: Object colResultFullExotic() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colResultFullExotic(intCounter) = Nothing
            'UPGRADE_NOTE: Object colResultExoticAutototeV6() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colResultExoticAutototeV6(intCounter) = Nothing
            'UPGRADE_NOTE: Object colResultPrice() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colResultPrice(intCounter) = Nothing
            'UPGRADE_NOTE: Object colRunningOrder() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colRunningOrder(intCounter) = Nothing
            'UPGRADE_NOTE: Object colPoolTotals() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colPoolTotals(intCounter) = Nothing
            'UPGRADE_NOTE: Object colRunnerTotals() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colRunnerTotals(intCounter) = Nothing
            'UPGRADE_NOTE: Object colWillPays() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colWillPays(intCounter) = Nothing
            'UPGRADE_NOTE: Object colWillPaysAutototeV6() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colWillPaysAutototeV6(intCounter) = Nothing
            'UPGRADE_NOTE: Object colTeletimer() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colTeletimer(intCounter) = Nothing
        Next intCounter

        Exit Sub

ErrHndlr:
        'Implement error logging
    End Sub

    Private Sub ServiceCommTimer_Amtote()
        'Timming
        'Dim STXpsn As Integer
        'Dim ETXpsn As Integer
        'Dim BuffLen As Integer
        'Dim intPost As Integer
        '
        '  On Error GoTo TimerError
        '
        '  CommTimer.TimerBuffer = CommTimer.TimerBuffer & CommTimer.CommTimer.Input
        '
        '  'test for start clock
        '  intPost = InStr(CommTimer.TimerBuffer, Chr(26) & Chr(48) & Chr(28))
        '  If intPost > 0 Then
        '    RaiseEvent StartTimerClock(True)
        '    CommTimer.TimerBuffer = Mid(CommTimer.TimerBuffer, intPost + 3)
        '  End If
        '
        '  'test for possible message
        '  STXpsn = InStr(CommTimer.TimerBuffer, Chr$(1))
        '  If STXpsn > 0 Then
        '    BuffLen = Len(CommTimer.TimerBuffer)
        '    If STXpsn > 1 Then
        '      'stx present - strip preceeding data
        '      CommTimer.TimerBuffer = Right$(CommTimer.TimerBuffer, BuffLen - STXpsn + 1)
        '      BuffLen = Len(CommTimer.TimerBuffer)
        '    End If
        '    If BuffLen > 1 Then
        '      ETXpsn = Asc(Mid(CommTimer.TimerBuffer, 2, 1))
        '      If BuffLen = ETXpsn Then
        '        'message available
        '        p_objTimerData.ProcTimerData Left$(CommTimer.TimerBuffer, ETXpsn)
        '        RaiseEvent TimerMsg(Left$(CommTimer.TimerBuffer, ETXpsn))
        '        CommTimer.TimerBuffer = Right$(CommTimer.TimerBuffer, BuffLen - ETXpsn)
        '      End If
        '    End If
        '  End If
        '
        '  On Error GoTo 0
        '
        '  Exit Sub
        '
        'TimerError:
        '  On Error GoTo 0
    End Sub

    Private Sub ServiceCommTimer_CanadaNew()
        'Timming
        'Dim ETXpsn As Integer
        'Dim strCurrentMsg As String
        'Dim STXpsn As Integer
        'Dim BuffLen As Integer
        'Static strTemp As String
        'Dim intCtr As Integer
        '
        'On Error GoTo TimerError
        '
        '  CommTimer.TimerBuffer = CommTimer.TimerBuffer & CommTimer.CommTimer.Input
        '  'test for possible message
        '
        '  STXpsn = InStr(CommTimer.TimerBuffer, Chr$(2))
        '  If STXpsn > 0 Then
        '    BuffLen = Len(CommTimer.TimerBuffer)
        '    If STXpsn > 1 Then
        '      'stx present - strip preceeding data
        '      CommTimer.TimerBuffer = Right$(CommTimer.TimerBuffer, BuffLen - STXpsn + 1)
        '    End If
        '    BuffLen = Len(CommTimer.TimerBuffer)
        '    ETXpsn = InStr(CommTimer.TimerBuffer, Chr$(13))
        '    If ETXpsn > 0 Then
        '      'message availabl
        '
        '      If strTemp <> Left$(CommTimer.TimerBuffer, ETXpsn) Then
        '        p_objTimerData.ProcTimerData Left$(CommTimer.TimerBuffer, ETXpsn)
        '        RaiseEvent TimerMsg(Left$(CommTimer.TimerBuffer, ETXpsn))
        '        'Debug.Print " >> IN <<"
        '      Else
        '        'Debug.Print "OUT"
        '      End If
        '      strTemp = Left$(CommTimer.TimerBuffer, ETXpsn)
        '      CommTimer.TimerBuffer = Right$(CommTimer.TimerBuffer, BuffLen - ETXpsn)
        '    End If
        '  End If
        '
        '  On Error GoTo 0
        '  Exit Sub
        'TimerError:
        '  'Me.Caption = "RaceFx Comm Server (" & Now() & " " & Error() & ")"
        'On Error GoTo 0
    End Sub

    Private Sub ServiceCommTimer_Canada()
        'Timming
        'Dim ETXpsn As Integer
        'Dim strCurrentMsg As String
        'On Error GoTo TimerError
        '
        '  CommTimer.TimerBuffer = CommTimer.TimerBuffer & CommTimer.CommTimer.Input
        '
        '  'test for possible message
        '  ETXpsn = InStr(CommTimer.TimerBuffer, Chr(13))
        '  If ETXpsn > 0 Then
        '    If Len(Mid(CommTimer.TimerBuffer, 1, ETXpsn)) >= 9 Then
        '      strCurrentMsg = Mid(CommTimer.TimerBuffer, ETXpsn - 8, 9)
        '      If Mid(strCurrentMsg, 8, 1) = ">" Then
        '        Select Case Mid(strCurrentMsg, 7, 1)
        '          Case " ", "-", ">"
        '            'message available
        '            p_objTimerData.ProcTimerData strCurrentMsg
        '            RaiseEvent TimerMsg(strCurrentMsg)
        '        End Select
        '      End If
        '    ElseIf Len(Mid(CommTimer.TimerBuffer, 1, ETXpsn)) >= 7 Then
        '      strCurrentMsg = Mid(CommTimer.TimerBuffer, ETXpsn - 6, 7)
        '      If Mid(strCurrentMsg, 6, 1) = ">" Then
        '        Select Case LCase(Mid(strCurrentMsg, 1, 1))
        '          Case "a", "b", "d", "i"
        '            p_objTimerData.ProcTimerData strCurrentMsg
        '            RaiseEvent TimerMsg(strCurrentMsg)
        '        End Select
        '      End If
        '    End If
        '    CommTimer.TimerBuffer = Right$(CommTimer.TimerBuffer, (Len(CommTimer.TimerBuffer)) - ETXpsn)
        '  End If
        '
        '  On Error GoTo 0
        '  Exit Sub
        'TimerError:
        '  'Me.Caption = "RaceFx Comm Server (" & Now() & " " & Error() & ")"
        'On Error GoTo 0
    End Sub
End Class