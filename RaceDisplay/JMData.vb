Public Class JMData
    Private _trackMeedID As String
    Private _raceNumber As String
    Private _official As Boolean
    Private _objection As Boolean
    Private _inquiry As Boolean
    Private _photo As Boolean
    Private _deadheat As Boolean
    Private _objectionRunners() As String = New String(4) {}
    Private _inqueryRunners() As String = New String(4) {}
    Private _deadheatRunners() As String = New String(4) {}
    Private _photoRunners() As String = New String(4) {}

    Public Property TrackMeetID() As String
        Get
            Return _trackMeedID
        End Get
        Set(ByVal value As String)
            _trackMeedID = value
        End Set
    End Property

    Public Property RaceNumber() As String
        Get
            Return _raceNumber
        End Get
        Set(ByVal value As String)
            _raceNumber = value
        End Set
    End Property

    Public Property Official() As Boolean
        Get
            Return _official
        End Get
        Set(ByVal value As Boolean)
            _official = value
        End Set
    End Property

    Public Property Objection() As Boolean
        Get
            Return _objection
        End Get
        Set(ByVal value As Boolean)
            _objection = value
        End Set
    End Property

    Public Property Inquiry() As Boolean
        Get
            Return _inquiry
        End Get
        Set(ByVal value As Boolean)
            _inquiry = value
        End Set
    End Property

    Public Property Photo() As Boolean
        Get
            Return _photo
        End Get
        Set(ByVal value As Boolean)
            _photo = value
        End Set
    End Property

    Public Property Deaheat() As Boolean
        Get
            Return _deadheat
        End Get
        Set(ByVal value As Boolean)
            _deadheat = value
        End Set
    End Property

    Public Property ObjectionRunners() As String()
        Get
            Return _objectionRunners
        End Get
        Set(ByVal value() As String)
            _objectionRunners = value
        End Set
    End Property

    Public Property InquiryRunners() As String()
        Get
            Return _inqueryRunners
        End Get
        Set(ByVal value() As String)
            _inqueryRunners = value
        End Set
    End Property

    Public Property DeaheatRunners() As String()
        Get
            Return _deadheatRunners
        End Get
        Set(ByVal value() As String)
            _deadheatRunners = value
        End Set
    End Property

    Public Sub SetObjectionRunners(ByVal runners As String)
        _objectionRunners = ExtractRunners(runners)
    End Sub

    Public Sub SetInquiryRunners(ByVal runners As String)
        _inqueryRunners = ExtractRunners(runners)
    End Sub

    Public Sub SetDeadheatRunners(ByVal runners As String)
        _deadheatRunners = ExtractRunners(runners)
    End Sub

    Public Sub SetPhotoRunners(ByVal runners As String)
        _photoRunners = ExtractRunners(runners)
    End Sub

    Public ReadOnly Property FlashingRunnersPresent() As Boolean
        Get
            Return _objectionRunners.Any(Function(dh) Not String.IsNullOrWhiteSpace(dh)) _
                Or _inqueryRunners.Any(Function(dh) Not String.IsNullOrWhiteSpace(dh)) _
                Or _deadheatRunners.Any(Function(dh) Not String.IsNullOrWhiteSpace(dh)) _
                Or _photoRunners.Any(Function(dh) Not String.IsNullOrWhiteSpace(dh))
        End Get
    End Property

    'Get list of the firs status runners it finds in the order in code
    Public Function GetStatusRuners() As String()
        Dim statusRunners() As String = Nothing

        If _objectionRunners.Any(Function(dh) Not String.IsNullOrWhiteSpace(dh)) Then
            statusRunners = _objectionRunners
        ElseIf _inqueryRunners.Any(Function(dh) Not String.IsNullOrWhiteSpace(dh)) Then
            statusRunners = _inqueryRunners
        ElseIf _deadheatRunners.Any(Function(dh) Not String.IsNullOrWhiteSpace(dh)) Then
            statusRunners = _deadheatRunners
        ElseIf _photoRunners.Any(Function(dh) Not String.IsNullOrWhiteSpace(dh)) Then
            statusRunners = _photoRunners
        End If

        Return statusRunners
    End Function


    Private Function ExtractRunners(ByVal runners As String) As String()
        Dim runnerData() As String = New String(4) {}

        If runners.Length < 15 Then
            ExtractRunners = runnerData
            Exit Function
        End If

        runners = runners.Substring(0, 15)

        runnerData(0) = runners.Substring(0, 3).Trim()
        runnerData(1) = runners.Substring(3, 3).Trim()
        runnerData(2) = runners.Substring(6, 3).Trim()
        runnerData(3) = runners.Substring(9, 3).Trim()
        runnerData(4) = runners.Substring(12, 3).Trim()

        ExtractRunners = runnerData
    End Function

End Class