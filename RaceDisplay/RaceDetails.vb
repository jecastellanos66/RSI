Public Class RaceDetails

#Region " Declarations "

#End Region

#Region " Constructor "

    Public Sub New()
        Me.myDs = New dsRaceDetails
    End Sub

#End Region

#Region " Properties "

    Public ReadOnly Property Dataset() As dsRaceDetails
        Get
            Return Me.myDs
        End Get
    End Property

#End Region

#Region " Public Methods "

#End Region

End Class
