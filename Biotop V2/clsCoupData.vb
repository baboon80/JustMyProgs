Public Class clsCoupData


    Public Sub New(Optional ByVal Ldf As Integer = 0, Optional ByVal Coup As String = "", Optional ByVal TP As String = "", Optional ByVal TM As String = "", Optional ByVal TPS As String = "", Optional ByVal TPI As String = "", Optional ByVal TPR As String = "" _
        , Optional ByVal TPR7 As String = "", Optional ByVal TPRS As String = "", Optional ByVal TMS As String = "", Optional ByVal TMI As String = "", Optional ByVal TMR As String = "", Optional ByVal TMR7 As String = "" _
        , Optional ByVal TMRS As String = "", Optional ByVal S As String = "", Optional ByVal R As String = "", Optional ByVal SS As String = "", Optional ByVal SI As String = "", Optional ByVal SR As String = "" _
        , Optional ByVal SR7 As String = "", Optional ByVal SRS As String = "", Optional ByVal RS As String = "", Optional ByVal RI As String = "", Optional ByVal RR As String = "", Optional ByVal RR7 As String = "", Optional ByVal RRS As String = "")
        Me.dLdf = Ldf
        Me.dCoup = Coup
    End Sub
    Public Property Ldf() As String
        Get
            Return dLdf
        End Get
        Set(ByVal Value As String)
            dLdf = Value
        End Set
    End Property
    Public Property Coup() As String
        Get
            Return dCoup
        End Get
        Set(ByVal Value As String)
            dCoup = Value
        End Set
    End Property

    'Transformator
    Public Property TP() As String
        Get
            Return dTP
        End Get
        Set(ByVal Value As String)
            dTP = Value
        End Set
    End Property
    Public Property TM() As String
        Get
            Return dTM
        End Get
        Set(ByVal Value As String)
            dTM = Value
        End Set
    End Property

    'Plus
    Public Property TPS() As String
        Get
            Return dTPS
        End Get
        Set(ByVal Value As String)
            dTPS = Value
        End Set
    End Property
    Public Property TPI() As String
        Get
            Return dTPI
        End Get
        Set(ByVal Value As String)
            dTPI = Value
        End Set
    End Property
    Public Property TPR() As String
        Get
            Return dTPR
        End Get
        Set(ByVal Value As String)
            dTPR = Value
        End Set
    End Property
    Public Property TPR7() As String
        Get
            Return dTPR7
        End Get
        Set(ByVal Value As String)
            dTPR7 = Value
        End Set
    End Property
    Public Property TPRS() As String
        Get
            Return dTPRS
        End Get
        Set(ByVal Value As String)
            dTPRS = Value
        End Set
    End Property

    'Minus
    Public Property TmS() As String
        Get
            Return dTMS
        End Get
        Set(ByVal Value As String)
            dTMS = Value
        End Set
    End Property
    Public Property TmI() As String
        Get
            Return dTMI
        End Get
        Set(ByVal Value As String)
            dTMI = Value
        End Set
    End Property
    Public Property TmR() As String
        Get
            Return dTMR
        End Get
        Set(ByVal Value As String)
            dTMR = Value
        End Set
    End Property
    Public Property TmR7() As String
        Get
            Return dTMR7
        End Get
        Set(ByVal Value As String)
            dTMR7 = Value
        End Set
    End Property
    Public Property TmRS() As String
        Get
            Return dTMRS
        End Get
        Set(ByVal Value As String)
            dTMRS = Value
        End Set
    End Property

    'Schwarz Rot
    Public Property S() As String
        Get
            Return dS
        End Get
        Set(ByVal Value As String)
            dS = Value
        End Set
    End Property
    Public Property R() As String
        Get
            Return dR
        End Get
        Set(ByVal Value As String)
            dR = Value
        End Set
    End Property

    'Schwarz
    Public Property SS() As String
        Get
            Return dSS
        End Get
        Set(ByVal Value As String)
            dSS = Value
        End Set
    End Property
    Public Property SI() As String
        Get
            Return dSI
        End Get
        Set(ByVal Value As String)
            dSI = Value
        End Set
    End Property
    Public Property SR() As String
        Get
            Return dSR
        End Get
        Set(ByVal Value As String)
            dSR = Value
        End Set
    End Property
    Public Property SR7() As String
        Get
            Return dSR7
        End Get
        Set(ByVal Value As String)
            dSR7 = Value
        End Set
    End Property
    Public Property SRS() As String
        Get
            Return dSRS
        End Get
        Set(ByVal Value As String)
            dSRS = Value
        End Set
    End Property

    'Rot
    Public Property RS() As String
        Get
            Return dRS
        End Get
        Set(ByVal Value As String)
            dRS = Value
        End Set
    End Property
    Public Property RI() As String
        Get
            Return dRI
        End Get
        Set(ByVal Value As String)
            dRI = Value
        End Set
    End Property
    Public Property RR() As String
        Get
            Return dRR
        End Get
        Set(ByVal Value As String)
            dRR = Value
        End Set
    End Property
    Public Property RR7() As String
        Get
            Return dRR7
        End Get
        Set(ByVal Value As String)
            dRR7 = Value
        End Set
    End Property
    Public Property RRS() As String
        Get
            Return dRRS
        End Get
        Set(ByVal Value As String)
            dRRS = Value
        End Set
    End Property

    Private dLdf As Integer
    Private dCoup As String
    Private dTP As String
    Private dTM As String
    Private dTPS As String
    Private dTPI As String
    Private dTPR As String
    Private dTPR7 As String
    Private dTPRS As String
    Private dTMS As String
    Private dTMI As String
    Private dTMR As String
    Private dTMR7 As String
    Private dTMRS As String
    Private dS As String
    Private dR As String
    Private dSS As String
    Private dSI As String
    Private dSR As String
    Private dSR7 As String
    Private dSRS As String
    Private dRS As String
    Private dRI As String
    Private dRR As String
    Private dRR7 As String
    Private dRRS As String

End Class
