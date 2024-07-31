
Public Class PigBaseLocal
    Inherits PigToolsLiteLib.PigBase

    Public Sub New(Version As String)
        MyBase.New(Version, System.Reflection.Assembly.GetExecutingAssembly().GetName.Version.ToString, System.Reflection.Assembly.GetExecutingAssembly().GetName.Name)
    End Sub

    Friend Structure StruStepLog
        Public SubName As String
        Public StepName As String
        Public LogInf As String
        Public IsHardDebug As Boolean
        Public Ret As String
        Public Sub AddStepNameInf(AddInf As String)
            Me.StepName &= "(" & AddInf & ")"
        End Sub
        Public ReadOnly Property StepLogInf As String
            Get
                With Me
                    StepLogInf = "[" & Me.SubName & "]"
                    StepLogInf &= "[" & Me.StepName & "]"
                    If Me.Ret <> "" Then StepLogInf &= "[" & Me.Ret & "]"
                    If Me.TrackID <> "" Then StepLogInf &= "[TrackID:" & Me.TrackID & "]"
                    If Me.IsLogUseTime = True Then StepLogInf &= "[" & Me.AllDiffSeconds.ToString & "]"
                End With
            End Get
        End Property

        Private mstrTrackID As String
        Public Property TrackID As String
            Get
                If mstrTrackID Is Nothing Then mstrTrackID = ""
                Return mstrTrackID
            End Get
            Friend Set(value As String)
                mstrTrackID = value
            End Set
        End Property

        Private mbolIsLogUseTime As Boolean
        Public Property IsLogUseTime As Boolean
            Get
                Return mbolIsLogUseTime
            End Get
            Friend Set(value As Boolean)
                mbolIsLogUseTime = value
            End Set
        End Property

        Public BeginTime As DateTime
        Public EndTime As DateTime
        Public CurrBeginTime As DateTime
        Public TimeDiff As TimeSpan
        Public Sub GoBegin()
            BeginTime = Now
            CurrBeginTime = BeginTime
            TimeDiff = Nothing
            TimeDiff = New TimeSpan
        End Sub
        Public Function AllDiffSeconds() As Decimal
            AllDiffSeconds = TimeDiff.TotalMilliseconds / 1000
        End Function
        Public Sub ToEnd()
            Me.EndTime = Now
            TimeDiff = Now - BeginTime
        End Sub

    End Structure

    Friend Shadows Function PrintDebugLog(SubName As String, LogInf As String) As String
        Return MyBase.PrintDebugLog(SubName, LogInf)
    End Function

    Friend Shadows Function PrintDebugLog(SubName As String, StepName As String, LogInf As String) As String
        Return MyBase.PrintDebugLog(SubName, StepName, LogInf)
    End Function

    Friend Shadows Function PrintDebugLog(SubName As String, StepName As String, LogInf As String, IsHardDebug As Boolean) As String
        Return MyBase.PrintDebugLog(SubName, StepName, LogInf, IsHardDebug)
    End Function

    Friend Shadows Function PrintDebugLog(SubName As String, LogInf As String, IsHardDebug As Boolean) As String
        Return MyBase.PrintDebugLog(SubName, LogInf, IsHardDebug)
    End Function

    Friend Shadows Sub ClearErr()
        MyBase.ClearErr()
    End Sub

    Friend Shadows Sub SetDebug(DebugFilePath As String)
        MyBase.SetDebug(DebugFilePath)
    End Sub

    Friend Shadows Sub SetDebug(DebugFilePath As String, IsHardDebug As Boolean)
        MyBase.SetDebug(DebugFilePath, IsHardDebug)
    End Sub

    Friend Shadows Sub SetSubErrInf(SubName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False)
        MyBase.SetSubErrInf(SubName, exIn, IsStackTrace)
    End Sub

    Friend Shadows Sub SetSubErrInf(SubName As String, StepName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False)
        MyBase.SetSubErrInf(SubName, StepName, exIn, IsStackTrace)
    End Sub


    Friend Shadows Function GetSubErrInf(SubName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False) As String
        Return MyBase.GetSubErrInf(SubName, exIn, IsStackTrace)
    End Function

    Friend Shadows Function GetSubErrInf(SubName As String, StepName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False) As String
        Return MyBase.GetSubErrInf(SubName, StepName, exIn, IsStackTrace)
    End Function

    Friend Shadows ReadOnly Property MyID() As String
        Get
            Return MyBase.MyID
        End Get
    End Property

    Friend Shadows ReadOnly Property AppVersion() As String
        Get
            Return MyBase.AppVersion
        End Get
    End Property

    Friend Shadows ReadOnly Property AppTitle() As String
        Get
            Return MyBase.AppTitle
        End Get
    End Property

    Friend Shadows ReadOnly Property AppPath() As String
        Get
            Return MyBase.AppPath
        End Get
    End Property

    Public Shadows ReadOnly Property OsCrLf() As String
        Get
            Return MyBase.OsCrLf
        End Get
    End Property

    Public Shadows ReadOnly Property OsPathSep() As String
        Get
            Return MyBase.OsPathSep
        End Get
    End Property

    Public Shadows ReadOnly Property IsWindows() As Boolean
        Get
            Return MyBase.IsWindows
        End Get
    End Property

    Friend Shadows ReadOnly Property IsDebug() As Boolean
        Get
            Return MyBase.IsDebug
        End Get
    End Property

    Friend Shadows ReadOnly Property IsHardDebug() As Boolean
        Get
            Return MyBase.IsHardDebug
        End Get
    End Property

    Friend Shadows Sub OpenDebug()
        MyBase.OpenDebug()
    End Sub

    Friend Shadows Sub OpenDebug(LogFileDir As String)
        MyBase.OpenDebug(LogFileDir)
    End Sub

    Friend Shadows Sub OpenDebug(LogFileDir As String, IsIncDate As Boolean)
        MyBase.OpenDebug(LogFileDir, IsIncDate)
    End Sub

    Friend Shadows Sub OpenDebug(IsIncDate As Boolean)
        MyBase.OpenDebug(IsIncDate)
    End Sub

End Class

