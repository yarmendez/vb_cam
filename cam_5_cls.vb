' YRM : Global wrapper for digital microscopes
'
' 

Public Class cam_cls
    ' All real function happens in the form.
    '
    Private myForm As cam_frm = Nothing
    '
    ' Public methods, as specified in SRS Document
    '   Note: changes to the pdf:
    '       1. Defined error codes here
    '       2. Added some additional methods.
    '       3. changed parameters:
    ' 
    '   Initialize(ByRef statusCode As Integer)
    '   Connect(ByVal cameraNumber As Integer, ByVal videoMode As Integer, ByRef statusCode As Integer)
    '   Record and Picture both have additional parameters,
    '   UseTimestamp: signals to add a timestamp to the output file name. 
    '
    ' Define return status codes
    Public Const ErrorCodeNone As Integer = 100             ' no error reported, most likely it worked, but not guaranteed
    Public Const ErrorCodeInvalidParameter As Integer = 101 ' wrong camera or video mode
    Public Const ErrorCodeOperationFail As Integer = 102    ' sdk returned fail
    Public Const ErrorCodeOperationNotPossible As Integer = 103 ' must be viewing live before trying a video or image save
    Public Const ErrorCodeUnsupportCommand As Integer = 104     'the DisplayLiveVideo command returns this and is redundant as the connect command starts live video
    Public Const ErrorCodeNotInitialized As Integer = 105       ' exception accoured in sdk call
    Public Const ErrorCodeException As Integer = 109            ' exception accoured in sdk call

    'Default values
    Public iTimerInterval As Integer = 10000
    Public iDaysToKeepFiles As Integer
    Public bFlagToDelete As Boolean = False
    Public objDateToDelete As Object = CDate(Format(Now, "yyy-MM-dd"))
    Public bIsServiceInitialized As Boolean = False ' 0224 YRM

    Public Shared errorLogFile As String = "ErrorLogFile.txt"
    Public Shared Sub LogErrorMessage(ByVal emsg As String)
        Try
            My.Computer.FileSystem.WriteAllText(errorLogFile, vbCrLf + emsg, True)
        Catch ex As Exception
        End Try
    End Sub

    ' Cam video capture mode
    Public Enum VideoMode
        Mode2592 = 1
        Mode2048 = 2
        Mode1600 = 3
        Mode1280 = 4
        Mode640 = 5
    End Enum

    ' Define extra functions in addition to SRS
    Public Sub SetErrorLogFile(ByVal filepath As String, ByRef statusCode As Integer)
        Try
            My.Computer.FileSystem.WriteAllText(filepath, "Set Net Log file path:" + filepath, True)
            errorLogFile = filepath
            statusCode = ErrorCodeNone
        Catch ex As Exception
            initializeSaveException = ex
            statusCode = ErrorCodeException
        End Try
    End Sub

    Public Function GetErrorLogFile() As String
        Return errorLogFile
    End Function

    Public Sub CloseDisplayForm(ByRef statusCode As Integer)
        ' If there are unknown errors or fail to function, the video display form can 
        '       be closed and reopened. This may clear persistent errors. 
        ' If this is called, the Initialize method will need to be called before any other.
        statusCode = cam_cls.ErrorCodeNone
        If cam_d_frm IsNot Nothing Then
            statusCode = cam_d_frm.FormClose()
            cam_d_frm = Nothing
        End If
    End Sub

    Private initializeSaveException As Exception = Nothing
    Public Function GetLastException(ByRef statusCode As Integer) As Exception
        statusCode = cam_cls.ErrorCodeNone
        If cam_d_frm Is Nothing Then
            If initializeSaveException IsNot Nothing Then
                Return initializeSaveException
            End If
        Else
            Return cam_d_frm.lastException
        End If
        statusCode = cam_cls.ErrorCodeNotInitialized
        Return Nothing
    End Function

    Public Function GetLastErrorMessage(ByRef statusCode As Integer) As String
        If cam_d_frm IsNot Nothing Then
            Return cam_d_frm.lastErrorString
        End If
        statusCode = ErrorCodeNotInitialized
        Return ""
    End Function

    Public Function getLastStatus(ByRef statusCode As Integer) As String
        If cam_d_frm Is Nothing Then
            statusCode = ErrorCodeNotInitialized
            Return ""
        Else
            Return cam_d_frm.DNVgetLastStatusString()
        End If
        Return ""
    End Function

    Public Function getStatusData(ByVal CamNumber As Integer,
        ByRef connected As Boolean,
        ByRef resolutionX As Integer,
        ByRef resolutionY As Integer,
        ByRef framesPerSecond As Single,
        ByRef capturing As Boolean,
        ByRef secondsCapturing As Integer,
        ByRef framesCaptured As Integer,
        ByRef framesDropped As Integer) As Integer
        If cam_d_frm Is Nothing Then
            Return ErrorCodeNotInitialized
        Else
            Return cam_d_frm.DNVgetStatusData(CamNumber,
             connected,
             resolutionX,
             resolutionY,
             framesPerSecond,
             capturing,
             secondsCapturing,
             framesCaptured,
             framesDropped)
        End If
    End Function

    Public Function getStatusDataString(ByVal CamNumber) As String
        Dim connected As Boolean
        Dim resolutionX, resolutionY, framesPerSecond As Integer
        Dim capturing As Boolean
        Dim secondsCapturing, framesCaptured, framesDropped As Integer

        getStatusData(CamNumber, connected,
            resolutionX, resolutionY, framesPerSecond,
            capturing,
            secondsCapturing, framesCaptured, framesDropped)

        Dim result1 As String
        result1 = "Cam: " & CamNumber & ", Conn:" & connected &
            ", Res:" & resolutionX & " x " & resolutionY & ", " &
            framesPerSecond & " FPS, " & ", Capturing:" & capturing & ", " &
            secondsCapturing & " seconds, " & framesCaptured &
            " frames, " & framesDropped & " dropped"
        Return result1
    End Function

    ' Set size for cam form
    Public Sub SetFormSize(ByVal horzX As Integer, ByVal vertY As Integer, ByRef statusCode As Integer)
        ' Autosize the video display to best fit the new form size
        If cam_d_frm Is Nothing Then
            statusCode = ErrorCodeNotInitialized
        Else
            cam_d_frm.SetFormSize(horzX, vertY)
        End If
    End Sub

    ' Define functions per document
    Public Sub Initialize(ByRef statusCode As Integer)

        statusCode = ErrorCodeNone
        initializeSaveException = Nothing
        ' Dim x As Integer = 1
        Try
            If cam_d_frm IsNot Nothing Then
                CloseDisplayForm(statusCode)
            End If
            cam_d_frm = New cam_frm
            cam_d_frm.Show()
            '    x = 2

            ' Enable timer to delete old files
            cam_d_frm.FilesToKeepTimer.Enabled = True

            bIsServiceInitialized = True ' 140224 YRM

        Catch ex1 As Exception
            initializeSaveException = ex1
            '  x = 4
            statusCode = cam_cls.ErrorCodeException
        Finally
            '   x = 3

        End Try
        '  x = 6
    End Sub

    ' Connect to cam#
    Public Sub Connect(ByVal cameraNumber As Integer, ByVal videoMode As Integer, ByRef statusCode As Integer)
        If cam_d_frm Is Nothing Then
            statusCode = ErrorCodeNotInitialized
        Else
            If videoMode < 1 Or videoMode > 5 Then
                statusCode = ErrorCodeInvalidParameter
                Exit Sub
            End If
            If cameraNumber = 1 Then
                statusCode = cam_d_frm.DNVconnectV1(videoMode)
                Exit Sub
            End If
            If cameraNumber = 2 Then
                statusCode = cam_d_frm.DNVconnectV2(videoMode)
                Exit Sub
            End If
            statusCode = ErrorCodeInvalidParameter
        End If
    End Sub

    ' Disconnect cam#
    Public Sub Disconect(ByVal cameraNumber As Integer, ByRef statusCode As Integer)
        If cam_d_frm Is Nothing Then
            statusCode = ErrorCodeNotInitialized
        Else
            If cameraNumber = 1 Then
                statusCode = cam_d_frm.DNVdisConnectV1()
                Exit Sub
            End If
            If cameraNumber = 2 Then
                statusCode = cam_d_frm.DNVdisConnectV2()
                Exit Sub
            End If
            statusCode = ErrorCodeInvalidParameter
        End If
    End Sub

    ' Record pic with cam#
    Public Sub RecordPicture(ByVal cameraNumber As Integer, ByRef statusCode As Integer, ByVal useTimestamp As Boolean, Optional ByVal fileName As String = "")
        If cam_d_frm Is Nothing Then
            statusCode = ErrorCodeNotInitialized
        Else
            If cameraNumber = 1 Then
                statusCode = cam_d_frm.DNVsaveStill1(useTimestamp, fileName)
                Exit Sub
            End If
            If cameraNumber = 2 Then
                statusCode = cam_d_frm.DNVsaveStill2(useTimestamp, fileName)
                Exit Sub
            End If
            statusCode = ErrorCodeInvalidParameter
        End If
    End Sub

    ' Record video with cam#
    Public Sub RecordVideo(ByVal cameraNumber As Integer, ByRef statusCode As Integer, ByVal useTimestamp As Boolean, Optional ByVal fileName As String = "")
        If cam_d_frm Is Nothing Then
            statusCode = ErrorCodeNotInitialized
        Else
            If cameraNumber = 1 Then
                statusCode = cam_d_frm.DNVstartCapture1(useTimestamp, fileName)
                Exit Sub
            End If
            If cameraNumber = 2 Then
                statusCode = cam_d_frm.DNVstartCapture2(useTimestamp, fileName)
                Exit Sub
            End If
            statusCode = ErrorCodeInvalidParameter
        End If
    End Sub

    ' Stop record video with cam#
    Public Sub StopRecordVideo(ByVal cameraNumber As Integer, ByRef statusCode As Integer)
        If cam_d_frm Is Nothing Then
            statusCode = ErrorCodeNotInitialized
        Else
            If cameraNumber = 1 Then
                statusCode = cam_d_frm.DNVstopCapture1()
                Exit Sub
            End If
            If cameraNumber = 2 Then
                statusCode = cam_d_frm.DNVstopCapture2()
                Exit Sub
            End If
            statusCode = ErrorCodeInvalidParameter
        End If
    End Sub

    ' Live video with cam#
    Public Sub DisplayLiveVideo(ByVal cameraNumber As Integer, ByRef statusCode As Integer)
        statusCode = ErrorCodeUnsupportCommand
    End Sub

    Public Sub SetGuiFullAccess(ByRef statusCode As Integer)
        If cam_d_frm Is Nothing Then
            statusCode = ErrorCodeNotInitialized
        Else
            cam_d_frm.SetGuiFullAccess()
        End If
    End Sub

    Public Sub SetGuiNoAccess(ByRef statusCode As Integer)
        If cam_d_frm Is Nothing Then
            statusCode = ErrorCodeNotInitialized
        Else
            cam_d_frm.SetGuiNoAccess()
        End If
    End Sub

    Public Sub SetGuiLowAccess(ByRef statusCode As Integer)
        If cam_d_frm Is Nothing Then
            statusCode = ErrorCodeNotInitialized
        Else
            cam_d_frm.SetGuiLowAccess()
        End If
    End Sub

    Public Sub SetPicturePathFile(ByVal iCameraNumber As Integer, ByVal strPicturePathFile As String)
        cam_d_frm.SetPicturePathFile(iCameraNumber, strPicturePathFile)
    End Sub
    Public Sub SetPictureName(ByVal iCameraNumber As Integer, ByVal strPicturePathFile As String)
        cam_d_frm.SetPicturePathFile(iCameraNumber, strPicturePathFile)
    End Sub
    Public Sub SetVideoRecordDuration_Secs(ByVal sngSeconds As Single)
        cam_d_frm.SetVideoRecordDuration_Secs(sngSeconds)
    End Sub

    Public Sub DeleteOldFiles()
        cam_d_frm.SetDaysToKeepFiles = iDaysToKeepFiles
        cam_d_frm.SetTimeIntervalToDelete = iTimerInterval

        cam_d_frm.DeleteOldFiles()

    End Sub

    ' Utility private  functions
    Private Sub MakeForm()
        cam_d_frm = New cam_frm

    End Sub

    ' Load cam form
    Private Function formLoad()
        Dim s As String
        s = "."

        Try
            cam_d_frm.Show()
            cam_d_frm.TextboxStatus1.Text = " Reshow at:" & Now
            s = cam_d_frm.TextboxStatus1.Text
        Catch ex As Exception
            MsgBox("error showing form" & ex.Message)
        End Try
        Return s
    End Function

    ' Hide cam form
    Private Function formHide()
        cam_d_frm.Hide()
        cam_d_frm.TextboxStatus1.Text = " ..Hide"
        Dim s As String
        s = cam_d_frm.TextboxStatus1.Text
        Return s
    End Function

    ' Close cam form
    Private Function formClose()
        cam_d_frm.TextboxStatus1.Text = " ..close"
        Dim s As String
        s = cam_d_frm.TextboxStatus1.Text
        cam_d_frm.Close()

        Return s
    End Function

End Class

' Class flags
Class FlagsClass
    Public Connect As Boolean = False
    Public ConnectMode As Integer
    Public Discon As Boolean
    Public Record As Boolean
    Public RecordFilename As String
    Public UseTimestamp As Boolean
    Public StopRecord As Boolean
    Public SnapImage As Boolean
    Public SnapImageFilename As String
    Public Sub reset()
        Connect = False
        ConnectMode = 1
        Discon = False
        Record = False
        RecordFilename = ""
        StopRecord = False
        SnapImage = False
        SnapImageFilename = ""
    End Sub
End Class
