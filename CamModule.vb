' YRM : Webservice contract and function exec
'
' 
Imports System.ServiceModel
Imports System.ServiceModel.Description

Module Module1WsVProControl

    Private webServiceStatus As String

    ' YRM : Service contract set up 
    <ServiceContract()>
    Public Interface IVproControlService
        <OperationContract()> Function Echo(ByVal message As String) As String
        <OperationContract()> Sub Initialize(ByVal camNum As Integer, ByRef statusCode As Integer)
        <OperationContract()> Sub Connect(ByVal camNum As Integer, ByRef statusCode As Integer)
        <OperationContract()> Sub DisConnect(ByVal camNum As Integer, ByRef statusCode As Integer)
        <OperationContract()> Sub RecordPicture(ByVal camNum As Integer, ByRef statusCode As Integer, ByVal filename As String, ByVal useTimestamp As Boolean)
        <OperationContract()> Sub SetGuiNoAccess(ByRef statusCode As Integer)
        <OperationContract()> Sub SetGuiLowAccess(ByRef statusCode As Integer)
        <OperationContract()> Sub SetGuiFullAccess(ByRef statusCode As Integer)
        <OperationContract()> Sub LoadedProject(ByVal camNum As Integer, ByVal filename As String, ByRef bLoaded As Boolean)
        <OperationContract()> Sub LoadProject(ByVal camNum As Integer, ByRef statusCode As Integer, ByVal filename As String, ByRef bIsLoaded As Boolean, ByRef strLoadedTool As String)
        <OperationContract()> Sub ExecuteInspect(ByVal camNum As Integer, ByRef statusCode As Integer)
        <OperationContract()> Sub GetInspectResults(ByVal camNum As Integer, ByRef statusCode As Integer, ByRef Results As String)
        <OperationContract()> Sub GetServiceStatus(ByRef isServiceUp As Boolean)
        <OperationContract()> Sub DeleteOldFiles(ByRef statusCode As Integer, ByVal iDate As Integer, ByVal strFilePath As String, ByVal iTimerInterval As Integer, Optional objComparisonDate As Object = Nothing)
        <OperationContract()> Sub IsServiceInitialized(ByRef bIsServiceInitialized As Boolean)

    End Interface

    ' YRM : Implementation of control services
    Public Class VproControlService
        Implements IVproControlService

        ' Comm function
        Public Function Echo(ByVal message As String) As String Implements IVproControlService.Echo
            statusAdd("Send:" + message)
            Return message
        End Function
        ' initialize cam# global function
        Public Sub Initialize(ByVal camNum As Integer, ByRef statusCode As Integer) Implements IVproControlService.Initialize
            FormVProControl.Initialize(camNum, statusCode)
            statusAdd("Initialize:" + camNum.ToString + ", retcode=" + statusCode.ToString)
        End Sub
        ' Connect to cam# with global function
        Public Sub Connect(ByVal camNum As Integer, ByRef statusCode As Integer) Implements IVproControlService.Connect
            FormVProControl.Connect(camNum, statusCode)
            statusAdd("Connect:" + camNum.ToString + ", retcode=" + statusCode.ToString)
        End Sub
        ' Connect to cam# with global function
        Public Sub DisConnect(ByVal camNum As Integer, ByRef statusCode As Integer) Implements IVproControlService.DisConnect
            FormVProControl.DisConnect(camNum, statusCode)
            statusAdd("DisConnect:" + camNum.ToString + ", retcode=" + statusCode.ToString)
        End Sub
        ' Record picture global fcn on cam#
        Public Sub RecordPicture(ByVal camNum As Integer, ByRef statusCode As Integer, ByVal filename As String, ByVal useTimestamp As Boolean) Implements IVproControlService.RecordPicture
            statusAdd("RecordPicture Start: " + camNum.ToString)
            FormVProControl.RecordPicture(camNum, statusCode, filename, useTimestamp)
            statusAdd("RecordPicture:" + camNum.ToString + ", filename=" + filename + ", retcode =" + statusCode.ToString)
        End Sub
        ' Set GUI access rules : 0
        Public Sub SetGuiNoAccess(ByRef statusCode As Integer) Implements IVproControlService.SetGuiNoAccess
            FormVProControl.SetGuiNoAccess(statusCode)
            statusAdd("SetGuiNoAccess:, retcode=" + statusCode.ToString)
        End Sub
        ' Set GUI access rules: 1
        Public Sub SetGuiLowAccess(ByRef statusCode As Integer) Implements IVproControlService.SetGuiLowAccess
            FormVProControl.SetGuiLowAccess(statusCode)
            statusAdd("SetGuiLowAccess:, retcode=" + statusCode.ToString)
        End Sub
        ' Set GUI access rules: 2
        Public Sub SetGuiFullAccess(ByRef statusCode As Integer) Implements IVproControlService.SetGuiFullAccess
            FormVProControl.SetGuiFullAccess(statusCode)
            statusAdd("SetGuiFullAccess:, retcode=" + statusCode.ToString)
        End Sub
        ' Check for project load status: 
        Public Sub LoadedProject(ByVal camNum As Integer, ByVal filename As String, ByRef bLoaded As Boolean) Implements IVproControlService.LoadedProject
            statusAdd("LoadedProject Start: " + camNum.ToString)
            FormVProControl.LoadedProject(camNum, filename, bLoaded)
            If bLoaded Then
                statusAdd(filename + " is loaded")
            Else
                statusAdd(filename + " is not loaded ")
            End If
        End Sub
        ' Load project to cam#
        Public Sub LoadProject(ByVal camNum As Integer, ByRef statusCode As Integer, ByVal filename As String, ByRef bIsLoaded As Boolean, ByRef strToolLoaded As String) Implements IVproControlService.LoadProject
            statusAdd("LoadProject Start: " + camNum.ToString)
            FormVProControl.LoadProject(camNum, statusCode, filename, bIsLoaded, strToolLoaded)
            statusAdd("ToolBlockLoad:" + camNum.ToString + ", retcode=" + statusCode.ToString)
        End Sub
        ' Execute inpection tools on cam#
        Public Sub ExecuteInspect(ByVal camNum As Integer, ByRef statusCode As Integer) Implements IVproControlService.ExecuteInspect
            statusAdd("ExecuteInspect Start: " + camNum.ToString)
            FormVProControl.ExecuteInspect(camNum, statusCode)
            statusAdd("ExecuteInspect:" + camNum.ToString + ", retcode=" + statusCode.ToString)
        End Sub
        ' Request results on cam#
        Public Sub GetInspectResults(ByVal camNum As Integer, ByRef statusCode As Integer, ByRef Results As String) Implements IVproControlService.GetInspectResults
            statusAdd("GetInspectResults Start: " + camNum.ToString)
            FormVProControl.GetInspectResults(camNum, statusCode, Results)
            statusAdd("GetInspectResults:" + camNum.ToString + ", retcode=" + statusCode.ToString)
        End Sub
        ' Get service status
        Public Sub GetServiceStatus(ByRef isServiceUp As Boolean) Implements IVproControlService.GetServiceStatus
            Dim serviceMsg As String
            serviceMsg = "Error".ToUpper

            statusAdd("Service is : " + webServiceStatus)

            getWsStatus()

            If webServiceStatus.ToUpper.Contains(serviceMsg) Then
                isServiceUp = False
            Else
                isServiceUp = True
            End If

        End Sub
        ' Delete old saved files.
        Public Sub DeleteOldFiles(ByRef statusCode As Integer, ByVal iDate As Integer, ByVal strFilePath As String, ByVal iTimerInterval As Integer, Optional objComparisonDate As Object = Nothing) Implements IVproControlService.DeleteOldFiles
            FormVProControl.DeleteOldFiles(statusCode, iDate, strFilePath, iTimerInterval, objComparisonDate)
        End Sub

        ' YRM : Is service initialized status
        Public Sub IsServiceInitialized(ByRef bIsServiceInitialized As Boolean) Implements IVproControlService.IsServiceInitialized
            FormVProControl.isServiceInitialized(bIsServiceInitialized)
        End Sub

    End Class

    Public Sub statusAdd(ByVal status As String)
        webServiceStatus = Now.ToString + "." + Now.Millisecond.ToString + " : " + status + vbCrLf + webServiceStatus.Substring(0, Math.Min(webServiceStatus.Length, 10000)) 'limit max...  
    End Sub

    Private myHost
    ' Close service host
    Sub ServiceClose()
        Try
            myHost.close()
            webServiceStatus = String.Format("Service Closed")
            FormVProControl.bIsServiceInitialized = False ' 140224 YRM
        Catch ex As Exception
            webServiceStatus = String.Format("Error closing service") + ", Error=" + ex.Message
        End Try
    End Sub
    ' Start service to host
    Sub ServiceStart(ipaddress As String)
        Dim urlString As String = Nothing
        Dim baseAddress As Uri = Nothing
        Try

            'Dim baseAddress As Uri = New Uri("http://localhost:8080/hello")
            urlString = "http://" + ipaddress + "/vProControl"
            baseAddress = New Uri(urlString)

            ' Create the ServiceHost.
            ' Using host As New ServiceHost(GetType(HelloWorldService), baseAddress)
            myHost = New ServiceHost(GetType(VproControlService), baseAddress)

            ' Enable metadata publishing.
            Dim smb As New ServiceMetadataBehavior()
            smb.HttpGetEnabled = True
            smb.MetadataExporter.PolicyVersion = PolicyVersion.Policy15
            myHost.Description.Behaviors.Add(smb)

            ' Open the ServiceHost to start listening for messages. Since
            ' no endpoints are explicitly configured, the runtime will create
            ' one endpoint per base address for each service contract implemented
            ' by the service.
            myHost.Open()

            webServiceStatus = String.Format("The service is ready at {0}", baseAddress)
        Catch ex As Exception
            webServiceStatus = String.Format("Error opening service at {0}", baseAddress) + ", Error=" + ex.Message
        End Try

    End Sub

    Function getWsStatus() As String
        Return webServiceStatus
    End Function

End Module
