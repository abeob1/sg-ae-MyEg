Option Explicit On
Option Strict Off
Imports System.Windows.Forms
Imports System.ServiceProcess

Module SubMain
    Public Sub Main()

        Dim obj As New Add_on
        Application.Run()

        'Dim f As New DebugForm
        'f.ShowDialog()

        'Dim Args As String() = Environment.GetCommandLineArgs()
        'If Args.Length = 1 Then
        '    Dim obj As New Add_on
        '    Application.Run()
        'Else
        '    If Args(1) = "-service" Then
        '        Dim ServicesToRun As ServiceBase()
        '        ServicesToRun = New ServiceBase() {New Service1()}
        '        ServiceBase.Run(ServicesToRun)
        '    Else
        '        Dim obj As New Add_on
        '        Application.Run()
        '    End If
        'End If


        'Application.EnableVisualStyles()
        'Application.SetCompatibleTextRenderingDefault(False)
        'Application.Run(New frmItemStructure())


        'Dim Args As String() = Environment.GetCommandLineArgs()

        'If Args.Length = 1 Then
        '    Application.EnableVisualStyles()
        '    Application.SetCompatibleTextRenderingDefault(False)
        '    Application.Run(New frmQCOpenOrderList())
        'ElseIf Args(1) = "-service" Then
        '    Dim ServicesToRun As ServiceBase()
        '    ServicesToRun = New ServiceBase() {New Service1()}
        '    ServiceBase.Run(ServicesToRun)
        'ElseIf Args(1) = "-addon" Then
        '    Dim obj As New Add_on
        '    Application.Run()
        'End If
    End Sub
End Module
