Imports System.Management
Imports System.Runtime.InteropServices

Public Class ServiceParam

    Public Serv As Form1.Service
    Public ItsAdd As Boolean = False
    Public ItsEdit As Boolean = False

    Private Sub ServiceParam_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        '//ДИ старт 25.01.2026  { 
        'ServiceName.Text = Serv.Name
        '//  } ДИ финиш 25.01.2026 

        ExeFile.Text = Serv.ExeFile
        ClusterFiles.Text = Serv.ClusterFiles
        PortAgent.Text = Serv.PortAgent
        PortMngr.Text = Serv.PortMngr
        PortProcessBegin.Text = Serv.PortProcessBegin
        PortProcessEnd.Text = Serv.PortProcessEnd

        CheckBoxDebug.Checked = Serv.Debug

        '//ДИ старт 2026.01.22
        CheckBoxHttp.Checked = Serv.Http
        '//ДИ финиш 2026.01.22

        'DescriptionService.Text = Serv.Description
        DisplayName.Text = Serv.DisplayName

        'Локальная(система)
        'Локальная(служба)
        'Сетевая(служба)

        'If Serv.User = "NT AUTHORITY\NetworkService" Then
        '    RadioButtonUser1.Checked = True
        '    TechUser.Text = TechUser.Items(2)

        'ElseIf Serv.User = "NT AUTHORITY\LocalService" Then
        '    RadioButtonUser1.Checked = True
        '    TechUser.Text = TechUser.Items(1)

        'Else
        If Serv.User = "LocalSystem" Then
            RadioButtonUser1.Checked = True
        Else
            RadioButtonUser2.Checked = True
            Login.Text = Serv.User
            Password.Text = ""
        End If

        If ItsAdd Then
            Text = "Добавление новой службы сервера 1С"
            '//ДИ старт 25.01.2026  { 
            ServiceName.Text = GetNewNameForService()
            '//  } ДИ финиш 25.01.2026 
        ElseIf ItsEdit Then
            Text = "Изменение параметров существующей службы сервера 1С"
            '//ДИ старт 25.01.2026  { 
            ServiceName.Text = Serv.Name
            '//  } ДИ финиш 25.01.2026 
        End If

        ' //ДИ старт 11.09.2026: установка режима запуска службы в ComboBox
        If Serv IsNot Nothing AndAlso Not String.IsNullOrEmpty(Serv.StartType) Then

            Dim mode As String = Serv.StartType.Trim().ToLower()

            Select Case mode
                Case "auto", "automatic"
                    ComboBoxStartType.SelectedIndex = 0 ' Автоматически

                Case "manual", "demand"
                    ComboBoxStartType.SelectedIndex = 1 ' Вручную

                Case "disabled"
                    ComboBoxStartType.SelectedIndex = 2 ' Отключена

                Case Else
                    ComboBoxStartType.SelectedIndex = 0 ' По умолчанию (Автоматически)
            End Select

        Else
            ComboBoxStartType.SelectedIndex = 0
        End If
        ' // } ДИ финиш 11.09.2026

    End Sub

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs) Handles Button6.Click

        OpenFileDialog.FileName = ExeFile.Text
        OpenFileDialog.ShowDialog()
        ExeFile.Text = OpenFileDialog.FileName

    End Sub



    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        FolderBrowserDialog.RootFolder = Environment.SpecialFolder.MyComputer
        FolderBrowserDialog.SelectedPath = ClusterFiles.Text
        FolderBrowserDialog.ShowDialog()
        ClusterFiles.Text = FolderBrowserDialog.SelectedPath

    End Sub

    Private Sub Button3_Click_1(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        PortAgent.Text = (Convert.ToInt32(PortAgent.Text) - 1000).ToString
        PortMngr.Text = (Convert.ToInt32(PortMngr.Text) - 1000).ToString
        PortProcessBegin.Text = (Convert.ToInt32(PortProcessBegin.Text) - 1000).ToString
        PortProcessEnd.Text = (Convert.ToInt32(PortProcessEnd.Text) - 1000).ToString
    End Sub

    Private Sub Button2_Click_1(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        PortAgent.Text = (Convert.ToInt32(PortAgent.Text) + 1000).ToString
        PortMngr.Text = (Convert.ToInt32(PortMngr.Text) + 1000).ToString
        PortProcessBegin.Text = (Convert.ToInt32(PortProcessBegin.Text) + 1000).ToString
        PortProcessEnd.Text = (Convert.ToInt32(PortProcessEnd.Text) + 1000).ToString
    End Sub

    Function GetNewNameForService() As String

        Dim BaseName = "1C:Enterprise 8 Server Agent #"
        Dim i = 1

        GetNewNameForService = BaseName + i.ToString

        While Not CheckNameForService(GetNewNameForService)

            i = i + 1

            GetNewNameForService = BaseName + i.ToString

        End While

    End Function

    Function CheckNameForService(Name As String) As Boolean

        Dim search As New ManagementObjectSearcher("SELECT * FROM Win32_Service WHERE Name = '" + Name + "'")

        CheckNameForService = (search.Get().Count = 0)

    End Function

    Sub SetDefaultColor()

        DisplayName.BackColor = Color.White
        ExeFile.BackColor = Color.White
        ClusterFiles.BackColor = Color.White

        PortAgent.BackColor = Color.White
        PortMngr.BackColor = Color.White
        PortProcessBegin.BackColor = Color.White
        PortProcessEnd.BackColor = Color.White

    End Sub


    Private Sub ButtonSave_Click(sender As System.Object, e As System.EventArgs) Handles ButtonSave.Click

        SetDefaultColor()

        If DisplayName.Text = "" Then
            DisplayName.BackColor = Color.Pink
            MsgBox("Отображаемое имя службы должно быть указано", , Text)
            Return
        End If

        If ItsAdd Then
            Dim search As New ManagementObjectSearcher("SELECT * FROM Win32_Service WHERE DisplayName = '" + DisplayName.Text + "'")

            If search.Get().Count > 0 Then
                DisplayName.BackColor = Color.Pink
                MsgBox("Служба с таким именем уже зарегистрирована. Укажите другое наименование.", , Text)
                Return
            End If
        End If

        If ExeFile.Text = "" Then
            ExeFile.BackColor = Color.Pink
            MsgBox("Путь к исполняемому файлу должен быть указан", , Text)
            Return
        End If

        If Not My.Computer.FileSystem.FileExists(ExeFile.Text) Then
            ExeFile.BackColor = Color.Pink
            MsgBox("Исполняемый файл не существует", , Text)
            Return
        End If

        If ClusterFiles.Text = "" Then
            ClusterFiles.BackColor = Color.Pink
            MsgBox("Путь к каталогу файлов кластера должен быть указан", , Text)
            Return
        End If

        If PortProcessEnd.Text = 0 _
         Or PortProcessBegin.Text = 0 _
         Or PortAgent.Text = 0 _
         Or PortMngr.Text = 0 _
         Then

            PortAgent.BackColor = Color.Pink
            PortMngr.BackColor = Color.Pink
            PortProcessBegin.BackColor = Color.Pink
            PortProcessEnd.BackColor = Color.Pink

            MsgBox("Значения портов агента сервера, менеджера кластера и рабочих процессов должны быть указаны", , Text)
            Return
        End If

        If Not PortProcessEnd.Text > PortProcessBegin.Text Then
            PortProcessEnd.BackColor = Color.Pink
            PortProcessBegin.BackColor = Color.Pink
            MsgBox("Конечный порт рабочего процесса должен быть больше начального порта", , Text)
            Return
        End If

        '//ДИ старт 2026.01.22
        Dim debugParams As String = ""
        If CheckBoxDebug.Checked Then
            debugParams = "-debug"
            If CheckBoxHttp.Checked Then
                Dim debugServerPort = (Convert.ToInt32(PortAgent.Text) + 10).ToString
                debugParams &= " -http -debugServerPort " + debugServerPort
            End If
        End If
        '//ДИ финиш 2026.01.22

        Dim PathNameTemplate = """" + ExeFile.Text + """ {0} -srvc -agent -regport {1} -port {2} -range {3}:{4} -d ""{5}"""
        Dim PathName = String.Format(PathNameTemplate, debugParams, PortMngr.Text,
                      PortAgent.Text, PortProcessBegin.Text, PortProcessEnd.Text, ClusterFiles.Text)

        Dim lpDependencies = "Tcpip" + Char.MinValue + "Dnscache" + Char.MinValue + "lanmanworkstation" + Char.MinValue + "lanmanserver"

        Dim User = ""
        Dim Pwd = ""

        If RadioButtonUser1.Checked Then
            User = "LocalSystem"
        Else
            User = Login.Text
            Pwd = Password.Text
        End If

        Dim TargetServiceName As String = ""

        If ItsAdd Then

            '//ДИ старт 25.01.2026   { 
            If String.IsNullOrWhiteSpace(ServiceName.Text) Then
                TargetServiceName = GetNewNameForService()
            Else
                TargetServiceName = ServiceName.Text
            End If
            '//  } ДИ финиш 25.01.2026 

            If Not ObjTec.Services.ServiceInstaller.InstallService(PathName, TargetServiceName, DisplayName.Text, lpDependencies, User, Pwd) Then
                Dim ErrorCode = Marshal.GetLastWin32Error()
                MsgBox("Ошибка установки сервиса " + Form1.GetErrorDescription(ErrorCode), MsgBoxStyle.Critical, Text)
                Return
            End If

        ElseIf ItsEdit Then

            TargetServiceName = Serv.Name

            If ObjTec.Services.ServiceInstaller.ChangeServiceParameters(PathName, Serv.Name, DisplayName.Text, lpDependencies, User, Pwd) Then
                Dim sc = New System.ServiceProcess.ServiceController(Serv.Name)
                If sc.Status = ServiceProcess.ServiceControllerStatus.Running Then
                    If MsgBox("Параметры успешно изменены, но служба в настоящий момент работает." +
                           vbNewLine + "Перезапустить службу для применения изменений?", MsgBoxStyle.YesNo, Text) = MsgBoxResult.Yes Then
                        sc.Stop()
                        sc.WaitForStatus(ServiceProcess.ServiceControllerStatus.Stopped)
                        sc.Start()
                        sc.WaitForStatus(ServiceProcess.ServiceControllerStatus.Running)
                    End If
                End If
            Else
                Dim ErrorCode = Marshal.GetLastWin32Error()
                MsgBox("Ошибка изменения параметров сервиса " + Form1.GetErrorDescription(ErrorCode), MsgBoxStyle.Critical, Text)
                Return
            End If

        End If

        '//ДИ старт 11.09.2026: Установка типа запуска службы (Авто / Вручную / Отключена)
        If Not String.IsNullOrEmpty(TargetServiceName) Then
            Dim startTypeArg As String = ""

            Select Case ComboBoxStartType.SelectedIndex
                Case 0
                    startTypeArg = "auto"     ' Автоматически
                Case 1
                    startTypeArg = "demand"   ' Вручную
                Case 2
                    startTypeArg = "disabled" ' Отключена
            End Select

            If Not String.IsNullOrEmpty(startTypeArg) Then
                Try
                    Dim proc As New Process()
                    Dim startInfo As New ProcessStartInfo()

                    startInfo.FileName = "sc.exe"
                    startInfo.Arguments = String.Format("config ""{0}"" start= {1}", TargetServiceName, startTypeArg)
                    startInfo.CreateNoWindow = True
                    startInfo.UseShellExecute = False

                    proc.StartInfo = startInfo
                    proc.Start()
                    proc.WaitForExit()
                Catch ex As Exception
                    MsgBox("Ошибка при изменении типа запуска службы: " & ex.Message, MsgBoxStyle.Exclamation, Text)
                End Try
            End If
        End If
        '//ДИ финиш 11.09.2026

        Close()

    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Close()
    End Sub

    Private Sub CheckBoxDebug_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxDebug.CheckedChanged

    End Sub

    Private Sub PortAgent_MaskInputRejected(sender As Object, e As MaskInputRejectedEventArgs) Handles PortAgent.MaskInputRejected

    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click

    End Sub

End Class