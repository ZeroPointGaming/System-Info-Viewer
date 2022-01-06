Imports System.Text.RegularExpressions
Imports System.Management
Imports System.Net
Imports System.Diagnostics
Imports System.Text

Public Class Form1
    Public ipaddr As String
    'Public cc As String
    'Public cn As String
    'Public rc As String
    'Public rn As String
    'Public city As String
    'Public lat As Double
    'Public longitude As Double
    Public Manufact As String
    Public SerialNum As String
    Public ModelName As String
    Public pxWidth As String = SystemInformation.PrimaryMonitorSize.Width.ToString
    Public pxHeight As String = SystemInformation.PrimaryMonitorSize.Height.ToString
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ListBox1.Items.Add("OS Information: " + My.Computer.Info.OSFullName + " / " + My.Computer.Info.OSPlatform + " / " + My.Computer.Info.OSVersion)
        ListBox1.Items.Add("User Information: " + SystemInformation.UserDomainName + "/" + SystemInformation.UserName)
        ListBox1.Items.Add("DNS Hostname: " + Net.Dns.GetHostName)
        ListBox1.Items.Add("Local Time/Timezone: " + My.Computer.Clock.LocalTime + " / " + System.TimeZone.CurrentTimeZone.StandardName)
        ListBox1.Items.Add("Main Display Resolution: " + SystemInformation.PrimaryMonitorSize.Width.ToString + "x" + SystemInformation.PrimaryMonitorSize.Height.ToString)
        ListBox1.Items.Add("Boot Mode: " + SystemInformation.BootMode.ToString)
        ListBox1.Items.Add("Monitor Count: " + SystemInformation.MonitorCount.ToString)
        ListBox1.Items.Add("Mouse Speed: " + SystemInformation.MouseSpeed.ToString)
        ListBox1.Items.Add("Network Capable: " + SystemInformation.Network.ToString)
        ListBox1.Items.Add("Uses Security: " + SystemInformation.Secure.ToString)
        ListBox1.Items.Add("Screen Orientation: " + SystemInformation.ScreenOrientation.ToString)
        ListBox1.Items.Add("CPU Cores: " + Environment.ProcessorCount.ToString)
        ListBox1.Items.Add("Current Clipboard Contet: " + My.Computer.Clipboard.GetText)
        ListBox1.Items.Add("Is 64bit: " + Environment.Is64BitOperatingSystem.ToString)
        ListBox1.Items.Add("Graphics Card Name: " + GetGraphicsCardName())
        ListBox1.Items.Add("Processing Core Name: " + GetProcessorName())
        ListBox1.Items.Add("Motherboard: " + GetMotherboardName() & " / " & GetMotherboardModel() & " / " & GetMotherboardSerial())
        Dim ItemsList As String = ListBox1.Items.ToString
        ListBox1.Items.Add("IP Address: " + IpAddress())
        'ListBox1.Items.Add("Country Code: " + cc)
        'ListBox1.Items.Add("Counrty Name: " + cn)
        'ListBox1.Items.Add("Region Name: " + rn)
        'ListBox1.Items.Add("Region Code: " + rc)
        'ListBox1.Items.Add("City: " + city)
        'ListBox1.Items.Add(lat)
        'ListBox1.Items.Add(longitude)
    End Sub
    Public Structure DatosIPExterna
        Dim IP As String
        Dim CountryCode As String
        Dim CountryName As String
        Dim RegionCode As String
        Dim RegionName As String
        Dim City As String
        Dim Latitude As Double
        Dim Longitude As Double
    End Structure

    Private Function IpAddress() As String
        Using wc As New Net.WebClient
            Return System.Text.Encoding.UTF8.GetString(wc.DownloadData("http://tools.feron.it/php/ip.php"))
        End Using
    End Function

    Private Function BuildTheListBoxItemsToSend(ByVal lb As ListBox) As String
        Dim sb As New System.Text.StringBuilder
        For Each item As String In ListBox1.Items
            sb.Append(item & "  -|-  ")
        Next
        Return sb.ToString
    End Function
    Private Function GetGraphicsCardName() As String
        Dim GraphicsCardName = String.Empty
        Try
            Dim WmiSelect As New ManagementObjectSearcher("SELECT * FROM Win32_VideoController")
            For Each WmiResults As ManagementObject In WmiSelect.Get()
                GraphicsCardName = WmiResults.GetPropertyValue("Name").ToString
                If (Not String.IsNullOrEmpty(GraphicsCardName)) Then
                    Exit For
                End If
            Next
        Catch err As ManagementException
            MessageBox.Show(err.Message)
        End Try
        Return GraphicsCardName
    End Function
    Private Function GetProcessorName() As String
        Dim ProcessorName = String.Empty
        Try
            Dim WmiSelect As New ManagementObjectSearcher("SELECT * FROM Win32_Processor")
            For Each WmiResults As ManagementObject In WmiSelect.Get()
                ProcessorName = WmiResults.GetPropertyValue("Name").ToString
                If (Not String.IsNullOrEmpty(ProcessorName)) Then
                    Exit For
                End If
            Next
        Catch err As ManagementException
            MessageBox.Show(err.Message)
        End Try
        Return ProcessorName
    End Function
    Private Function GetMotherboardName() As String
        Dim MotherboardName = String.Empty
        Try
            Dim WmiSelect As New ManagementObjectSearcher("SELECT * FROM Win32_BaseBoard")

            For Each WmiResults As ManagementObject In WmiSelect.Get()
                MotherboardName = WmiResults.GetPropertyValue("Manufacturer").ToString
                'MotherboardName = WmiResults.GetPropertyValue("Model").ToString
                'MotherboardName = WmiResults.GetPropertyValue("SerialNumber").ToString


                If (Not String.IsNullOrEmpty(MotherboardName)) Then
                    Exit For
                End If
            Next
        Catch err As ManagementException
            MessageBox.Show(err.Message)
        End Try
        Return MotherboardName
    End Function
    Private Function GetMotherboardSerial() As String
        Dim MotherboardName = String.Empty
        Try
            Dim WmiSelect As New ManagementObjectSearcher("SELECT * FROM Win32_BaseBoard")

            For Each WmiResults As ManagementObject In WmiSelect.Get()
                'MotherboardName = WmiResults.GetPropertyValue("Manufacturer").ToString
                'MotherboardName = WmiResults.GetPropertyValue("Model").ToString
                MotherboardName = WmiResults.GetPropertyValue("Product").ToString


                If (Not String.IsNullOrEmpty(MotherboardName)) Then
                    Exit For
                End If
            Next
        Catch err As ManagementException
            MessageBox.Show(err.Message)
        End Try
        Return MotherboardName
    End Function
    Private Function GetMotherboardModel() As String
        Dim MotherboardName = String.Empty
        Try
            Dim WmiSelect As New ManagementObjectSearcher("SELECT * FROM Win32_BaseBoard")

            For Each WmiResults As ManagementObject In WmiSelect.Get()
                'MotherboardName = WmiResults.GetPropertyValue("Manufacturer").ToString
                MotherboardName = WmiResults.GetPropertyValue("Name").ToString
                'MotherboardName = WmiResults.GetPropertyValue("SerialNumber").ToString


                If (Not String.IsNullOrEmpty(MotherboardName)) Then
                    Exit For
                End If
            Next
        Catch err As ManagementException
            MessageBox.Show(err.Message)
        End Try
        Return MotherboardName
    End Function

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        My.Computer.Clipboard.SetText(ListBox1.SelectedItem.ToString)
    End Sub
End Class