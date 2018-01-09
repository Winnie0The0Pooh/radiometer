
Imports System.Net.Sockets
Imports System.Text

'Imports System
Imports System.IO
Imports System.Windows.Forms.DataVisualization.Charting

Imports System.Net
Imports System.Speech.Synthesis

Public Class Form1

    Public Shared addr As String = "192.168.1.155"
    Public Shared addrWB As String = "192.168.1.150"
    Public Shared port As Integer = 8080

    Dim path1 As String = "e:\3\rd.txt"
    Dim path2 As String = "e:\3\rd_error.txt"
    Dim path3 As String = "e:\3\rd_status.txt"
    Dim sw As StreamWriter = New StreamWriter(path1, True)
    Dim swerror As StreamWriter = New StreamWriter(path2, True)
    'Dim status As StreamWriter = New StreamWriter(path3, True)

    Public Shared mkr As New Queue(Of Integer)
    Public Shared TempC As New Queue(Of Double)
    Public Shared Humi As New Queue(Of Double)
    Public Shared HiVolt As New Queue(Of Double)

    Public Shared Hgmm As New Queue(Of Integer)
    Public Shared Hgmmr As New Queue(Of Integer)

    Public Shared CO2 As New Queue(Of Integer)
    Public Shared CO2r As New Queue(Of Integer)

    Public Shared Temp0 As New Queue(Of Double)
    Public Shared Temp1 As New Queue(Of Double)
    Public Shared Temp2 As New Queue(Of Double)
    Public Shared Temp3 As New Queue(Of Double)
    Public Shared Temp4 As New Queue(Of Double)

    Public Shared Temp0r As New Queue(Of Double)
    Public Shared Temp1r As New Queue(Of Double)
    Public Shared Temp2r As New Queue(Of Double)
    Public Shared Temp3r As New Queue(Of Double)
    Public Shared Temp4r As New Queue(Of Double)

    Public Shared qx As New Queue(Of String)
    Public Shared qxwb As New Queue(Of String)
    Public Shared diag As New Queue(Of String)
    Public Shared wbraw As New Queue(Of String)

    Public CLIENT As Class_client
    Public CLIENTwb As Class_client

    Function RetrieveLinkerTimestamp(ByVal filePath As String) As DateTime
        Const PeHeaderOffset As Integer = 60
        Const LinkerTimestampOffset As Integer = 8
        Dim b(2047) As Byte
        Dim s As Stream
        Try
            s = New FileStream(filePath, FileMode.Open, FileAccess.Read)
            s.Read(b, 0, 2048)
        Finally
            If Not s Is Nothing Then s.Close()
        End Try
        Dim i As Integer = BitConverter.ToInt32(b, PeHeaderOffset)
        Dim SecondsSince1970 As Integer = BitConverter.ToInt32(b, i + LinkerTimestampOffset)
        Dim dt As New DateTime(1970, 1, 1, 0, 0, 0)
        dt = dt.AddSeconds(SecondsSince1970)
        dt = dt.AddHours(TimeZone.CurrentTimeZone.GetUtcOffset(dt).Hours)
        Return dt
    End Function

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim fs As String = Application.ExecutablePath
        Me.Text = Me.Text & " Build date: " & RetrieveLinkerTimestamp(fs)


        'memoryUsed = Process.WorkingSet64 / (1024.0F * 1024.0F)



        Me.Size = My.Settings.w_size
        Me.Location = My.Settings.w_loc


        Chart1.Series(0).Name = "Rad"
        Chart1.Series.Add("HV")
        Chart1.Series("HV").YAxisType = AxisType.Secondary

        For j = 0 To Chart1.Series.Count - 1
            Chart1.Series(j).ChartType = DataVisualization.Charting.SeriesChartType.Line
            Chart1.Series(j).BorderWidth = 3
        Next

        Chart1.ChartAreas("ChartArea1").AxisY.IsStartedFromZero = False
        Chart1.ChartAreas("ChartArea1").AxisY2.IsStartedFromZero = False

        Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time"
        'Chart2.ChartAreas("ChartArea1").AxisX.MajorGrid.Interval = 1
        'Chart2.ChartAreas("ChartArea1").AxisX.LabelStyle.Interval = 1
        Chart1.ChartAreas("ChartArea1").AxisY.Title = "мкр"
        Chart1.ChartAreas("ChartArea1").AxisY2.Title = "HV, v"

        Chart2.Series(0).Name = "Temp"
        Chart2.Series.Add("Humi")
        Chart2.Series("Humi").YAxisType = AxisType.Secondary


        For j = 0 To Chart2.Series.Count - 1
            Chart2.Series(j).ChartType = DataVisualization.Charting.SeriesChartType.Line
            Chart2.Series(j).BorderWidth = 3
        Next

        Chart2.ChartAreas("ChartArea1").AxisY.IsStartedFromZero = False
        Chart2.ChartAreas("ChartArea1").AxisY2.IsStartedFromZero = False

        Chart2.ChartAreas("ChartArea1").AxisX.Title = "Time"
        'Chart2.ChartAreas("ChartArea1").AxisX.MajorGrid.Interval = 1
        'Chart2.ChartAreas("ChartArea1").AxisX.LabelStyle.Interval = 1
        Chart2.ChartAreas("ChartArea1").AxisY.Title = "Temp, C"
        Chart2.ChartAreas("ChartArea1").AxisY2.Title = "Humidity, %"

        Chart3.Series(0).Name = "Atm Press"

        For j = 0 To Chart3.Series.Count - 1
            Chart3.Series(j).ChartType = DataVisualization.Charting.SeriesChartType.Line
            Chart3.Series(j).BorderWidth = 3
        Next

        Chart3.ChartAreas("ChartArea1").AxisY.IsStartedFromZero = False

        Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time"
        Chart1.ChartAreas("ChartArea1").AxisY.Title = "мм рт.ст."

        Chart5.Series(0).Name = "CO2"
        For j = 0 To Chart5.Series.Count - 1
            Chart5.Series(j).ChartType = DataVisualization.Charting.SeriesChartType.Line
            Chart5.Series(j).BorderWidth = 3
        Next

        Chart8.Series(0).Name = "Temp0"
        Chart8.Series.Add("Temp1")
        Chart8.Series.Add("Temp2")
        Chart8.Series.Add("Temp3")
        Chart8.Series.Add("Temp4")
        For j = 0 To Chart8.Series.Count - 1
            Chart8.Series(j).ChartType = DataVisualization.Charting.SeriesChartType.Line
            Chart8.Series(j).BorderWidth = 3
        Next


        CLIENT = New Class_client ' Инициализируем класс клиента. Иначе форма не закроется, если не подключится хотя бы раз.
        CLIENTwb = New Class_client
        s1()
        s1wb()

        'Dim ex As Boolean, dl As Long
        'ex = File.Exists(path3)
        'dl = FileSystem.FileLen(path3)



        If (File.Exists(path3)) Then
            If (FileSystem.FileLen(path3) > 100) Then
                'Using (BinaryReader reader = New BinaryReader(File.Open(path3, FileMode.Open)))
                Dim j As Integer, i As Integer

                Using fsr As New FileStream(path3, FileMode.Open)
                    Using w As New BinaryReader(fsr)
                        j = w.ReadInt32


                        For i = 0 To j
                            qx.Enqueue(w.ReadString)
                            mkr.Enqueue(w.ReadInt32)
                            TempC.Enqueue(w.ReadDouble)
                            Humi.Enqueue(w.ReadDouble)
                            HiVolt.Enqueue(w.ReadDouble)
                            Hgmmr.Enqueue(w.ReadInt32)
                            CO2r.Enqueue(w.ReadInt32)
                            Temp0r.Enqueue(w.ReadDouble)
                            Temp1r.Enqueue(w.ReadDouble)
                            Temp2r.Enqueue(w.ReadDouble)
                            Temp3r.Enqueue(w.ReadDouble)
                            Temp4r.Enqueue(w.ReadDouble)
                        Next
                    End Using
                End Using
                msg("Read " & j & " records" & vbCrLf)
            End If
        End If

        TextBox12.Text = "Запущена " & Now()



        Timer1.Interval = 10 * 60 * 1000 ' 10 минут
        Timer1.Enabled = True

    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        My.Settings.w_size = Me.Size 'New Size(Me.Width, Me.Height)
        My.Settings.w_loc = Me.Location 'New Size(Me.Location)

        My.Settings.Save()

        s2() 'при закрытии формы отключаем потоки, закрываем треды и т.д.
        s2wb()

        swerror.WriteLine(Now & " Выполнение программы завершено.")
        swerror.Flush()

    End Sub

    Private Sub s1()
        ' Подключение к серверу 

        CLIENT = New Class_client ' Инициализируем класс клиента
        ' Установка событий класса нужным процедурам. Описание событий в классе.
        AddHandler CLIENT.OnRead, AddressOf Me.OnRead_Invoke
        AddHandler CLIENT.Disconnected, AddressOf Me.Disconnect_Invoke
        AddHandler CLIENT.messege, AddressOf Me.msg_Invoke

        Dim p As Long = 0
        Do
            msg(vbCrLf & "Попытка " & p & vbCrLf)
            p = p + 1
        Loop Until (CLIENT.Connect(addr, port, AscW("+")))
        msg("Подключено" & vbCrLf)
        swerror.WriteLine(Now & " Поток подключен")
        swerror.Flush()
    End Sub
    Private Sub s1wb()
        ' Подключение к серверу 

        CLIENTwb = New Class_client ' Инициализируем класс клиента
        ' Установка событий класса нужным процедурам. Описание событий в классе.
        AddHandler CLIENTwb.OnRead, AddressOf Me.OnRead_Invokewb
        AddHandler CLIENTwb.Disconnected, AddressOf Me.Disconnect_Invokewb
        AddHandler CLIENTwb.messege, AddressOf Me.msg_Invoke

        Dim p As Long = 0
        Do
            msg(vbCrLf & "WBПопытка " & p & vbCrLf)
            p = p + 1
        Loop Until (CLIENTwb.Connect(addrwb, port, AscW("+")))
        msg("WBПодключено" & vbCrLf)
        swerror.WriteLine(Now & " WBПоток подключен")
        swerror.Flush()
    End Sub
    Private Sub s2()
        'Отключение от сервера, разрушение потока 
        CLIENT.Disconnect()
        RemoveHandler CLIENT.OnRead, AddressOf Me.OnRead_Invoke
        RemoveHandler CLIENT.Disconnected, AddressOf Me.Disconnect_Invoke
        RemoveHandler CLIENT.messege, AddressOf Me.msg_Invoke
        CLIENT.client_Thread.Abort()
        CLIENT = Nothing
        swerror.WriteLine(Now & " Отключение от потока")
        swerror.Flush()

    End Sub
    Private Sub s2wb()
        'Отключение от сервера, разрушение потока 
        CLIENTwb.Disconnectwb()
        RemoveHandler CLIENTwb.OnRead, AddressOf Me.OnRead_Invokewb
        RemoveHandler CLIENTwb.Disconnected, AddressOf Me.Disconnect_Invokewb
        RemoveHandler CLIENTwb.messege, AddressOf Me.msg_Invoke
        CLIENTwb.client_Thread.Abort()
        CLIENTwb = Nothing
        swerror.WriteLine(Now & " WBОтключение от потока")
        swerror.Flush()

    End Sub

    Private Sub OnRead(ByVal data As String)
        ' Процедура расшифровки принятых данных
        '249302193 - 40 - 60 - 12 mkr\hour
        '249302193, 40 imp, 60 sec, 12 mkr\hour
        '862557, 103 imp, 22.54 sec, 82 mkr\hour
        '12148898, 73 imp, 121.80 sec, 10 mkr\hour,  Temp: 3.26 C, Humidity: 61.48 %, 0
        '495617, 0 imp, 120.17 sec, 0 mkr\hour, 13.8299980163 C, 86.97 % 0 v
        '132674, 0 imp, 120.94 sec, 0 mkr\hour, -40.01 C, -4.69 %, 5 v
        '


        Dim regexp As New System.Text.RegularExpressions.Regex("[ -,]")

        Dim regexp1 As New System.Text.RegularExpressions.Regex("\d*, \d* imp, \d{1,3}\.\d{1,2} sec, \d{1,5} mkr\\hour, -?\d{1,2}\.\d{1,2} C, -?\d{1,2}\.\d{1,2} %, \d{1,4} v")
        '-?\d{1,3}\.\d{1,2}){0,}
        'Dim regexp1 As New System.Text.RegularExpressions.Regex("\d* - \d* - \d* - \d* mkr\\hour")

        Dim s() As String

        Dim CoeffHV = 100 * 5 / 1024 * (393 / 433) 'поправочный коэфф по мультиметру


        If regexp1.IsMatch(data) Then

            s = regexp.Split(data)

            Dim ap As String = ""
            For i = 0 To s.GetUpperBound(0)
                ap = ap & i & " " & s(i) & vbCrLf
            Next
            'Debug.Print(ap)

            mkr.Enqueue(CInt(s(8)))
            TempC.Enqueue(CDbl(s(11).Replace(".", ",")))
            Humi.Enqueue(CDbl(s(14).Replace(".", ",")))
            HiVolt.Enqueue(Math.Round((CDbl(s(18).Replace(".", ","))) * CoeffHV))


            TextBox2.Text = mkr.Last
            TextBox10.Text = HiVolt.Last

            If HiVolt.Last < 370 Then
                TextBox10.ForeColor = Color.Red
            Else
                TextBox10.ForeColor = Color.Green
            End If

            If Humi.Last >= 0 Then
                TextBox3.Text = Humi.Last
                TextBox8.Text = TempC.Last
                TextBox9.Text = "Data OK."
                TextBox9.ForeColor = Color.Green

            Else
                TextBox3.Text = "No data"
                TextBox8.Text = "No data"
                TextBox9.Text = "Sensors disconnected."
                TextBox9.ForeColor = Color.Red
            End If


            Dim ave As Integer = mkr.Average
            TextBox4.Text = ave
            TextBox6.Text = mkr.Min
            TextBox7.Text = mkr.Max
            Dim m() = mkr.ToArray

            Dim r As Double, sum As Double = 0
            Dim ct As Double = mkr.Count

            For j = m.GetLowerBound(0) To m.GetUpperBound(0)
                r = m(j) - ave
                r = r * r
                sum = sum + r
            Next

            sum = sum / ct
            Dim sko As Double = Math.Round(Math.Sqrt(sum), 2)
            TextBox5.Text = sko

            qx.Enqueue(Format(TimeOfDay, "HH:mm:ss"))

            Hgmmr.Enqueue(Hgmm.Average)

            CO2r.Enqueue(CO2.Average)

            Temp0r.Enqueue(Temp0.Average)
            Temp1r.Enqueue(Temp1.Average)
            Temp2r.Enqueue(Temp2.Average)
            Temp3r.Enqueue(Temp3.Average)
            Temp4r.Enqueue(Temp4.Average)

            Chart3.Series(0).Points.DataBindXY(qx, Hgmmr)

            Chart1.Series(0).Points.DataBindXY(qx, mkr)
            Chart1.Series("HV").Points.DataBindXY(qx, HiVolt)

            Chart2.Series("Temp").Points.DataBindXY(qx, TempC)
            Chart2.Series("Humi").Points.DataBindXY(qx, Humi)

            Chart5.Series(0).Points.DataBindXY(qx, CO2r)

            If Temp0r.Count > 0 Then
                Chart8.Series("Temp0").Points.DataBindXY(qx, Temp0r)
                Chart8.Series("Temp1").Points.DataBindXY(qx, Temp1r)
                Chart8.Series("Temp2").Points.DataBindXY(qx, Temp2r)
                Chart8.Series("Temp3").Points.DataBindXY(qx, Temp3r)
                Chart8.Series("Temp4").Points.DataBindXY(qx, Temp4r)
            End If

            If mkr.Count > 2400 Then
                qx.Dequeue()
                mkr.Dequeue()
                TempC.Dequeue()
                Humi.Dequeue()
                HiVolt.Dequeue()
                Hgmmr.Dequeue()
                CO2r.Dequeue()

                Temp0r.Dequeue()
                Temp1r.Dequeue()
                Temp2r.Dequeue()
                Temp3r.Dequeue()
                Temp4r.Dequeue()

            End If

            Dim P As Process
            P = Process.GetCurrentProcess
            Dim memoryUsed As Long
            memoryUsed = P.WorkingSet64

            diag.Enqueue(Now() & " Записей " & qx.Count & " Память: " & memoryUsed & vbCrLf)

            If diag.Count > 20 Then
                diag.Dequeue()
            End If


            Dim diagstr As String
            diagstr = ""

            For j = 0 To diag.Count - 1
                diagstr = diagstr & diag(j)
            Next
            TextBox13.Text = diagstr

            'TextBox13.Text = TextBox13.Text & " Память: " & memoryUsed & vbCrLf

            Try
                Using WebClientwebClient As New WebClient()
                    '// Адрес ресурса, к которому выполняется запрос
                    Dim url As String
                    'Dim otvet As String
                    url = "http://esp/data"
                    Dim responce As String
                    'Для GET запроса надо создать объект класса WebClient
                    'и выполнить его метод DownloadString().
                    ' Для запроса есть специальное свойство QueryString,
                    ' с помощью которго можно добавить параметры запроса в виде пар ключ,значение. 
                    'http://esp/data?atm_pressure=748&Out_temp=-12.7&Out_wet=72.7

                    '// Выполняем запрос по адресу и получаем ответ в виде строки
                    WebClientwebClient.QueryString.Add("atm_pressure", Int(Hgmm.Last).ToString)
                    WebClientwebClient.QueryString.Add("Out_temp", Math.Round(TempC.Last, 0).ToString()) 'Int(TempC.Last).ToString
                    WebClientwebClient.QueryString.Add("Out_wet", Int(Humi.Last).ToString)
                    responce = WebClientwebClient.DownloadString(url)
                    'WebClientwebClient.Dispose()
                End Using

            Catch ex As Exception
                msg(ex.Message) ' Сообщаем об ошибке
            End Try
        Else
            swerror.WriteLine(Now & " " & data)
            swerror.Flush()
        End If

        msg(data)

    End Sub

    Private Sub OnReadwb(ByVal data As String)
        ' Процедура расшифровки принятых данных
        ' 122014676 T 37 P 98391 737 CO2 801 T 25.04 H 38.54 CO I 82 40 SMO 18 H2 32 LPG 12 NH3 28 SPI 38 Mag X 227 Y 148 Z -298 Grav X 63 Y -269 Z 28 Gyro X 78 Y 36 Z -76 mem 188
        '3237495 T 30 P 99405 745 CO2 1074 T 25.66 H 32.40 CO N 218 223 SMO 191 H2 207 LPG 220 NH3 182 SPI 197 Mag X -148 Y 491 Z -142 Grav X 64 Y -275 Z -30 Gyro X 63 Y 21 Z 9 D0 20.00 D1 20.50 D2 20.50 D3 20.00 D4 20.50 mem 118
        '3237495 T 30 P 99405 745 CO2 1074 T 25.66 H 32.40 CO N 218 223 SMO 191 H2 207 LPG 220 NH3 182 SPI 197 Mag X -148 Y 491 Z -142 Grav X 64 Y -275 Z -30 Gyro X 63 Y 21 Z 9 D0 20.00 D1 20.50 D2 20.50 D3 20.00 D4 20.50 mem 1189
        '24.11.2013 0:03:56 >> 
        '\d* T \d* P \d* \d* CO2 \d* T \d*\.?\d CO (I|N) \d* \d* SMO \d* H2 \d* LPG \d* NH3 \d* SPI \d* Mag X \d* Y \d* Z \d* Grav X \d* Y \d* Z \d* Gyro X \d* y \d* Z \d* mem \d* 

        Dim regexp As New System.Text.RegularExpressions.Regex("[ ,]")
        '        Dim regexp1 As New System.Text.RegularExpressions.Regex("\d* T \d* P \d* \d* CO2 \d* T \d*\.?\d CO (I|N) \d* \d* SMO \d* H2 \d* LPG \d* NH3 \d* SPI \d* Mag X \d* Y \d* Z \d* Grav X \d* Y \d* Z \d* Gyro X \d* y \d* Z \d* mem \d*")
        Dim regexp1 As New System.Text.RegularExpressions.Regex("^\d{1,10} T \d{2} P \d{5,6} \d{3} CO2 \d{1,4} T -?\s?\d{1,2}\.\d{2} H \d{2}\.\d{2} CO (I|N) \d{1,4} \d{1,4} SMO \d{1,4} H2 \d{1,4} LPG \d{1,4} NH3 \d{1,4} SPI \d{1,4} Mag X -?\d{1,5} Y -?\d{1,5} Z -?\d{1,5} Grav X -?\d{1,5} Y -?\d{1,5} Z -?\d{1,5} Gyro X -?\d{1,5} Y -?\d{1,5} Z -?\d{1,5}( D\d -?\d{1,3}\.\d{1,2}){0,} L \d{1,6}\.\d{2} mem \d{1,4}")
        '23161567 T 37 P 99509 746 CO2 913 T 26.33 H 24.13 CO I 91 40 SMO 58 H2 85 LPG 57 NH3 87 SPI 79 Mag X 313 Y -11 Z 41 Grav X 1 Y -281 Z -28 Gyro X 81 Y 52 Z -55 D0 24.25 D1 24.00 D2 24.25 D3 23.75 D4 23.25 L 2393 mem 1179
        Dim s() As String
        'Dim strw1 As String
        wbraw.Enqueue(data)
        If wbraw.Count > 10 Then
            wbraw.Dequeue()
        End If

        Dim diagstr As String
        diagstr = ""

        For j = 0 To wbraw.Count - 1
            diagstr = diagstr & wbraw(j)
        Next
        TextBox14.Text = diagstr

        If regexp1.IsMatch(data) Then 'And data.Length < 250

            'strw1 = Now & " " & data

            'sw1.WriteLine(strw1)
            'sw1.Flush()

            'TempIntTB.BackColor = Color.LightGreen

            s = regexp.Split(data)

            Dim ap As String = ""
            For i = 0 To s.GetUpperBound(0)
                ap = ap & i & " " & s(i) & vbCrLf
            Next
            'Debug.Print(ap)


            'Tempint.Enqueue(CInt(s(2)))
            'TempIntTB.Text = Tempint.Last

            Hgmm.Enqueue(CInt(s(5)))
            TextBox11.Text = Hgmm.Last
            mmHg.Text = Hgmm.Last

            If CInt(s(7)) = 0 Then
                CO2.Enqueue(CO2.Last)
                CO2TB.Text = "ND"
            Else
                CO2.Enqueue(CInt(s(7)))
                CO2TB.Text = CO2.Last
            End If

            'Tempext.Enqueue(CDbl(s(9).Replace(".", ","))) ' внешняя температура из SHT11
            'TempExtTB.Text = Tempext.Last

            'Humi.Enqueue(CDbl(s(11).Replace(".", ","))) ' влажность
            'HumiTB.Text = Humi.Last

            'If s(13) = "N" Then
            '    'If COstat.Last = 0 Then nold = 1
            '    COstat.Enqueue(1) ' статус датчика CO  N нагрев, I измерение
            'Else
            '    'If COstat.Last = 1 Then nold = 1
            '    COstat.Enqueue(0)
            'End If

            'COmeas.Enqueue(CInt(s(14)))
            'COMeasTB.Text = COmeas.Last

            'COcurr.Enqueue(CInt(s(15)))
            'COcurrTB.Text = COcurr.Last

            'Smog.Enqueue(CInt(s(17)))
            'SmogTB.Text = Smog.Last

            'H2.Enqueue(CInt(s(19)))
            'H2TB.Text = H2.Last

            'LPG.Enqueue(CInt(s(21)))
            'lpgtb.Text = LPG.Last

            'NH3.Enqueue(CInt(s(23)))
            'NH3TB.Text = NH3.Last

            'Spirt.Enqueue(CInt(s(25)))
            'SpirtTB.Text = Spirt.Last

            'MagX.Enqueue(CInt(s(28)))
            'MagXTB.Text = MagX.Last

            'MagY.Enqueue(CInt(s(30)))
            'MagYTB.Text = MagY.Last

            'MagZ.Enqueue(CInt(s(32)))
            'MagZTB.Text = MagZ.Last

            'MagM.Enqueue(Math.Sqrt(Math.Pow(MagX.Last, 2) + Math.Pow(MagY.Last, 2) + Math.Pow(MagZ.Last, 2)))

            'GravX.Enqueue(CInt(s(35)))
            'GravXTB.Text = GravX.Last

            'GravY.Enqueue(CInt(s(37)))
            'GravYTB.Text = GravY.Last

            'GravZ.Enqueue(CInt(s(39)))
            'GravZTB.Text = GravZ.Last

            'GravM.Enqueue(Math.Sqrt(Math.Pow(GravX.Last, 2) + Math.Pow(GravY.Last, 2) + Math.Pow(GravZ.Last, 2)))

            'GyroX.Enqueue(CInt(s(42)))
            'GyroXTB.Text = GyroX.Last

            'GyroY.Enqueue(CInt(s(44)))
            'GyroYTB.Text = GyroY.Last

            'GyroZ.Enqueue(CInt(s(46)))
            'GyroZTB.Text = GyroZ.Last

            'GyroM.Enqueue(Math.Sqrt(Math.Pow(GyroX.Last, 2) + Math.Pow(GyroY.Last, 2) + Math.Pow(GyroZ.Last, 2)))

            'датчиков температуры может быть от 0 и далее. Надо определить - есть ли данные от датчиков темп и сколько их.



            '47:         D0()
            '48 32.19
            '49:         D1()
            '50 29.69
            '51:         D2()
            '52 30.37
            '53:         D3()
            '54 31.00
            '55:         D4()
            '56 28.56

            'If s(47) = "D0" Then
            Temp0.Enqueue(CDbl(s(48).Replace(".", ",")))
            Temp0TB.Text = Temp0.Last


            Temp1.Enqueue(CDbl(s(50).Replace(".", ",")))
                Temp1TB.Text = Temp1.Last

                Temp2.Enqueue(CDbl(s(52).Replace(".", ",")))
                Temp2TB.Text = Temp2.Last

                Temp3.Enqueue(CDbl(s(54).Replace(".", ",")))
                Temp3TB.Text = Temp3.Last

                Temp4.Enqueue(CDbl(s(56).Replace(".", ",")))
                Temp4TB.Text = Temp4.Last

            '    Light.Enqueue(CDbl(s(58).Replace(".", ",")))
            '    'Light1min.Enqueue(CInt(s(58)))
            '    LightTB.Text = Light.Last
            'Else
            '    Light.Enqueue(CDbl(s(58).Replace(".", ",")))
            '    'Light1min.Enqueue(CInt(s(48)))
            '    LightTB.Text = Light.Last
            'End If


            'qxwb.Enqueue(Format(TimeOfDay, "HH:mm:ss"))

            'If COstat.Count > 1 And COstat.Last <> nold Then

            '    strw1 = Now & " " & CInt(lq.Average) & " " & nold

            '    'sw1.WriteLine(strw1)
            '    'sw1.Flush()

            '    'lqav.Enqueue(CInt(lq.Average))
            '    lq.Clear()



            '    'lqx.Enqueue(Format(TimeOfDay, "HH:mm:ss"))
            '    'lqin.Enqueue(nold)

            'Else
            '    lq.Enqueue(Light.Last)

            'End If

            'nold = COstat.Last

            'Chart1.Series(0).Points.DataBindXY(qx, Tempint)
            'Chart1.Series(1).Points.DataBindXY(qx, Tempext)
            'Chart1.Series(2).Points.DataBindXY(qx, Humi)

            'Chart3.Series(0).Points.DataBindXY(qxwb, Hgmm)

            'Chart2.Series(1).Points.DataBindXY(qx, CO2)

            'Chart24.Series(0).Points.DataBindXY(qx, Light)

            'Chart3.Series("COmeas").Points.DataBindXY(qx, COmeas)

            'If COstat.Last = "N" Then
            '    Chart3.Series("COcurr").Color = Color.Red
            'Else
            '    Chart3.Series("COcurr").Color = Color.Green
            'End If
            'Chart3.Series("COcurr").Points.DataBindXY(qx, COcurr)

            'Chart3.Series("Smog").Points.DataBindXY(qx, Smog)
            'Chart3.Series("H2").Points.DataBindXY(qx, H2)
            'Chart3.Series("LPG").Points.DataBindXY(qx, LPG)
            'Chart3.Series("NH3").Points.DataBindXY(qx, NH3)
            'Chart3.Series("Spirt").Points.DataBindXY(qx, Spirt)
            'Chart4.Series("NAGR\IZM").Points.DataBindXY(qx, COstat)

            'Chart23.Series("COmeas").Points.DataBindXY(qx, COmeas)
            'Chart23.Series("COcurr").Points.DataBindXY(qx, COcurr)
            'Chart23.Series("Smog").Points.DataBindXY(qx, Smog)
            'Chart23.Series("H2").Points.DataBindXY(qx, H2)
            'Chart23.Series("LPG").Points.DataBindXY(qx, LPG)
            'Chart23.Series("NH3").Points.DataBindXY(qx, NH3)
            'Chart23.Series("Spirt").Points.DataBindXY(qx, Spirt)

            'If Temp0.Count > 0 Then
            '    Chart8.Series("Temp0").Points.DataBindXY(qx, Temp0)
            '    Chart8.Series("Temp1").Points.DataBindXY(qx, Temp1)
            '    Chart8.Series("Temp2").Points.DataBindXY(qx, Temp2)
            '    Chart8.Series("Temp3").Points.DataBindXY(qx, Temp3)
            '    Chart8.Series("Temp4").Points.DataBindXY(qx, Temp4)
            'End If


            'Chart10.Series(0).Points.DataBindXY(qx, MagX)
            'Chart13.Series(0).Points.DataBindXY(qx, MagY)
            'Chart14.Series(0).Points.DataBindXY(qx, MagZ)
            'Chart15.Series(0).Points.DataBindXY(qx, MagM)
            ''17 16 12 11
            'Chart17.Series(0).Points.DataBindXY(qx, GravX)
            'Chart16.Series(0).Points.DataBindXY(qx, GravY)
            'Chart12.Series(0).Points.DataBindXY(qx, GravZ)
            'Chart11.Series(0).Points.DataBindXY(qx, GravM)
            ''21 20 19 18
            'Chart21.Series(0).Points.DataBindXY(qx, GyroX)
            'Chart20.Series(0).Points.DataBindXY(qx, GyroY)
            'Chart19.Series(0).Points.DataBindXY(qx, GyroZ)
            'Chart18.Series(0).Points.DataBindXY(qx, GyroM)


            If Hgmm.Count > 40 Then
                'qxwb.Dequeue()
                'Tempint.Dequeue()
                Hgmm.Dequeue()
                CO2.Dequeue()
                'Tempext.Dequeue()
                'Humi.Dequeue()
                'COstat.Dequeue()

                'COmeas.Dequeue()
                'COcurr.Dequeue()
                'Smog.Dequeue()
                'H2.Dequeue()
                'LPG.Dequeue()
                'NH3.Dequeue()
                'Spirt.Dequeue()

                'MagX.Dequeue()
                'MagY.Dequeue()
                'MagZ.Dequeue()
                'MagM.Dequeue()

                'GravX.Dequeue()
                'GravY.Dequeue()
                'GravZ.Dequeue()
                'GravM.Dequeue()

                'GyroX.Dequeue()
                'GyroY.Dequeue()
                'GyroZ.Dequeue()
                'GyroM.Dequeue()

                'Light.Dequeue()

                Temp0.Dequeue()
                Temp1.Dequeue()
                Temp2.Dequeue()
                Temp3.Dequeue()
                Temp4.Dequeue()


            End If

            'If Tempint.Count > 30 Then '30*4s = 2 min
            '    Light1min.Dequeue()

            'End If

        Else
            'TempIntTB.BackColor = Color.LightPink
            'swerror.WriteLine(Now & " " & data)
            'swerror.Flush()
        End If

        'msg(data)


    End Sub

    Private Sub OnRead_Invoke(ByVal data As String) ' Процедура получения данных с сервера с синхронизацией потоков
        Try
            If Me.IsHandleCreated Then
                Me.Invoke(New Class_client.StatusInvoker(AddressOf Me.OnRead), data)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub OnRead_Invokewb(ByVal data As String) ' Процедура получения данных с сервера с синхронизацией потоков
        Try
            If Me.IsHandleCreated Then
                Me.Invoke(New Class_client.StatusInvoker(AddressOf Me.OnReadwb), data)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub Disconnect(ByVal reson As String)
        ' Процедура при отключении клиента
        msg(vbCrLf & "Подключение разорвано. Причина: " & reson & vbCrLf)

    End Sub

    Private Sub Disconnectwb(ByVal reson As String)
        ' Процедура при отключении клиента
        msg(vbCrLf & "WBПодключение разорвано. Причина: " & reson & vbCrLf)

    End Sub
    Private Sub Disconnect_Invoke(ByVal t As String) ' Процедура отключения от сервера с синхронизацией потоков
        Try
            If Me.IsHandleCreated Then
                Me.Invoke(New Class_client.StatusInvoker(AddressOf Me.Disconnect), t)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Disconnect_Invokewb(ByVal t As String) ' Процедура отключения от сервера с синхронизацией потоков
        Try
            If Me.IsHandleCreated Then
                Me.Invoke(New Class_client.StatusInvoker(AddressOf Me.Disconnectwb), t)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub msg(ByVal t As String) ' Процедура вывода сообщений в консоль
        Dim s() As String
        Dim o(10) As String

        If t = "Выпали по таймауту" & vbCrLf Then
            s2() 'отключим поток и почистим хвосты - уничтожим тред, отключим обработку событий 
            s1() 'подключимся заново
        End If

        If t = "WBВыпали по таймауту" & vbCrLf Then
            s2wb() 'отключим поток и почистим хвосты - уничтожим тред, отключим обработку событий 
            s1wb() 'подключимся заново
        End If

        s = TextBox1.Lines

        If s.GetUpperBound(0) > 10 Then
            'TextBox1.Lines(0) = ""

            For j = 0 To 9
                o(j) = s(j + 1)
            Next

            TextBox1.Lines = o
        End If

        TextBox1.AppendText(Now & " >> " + t)

    End Sub

    Private Sub msg_Invoke(ByVal t As String) ' Процедура выполняемая для вывода сообщения в консоль с синхронизацией потоков
        Try
            If Me.IsHandleCreated Then
                Me.Invoke(New Class_client.StatusInvoker(AddressOf Me.msg), t)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim strw As String

        If mkr.Count = 0 Then
            Exit Sub
        End If

        'Dim mkrsr As Integer = mkr.Average
        'Dim mkrsvo As Double = mkr.Min

        strw = Now & " Отсчетов: " & mkr.Count & " Среднее: " & Math.Round(mkr.Average, 0) & " Мин: " & mkr.Min & " Мах: " & mkr.Max & " СКО: " & TextBox5.Text & " Температура: " & TextBox8.Text & " Влажность: " & TextBox3.Text & " Высокое: " & TextBox10.Text

        sw.WriteLine(strw)
        sw.Flush()

        If CheckBox1.Checked Then
            Dim strText
            ', objVoice
            'objVoice = CreateObject("SPEECH.SpVoice")
            'SetAttr(objVoice = CreateObject("SPEECH.SpVoice"))
            'If WScript.Arguments.Count = 1 Then SetAttr(objVoice.voice = objVoice.GetVoices.Item(CInt(WScript.Arguments(0))))
            'SetAttr(objHTML = CreateObject("htmlfile"))
            'strText = objHTML.ParentWindow.ClipboardData.GetData("text")

            Dim temp As Integer

            Dim Dig_to_say As String
            Dim body As String

            Dig_to_say = TextBox8.Text
            temp = Math.Round(Convert.ToSingle(Dig_to_say), 0)

            Dig_to_say = temp.ToString

            body = "градус" & suffix(Dig_to_say, temp)

            strText = "Температура за бортом " & Dig_to_say & " " & body & " Цельсия."


            Dig_to_say = TextBox3.Text
            temp = Math.Round(Convert.ToSingle(Dig_to_say), 0)
            Dig_to_say = temp.ToString

            body = "процент" & suffix(Dig_to_say, temp)

            strText = strText & "    Влажность за бортом: " & Dig_to_say & " " & body & "."

            Dig_to_say = TextBox11.Text
            temp = Math.Round(Convert.ToSingle(Dig_to_say), 0)
            Dig_to_say = temp.ToString

            body = "миллиметр" & suffix(Dig_to_say, temp)

            strText = strText & " Атмосферное давление " & Dig_to_say & " " & body & " ртутного столба."

            If strText <> "" Then
                Dim speaker As New SpeechSynthesizer()
                speaker.Rate = 1
                speaker.Volume = 100
                speaker.Speak(strText)
                'objVoice.speak(strText)
                'objVoice.speak(strText)
            End If
        End If

        'Dim status As StreamWriter = New StreamWriter(path3, False) ' ReWrite file
        Dim j As Integer

        Using fs As New FileStream(path3, FileMode.Create)
            Using w As New BinaryWriter(fs)
                w.Write(qx.Count - 1)
                For j = 0 To qx.Count - 1
                    w.Write(qx(j))
                    w.Write(mkr(j))
                    w.Write(TempC(j))
                    w.Write(Humi(j))
                    w.Write(HiVolt(j))
                    w.Write(Hgmmr(j))
                    w.Write(CO2r(j))
                    w.Write(Temp0r(j))
                    w.Write(Temp1r(j))
                    w.Write(Temp2r(j))
                    w.Write(Temp3r(j))
                    w.Write(Temp4r(j))
                Next
            End Using
        End Using


        'status.Close()
        'status.Dispose()

    End Sub
    Function suffix(txtparam As String, temp As Integer) As String
        Dim LastDigit As String

        LastDigit = Trim(txtparam).Substring(txtparam.Length - 1, 1)

        If Math.Abs(temp) > 4 And Math.Abs(temp) < 21 Then
            suffix = "ов"
        Else
            Select Case LastDigit
                Case "1"
                    suffix = ""
                Case "2", "3", "4"
                    suffix = "а"
                Case Else
                    suffix = "ов"
            End Select
        End If

    End Function

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Location = New Size(57, 134)
        Me.Size = New Size(962, 348)

        My.Settings.w_size = Me.Size
        My.Settings.w_loc = Me.Location

        My.Settings.Save()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        msg("Height=" & Me.Size.Height & vbCrLf)
        msg("Width=" & Me.Size.Width & vbCrLf)
        msg("X=" & Me.Location.X & vbCrLf)
        msg("Y=" & Me.Location.Y & vbCrLf)
    End Sub

    Private Sub TextBox9_MouseHover(sender As Object, e As EventArgs) Handles TextBox9.MouseHover
        TextBox9.BringToFront()
    End Sub

    Private Sub TextBox9_MouseLeave(sender As Object, e As EventArgs) Handles TextBox9.MouseLeave
        TextBox9.SendToBack()
    End Sub
End Class


Public Class Class_client
    '
    'Public Shared addr As String = "192.168.1.155"
    'Public Shared port As Integer = 8080

    Public Delegate Sub StatusInvoker(ByVal t As String) ' делегат для синхронизации при выводе сообщений

    Dim client As TcpClient ' Текущий TcpClient связывающий клиента и сервер
    Dim stream As NetworkStream ' Поток данных связывающий клиента и сервер
    Dim flag As Byte ' Пока не используется. Нужен для индификации сервером "правильного" клиента

    Dim connection As Boolean = False ' Состояние подключения


    Public Event OnRead(ByVal _data As String) ' Событие, при получении данных от сервера
    Public Event Disconnected(ByVal _reson As String)  ' Событие, при отключении от сервера
    Public Event messege(ByVal text As String)  ' Событие, при необходимости передать сообщение из класса


    Dim clientwb As TcpClient ' Текущий TcpClient связывающий клиента и сервер
    Dim streamwb As NetworkStream ' Поток данных связывающий клиента и сервер
    Dim flagwb As Byte ' Пока не используется. Нужен для индификации сервером "правильного" клиента

    Dim connectionwb As Boolean = False ' Состояние подключения


    Public Event OnReadwb(ByVal _data As String) ' Событие, при получении данных от сервера
    Public Event Disconnectedwb(ByVal _reson As String)  ' Событие, при отключении от сервера

    'Dim s As Object
    Public client_Thread As Threading.Thread
    Public client_Threadwb As Threading.Thread

    Public Function Connect(ByVal server As String, ByVal port As Integer, ByVal _Flag As Byte) As Boolean ' Функция подключения клиента к серверу. Вернет True в случае удачи и False при ошибке.
        If connection = False Then ' Проверка не подключены ли мы уже?
            Try
                Me.client = New TcpClient(server, port)
                Me.stream = client.GetStream()
                Me.stream.ReadTimeout = 180000 'ждать данных 3мин интервал между посылками нормально 2 мин
                Me.flag = _Flag
                client_Thread = New Threading.Thread(AddressOf Me.doListen)

                client_Thread.IsBackground = True 'бекграунд поток закроется при закрытии формы.

                client_Thread.Start() ' Запуск отдельного потока на прослушивание данных с сервера.
                connection = True ' Состояние=Подключен
                Return True

                Return True
            Catch ex As Exception
                msg(ex.Message) ' Сообщаем об ошибке
                Return False
            End Try
        Else
            Return False ' Если уже подключены, тогда новое подключение не удалось
        End If

    End Function


    Public Function Connectwb(ByVal server As String, ByVal port As Integer, ByVal _Flag As Byte) As Boolean ' Функция подключения клиента к серверу. Вернет True в случае удачи и False при ошибке.
        If connectionwb = False Then ' Проверка не подключены ли мы уже?
            Try
                Me.clientwb = New TcpClient(server, port)
                Me.streamwb = client.GetStream()
                Me.streamwb.ReadTimeout = 10000 'ждать данных 3мин интервал между посылками нормально 2 мин
                Me.flagwb = _Flag
                client_Threadwb = New Threading.Thread(AddressOf Me.doListenwb)

                client_Threadwb.IsBackground = True 'бекграунд поток закроется при закрытии формы.

                client_Threadwb.Start() ' Запуск отдельного потока на прослушивание данных с сервера.
                connectionwb = True ' Состояние=Подключен
                Return True

                Return True
            Catch ex As Exception
                msg(ex.Message) ' Сообщаем об ошибке
                Return False
            End Try
        Else
            Return False ' Если уже подключены, тогда новое подключение не удалось
        End If

    End Function


    Public Sub Disconnect(Optional ByVal s As String = "") ' Процедура отключения от сервера. Возможно указать коментарий.
        If connection Then
            connection = False ' Состояние=Отключен
            stream.Close()
            client.Close()
            client_Thread.Abort()

            If s = "" Then
                RaiseEvent Disconnected("Отключен") ' Вызываем событие отключения с коментарием "по умолчанию"
            Else
                RaiseEvent Disconnected(s) ' Вызываем событие отключения с коментарием указанным при вызове текущей процедуры
            End If

        End If
    End Sub


    Public Sub Disconnectwb(Optional ByVal s As String = "") ' Процедура отключения от сервера. Возможно указать коментарий.
        If connectionwb Then
            connectionwb = False ' Состояние=Отключен
            streamwb.Close()
            clientwb.Close()
            client_Threadwb.Abort()

            If s = "" Then
                RaiseEvent Disconnected("Отключен") ' Вызываем событие отключения с коментарием "по умолчанию"
            Else
                RaiseEvent Disconnected(s) ' Вызываем событие отключения с коментарием указанным при вызове текущей процедуры
            End If

        End If
    End Sub


    Public Function Send(ByVal data As String) As Boolean ' Функция отправки сообщения серверу. Вернет True, если посылка удалась или False в обратном случаее
        If connection Then
            Try
                Dim databyte() As Byte = System.Text.Encoding.UTF8.GetBytes(data) ' Создаем массив байтов для отправки из Data
                stream.Write(databyte, 0, databyte.Length) ' Отправляем наш массив байт.
                stream.Flush() ' честно не знаю, что это делает. Работает и с этим и без, но в большинстве примеров есть, поэтому оставил.
                Return True ' Говорим, что посылка удалась
            Catch ex As Exception
                msg(ex.Message) ' Сообщаем об ошибке
                Disconnect("Проблема при отправке.") ' Отключаемся
                Return False
            End Try
        Else
            Return False ' Если не подключены, то отправка не удалась.
        End If
    End Function


    Private Sub doListen() ' Процедура прослушки сообщений от сервера


        Dim bytes(1024) As Byte ' 1мб буфер, чтобы данные не дробились.
        Dim data As String
        Dim i As Int32

        Dim sb As New StringBuilder()

        Try
            i = stream.Read(bytes, 0, bytes.Length) ' Ждем новых данных от сервера. При получении заносим данные в буфер bytes и длину данных в i
            While (i <> 0) ' Цикл бесконечный, пока длина считанных данных не окажется=0, что ознает, что соединение с сервером прекращено.
                data = System.Text.Encoding.UTF8.GetString(bytes, 0, i) ' Переводим байты в текст

                'If Mid(data, 1, 5) = "Time:" Then
                '    sb.Clear()
                '    sb.Append(data)
                'Else
                sb.Append(data)
                'End If

                If Right(data, 2) = vbCrLf Then

                    Send("+")
                    RaiseEvent OnRead(sb.ToString) ' Вызываем событие о получении новых данных
                    sb.Clear()
                End If

                i = stream.Read(bytes, 0, bytes.Length) ' Ждем новых данных от сервера. При получении заносим данные в буфер bytes и длину данных в i
            End While
            Disconnect() ' Так как вышли из цикла, то сервер закрыл соединение. Следовательно отключаемся.
        Catch ex As Exception

            msg("Выпали по таймауту" & vbCrLf)
            Disconnect(ex.Message) ' Ошибка, значит отключаемся, указав причину.

            'Dim p As Long = 0
            'Do

            '    msg(vbCrLf & "Попытка " & p & vbCrLf)
            '    p = p + 1

            'Loop Until (Connect(addr, port, AscW("+")))
        End Try

    End Sub

    Private Sub doListenwb() ' Процедура прослушки сообщений от сервера


        Dim bytes(1024) As Byte ' 1мб буфер, чтобы данные не дробились.
        Dim data As String
        Dim i As Int32

        Dim sb As New StringBuilder()

        Try
            i = streamwb.Read(bytes, 0, bytes.Length) ' Ждем новых данных от сервера. При получении заносим данные в буфер bytes и длину данных в i
            While (i <> 0) ' Цикл бесконечный, пока длина считанных данных не окажется=0, что ознает, что соединение с сервером прекращено.
                data = System.Text.Encoding.UTF8.GetString(bytes, 0, i) ' Переводим байты в текст

                'If Mid(data, 1, 5) = "Time:" Then
                '    sb.Clear()
                '    sb.Append(data)
                'Else
                sb.Append(data)
                'End If

                If Right(data, 2) = vbCrLf Then

                    Send("+")
                    RaiseEvent OnReadwb(sb.ToString) ' Вызываем событие о получении новых данных
                    sb.Clear()
                End If

                i = streamwb.Read(bytes, 0, bytes.Length) ' Ждем новых данных от сервера. При получении заносим данные в буфер bytes и длину данных в i
            End While
            Disconnectwb() ' Так как вышли из цикла, то сервер закрыл соединение. Следовательно отключаемся.
        Catch ex As Exception

            msg("WBВыпали по таймауту" & vbCrLf)
            Disconnect(ex.Message) ' Ошибка, значит отключаемся, указав причину.

            'Dim p As Long = 0
            'Do

            '    msg(vbCrLf & "Попытка " & p & vbCrLf)
            '    p = p + 1

            'Loop Until (Connect(addr, port, AscW("+")))
        End Try

    End Sub


    Public Sub msg(ByVal t As String) ' Процедура отправки сообщения
        Try
            RaiseEvent messege(t)  ' Вызываем событие о новом сообщении
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
