Imports DMXIT.USBDMXPro
Imports System.Timers
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.IO
Imports System.Configuration
Imports AxWMPLib
Imports MongoDB.Bson
Imports MongoDB.Driver
Imports System.Drawing.Text
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Public Class frmMain
    Private posArray As New ArrayList
    Private posArrayCount As Integer = 0
    Private MainClass As New USBDMXPro
    Private cls As New dataclass
    Private device As Integer
    Private Shared timer1Counter As Integer = 1
    Private Shared exitFlag As Boolean = False
    Public Shared dmxdata(0 To 511) As Byte
    Private dmx1Value As Boolean = False
    Private dmx2Value As Boolean = False
    Private dmx3Value As Boolean = False
    Private dmx4Value As Boolean = False
    Private dmx5Value As Boolean = False
    Private dmx6Value As Boolean = False
    Private dmx7Value As Boolean = False
    Private dmx8Value As Boolean = False
    Private dmx9Value As Boolean = False
    Private dmx10Value As Boolean = False
    Private dmx11Value As Boolean = False
    Private dmx12Value As Boolean = False
    Private dmx13Value As Boolean = False
    Private dmx14Value As Boolean = False
    Private dmx15Value As Boolean = False
    Private dmx16Value As Boolean = False
    Private dmx17Value As Boolean = False
    Private dmx18Value As Boolean = False
    Private dmx19Value As Boolean = False
    Private dmx20Value As Boolean = False
    Private dmx21Value As Boolean = False
    Private dmx22Value As Boolean = False
    Private dmx23Value As Boolean = False
    Private dmx24Value As Boolean = False
    Private dmx25Value As Boolean = False
    Private dmx26Value As Boolean = False
    Private dmx27Value As Boolean = False
    Private dmx28Value As Boolean = False
    Private dmx29Value As Boolean = False
    Private dmx30Value As Boolean = False
    Private dmx31Value As Boolean = False
    Private dmx32Value As Boolean = False
    Private dmx33Value As Boolean = False
    Private dmx34Value As Boolean = False
    Private dmx35Value As Boolean = False
    Private dmx36Value As Boolean = False
    Private dmx37Value As Boolean = False
    Private dmx38Value As Boolean = False
    Private dmx39Value As Boolean = False
    Private dmx40Value As Boolean = False
    Private dmx41Value As Boolean = False
    Private dmx42Value As Boolean = False
    Private dmx43Value As Boolean = False
    Private dmx44Value As Boolean = False
    Private dmx45Value As Boolean = False
    Private dmx46Value As Boolean = False
    Private dmx47Value As Boolean = False
    Private dmx48Value As Boolean = False

    Private Sub goRed(sender As Object, e As EventArgs) Handles btnRed.Click
        dmxdata(1) = 255
        dmxdata(2) = 0
        dmxdata(3) = 0
        MainClass.sendDMXdata(dmxdata, 4)
    End Sub

    Private Sub FormLoad(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim status As Integer


        status = MainClass.startDMX()
        If status = 0 Then
            Label2.Text = "Connected to DMX device"
            Label2.ForeColor = Color.Green
        Else
            Label2.Text = "DMX device error"
            Label2.ForeColor = Color.Red
        End If

        ' connect to mongo db
        If cls.InitiateDb() <> 0 Then
            Label6.Text = "Database error"
            Label6.ForeColor = Color.Red
        Else
            Label6.Text = ""
            Label6.ForeColor = Color.Green
        End If

        ' load default application settings
        txtLayoutName.Text = My.Settings.defaultLayout

        'Dim _fixture As New datamodel.DmxFixture
        '_fixture.name = "test name"
        '_fixture.manufacturer = "test manufacturer"
        '_fixture.model = "test model"

        'Collection.InsertOne(_fixture)

        Timer2.Interval = 1000

    End Sub

    Private Sub goGreen(sender As Object, e As EventArgs) Handles btnGreen.Click
        Timer2.Stop()
        dmxdata(1) = 0
        dmxdata(2) = 255
        dmxdata(3) = 0
        MainClass.sendDMXdata(dmxdata, 4)
    End Sub

    Private Sub goBlue(sender As Object, e As EventArgs) Handles btnBlue.Click
        Timer2.Stop()
        dmxdata(1) = 0
        dmxdata(2) = 0
        dmxdata(3) = 255
        MainClass.sendDMXdata(dmxdata)
    End Sub

    Private Sub goScene1(sender As Object, e As EventArgs) Handles btnScene1.Click
        Timer2.Stop()
        timer1Counter = 1
        Timer1.Interval = 100
        Timer1.Start()
    End Sub

    Private Sub Timer1_Process(sender As Object, e As EventArgs) Handles Timer1.Tick

        Select Case timer1Counter
            Case 1, 3, 5, 7
                MainClass.dmxdata(1) = 255
                MainClass.dmxdata(2) = 0
                MainClass.dmxdata(3) = 0
                MainClass.sendDMXpro()
            Case 2, 4, 6, 8
                MainClass.dmxdata(1) = 0
                MainClass.dmxdata(2) = 0
                MainClass.dmxdata(3) = 255
                MainClass.sendDMXpro()
            Case 9
                Timer1.Stop()
        End Select

        timer1Counter += 1

    End Sub

    Private Sub Timer2_Process(sender As Object, e As EventArgs) Handles Timer2.Tick
        test()

    End Sub

    Private Sub test()
        'Dim screenshot As Bitmap = TakeScreenShot()
        Dim clr As Color = GetWPFImageAverageRGB()

        dmxdata(1) = CByte(clr.R)
        dmxdata(2) = CByte(clr.G)
        dmxdata(3) = CByte(clr.B)
        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub goScreenRGB(sender As Object, e As EventArgs) Handles MyBase.Click, btnScreenRGB.Click
        Timer2.Start()
        test()
    End Sub

    Private Sub AvgColors(ByVal InBitmap As Bitmap, ByRef avgRed As Byte, ByRef avgGreen As Byte, ByRef avgBlue As Byte)

        Dim btPixels(InBitmap.Height * InBitmap.Width * 3 - 1) As Byte
        Dim hPixels As GCHandle = GCHandle.Alloc(btPixels, GCHandleType.Pinned)
        Dim bmp24Bpp As New Bitmap(InBitmap.Width, InBitmap.Height, InBitmap.Width * 3,
          Imaging.PixelFormat.Format24bppRgb, hPixels.AddrOfPinnedObject)

        Using gr As Graphics = Graphics.FromImage(bmp24Bpp)
            gr.DrawImageUnscaledAndClipped(InBitmap, New Rectangle(0, 0,
              bmp24Bpp.Width, bmp24Bpp.Height))
        End Using

        Dim sumRed As Long
        Dim sumGreen As Long
        Dim sumBlue As Long

        For i = 0 To btPixels.Length - 1 Step 3
            sumBlue += btPixels(i)
            sumGreen += btPixels(i + 1)
            sumRed += btPixels(i + 2)
            If i = 3000 Then
                'Debug.Print("test")
            End If
        Next
        hPixels.Free()

        avgRed = CByte(sumRed / (btPixels.Length))
        avgGreen = CByte(sumGreen / (btPixels.Length))
        avgBlue = CByte(sumBlue / (btPixels.Length))
    End Sub

    Public Function GetWPFImageAverageRGB() As System.Drawing.Color
        ' define inner reactangle selection area
        Dim rect As New Rectangle(0, 288, 1440, 864)

        ' sets size to FULL screen
        'Using b As New Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height)

        ' sets size to specific rectangle parameters
        Using b As New Bitmap(rect.Width, rect.Height)
            Using g = Graphics.FromImage(b)
                ' code copies FULL screen
                'g.CopyFromScreen(0, 0, 0, 0, b.Size)

                ' code copies rectangle selection only
                g.CopyFromScreen(rect.Location, New Point(0, 0), rect.Size)
            End Using

            Dim OutColor As Color
            Dim st As Date
            Dim TotalTime As Double


            ''Interpolation Method
            'Using bOut As New Bitmap(1, 1) '<<< - i keep the bout object outside of the loop ... because if you want speed you would not re-create the bitmap object all the time
            '    Using g = Graphics.FromImage(bOut) '<<< - same with this ""
            '        st = Now
            '        For i = 1 To 100
            '            g.InterpolationMode = Drawing2D.InterpolationMode.High
            '            g.DrawImage(b, New Rectangle(0, 0, 1, 1))
            '            OutColor = b.GetPixel(0, 0)
            '        Next
            '        TotalTime = Now.Subtract(st).TotalMilliseconds
            '        Debug.Print("Interpolation Method: " & OutColor.ToString & " " & TotalTime & "ms")
            '    End Using
            'End Using

            'LockBits Method
            st = Now
            For i = 1 To 100
                Dim bmpData As System.Drawing.Imaging.BitmapData = b.LockBits(New Rectangle(0, 0, b.Width, b.Height), Imaging.ImageLockMode.ReadOnly, Imaging.PixelFormat.Format32bppArgb)
                Dim bitmapAddress As IntPtr = bmpData.Scan0

                Dim iCount As Integer
                Dim Max = (b.Width * b.Height * 4) - 4 '<<< - Edit as faster than: CInt(bmpData.Stride * bmpData.Height) - 4
                Dim TotalR = 0&
                Dim TotalG = 0&
                Dim TotalB = 0&

                For iPixel = 0 To Max Step 4
                    iCount += 1
                    TotalR += System.Runtime.InteropServices.Marshal.ReadByte(bitmapAddress, iPixel + 2)
                    TotalG += System.Runtime.InteropServices.Marshal.ReadByte(bitmapAddress, iPixel + 1)
                    TotalB += System.Runtime.InteropServices.Marshal.ReadByte(bitmapAddress, iPixel)
                Next

                Dim TotalPixels = (Max / 4) + 1
                Return Color.FromArgb(TotalR \ TotalPixels, TotalG \ TotalPixels, TotalB \ TotalPixels)
                b.UnlockBits(bmpData)
            Next

        End Using
    End Function

    Private ReadOnly Property FromStream(ms As MemoryStream) As Object
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Private Function TakeScreenShot() As Bitmap
        Dim tmpImg As New Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height)
        Using g As Graphics = Graphics.FromImage(tmpImg)
            g.CopyFromScreen(Screen.PrimaryScreen.Bounds.Location, New Point(0, 0), Screen.PrimaryScreen.Bounds.Size)
        End Using
        Return tmpImg
    End Function

    Private Sub GetRGB(sender As Object, e As EventArgs) Handles btnGetRGB.Click
        Dim clr As Color = GetWPFImageAverageRGB()

        dmxdata(1) = CByte(clr.R)
        dmxdata(2) = CByte(clr.G)
        dmxdata(3) = CByte(clr.B)

        MainClass.sendDMXdata(dmxdata, 0)
        WriteToFile(CStr(clr.R), CStr(clr.G), CStr(clr.B))

    End Sub

    Private Sub WriteToFile(sRed, sGreen, sBlue)
        Dim filePath As String = String.Format("C:\ErrorLog_{0}.txt", DateTime.Today.ToString("dd-MMM-yyyy"))

        Dim fileExists As Boolean = File.Exists(filePath)

        'Using writer As New StreamWriter(filePath, True)
        'writer.WriteLine(sRed & "," & sGreen & "," & sBlue)
        ' End Using
    End Sub

    Private Async Sub btnLoad_Click(sender As Object, e As EventArgs) Handles btnLoad.Click

        Select Case cmboPreset.Text
            Case "New"
                If OpenFileDialog1.ShowDialog = DialogResult.OK Then
                    AxWindowsMediaPlayer1.URL = OpenFileDialog1.FileName

                End If
            Case "Greatest Show Audio"
                AxWindowsMediaPlayer1.URL = "c:\democontent\greatestshowdemo.mp3"
                txtContentID.Text = "0"

            Case "Harry Potter Video"
                AxWindowsMediaPlayer1.URL = "c:\democontent\harrypotter.mp4"
                txtContentID.Text = "1"
        End Select

        AxWindowsMediaPlayer1.settings.setMode("loop", False)
        AxWindowsMediaPlayer1.Ctlcontrols.stop()

        Await populateDG1(txtContentID.Text)


    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs)
        AxWindowsMediaPlayer1.Ctlcontrols.play()
    End Sub

    Private Sub btnStop_Click(sender As Object, e As EventArgs)
        AxWindowsMediaPlayer1.Ctlcontrols.stop()
    End Sub

    Private Sub btnPause_Click(sender As Object, e As EventArgs)
        AxWindowsMediaPlayer1.Ctlcontrols.pause()
    End Sub

    Private Sub btnTest_Click(sender As Object, e As EventArgs) Handles btnTest.Click
        dmxdata(47) = 255
        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click
        If dmx1Value = False Then
            btn1.BackColor = Color.RoyalBlue
            dmx1Value = True
            dmxdata(1) = 255
        ElseIf dmx1Value = True Then
            btn1.BackColor = Color.Transparent
            dmx1Value = False
            dmxdata(1) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)

    End Sub

    Private Sub btn2_Click(sender As Object, e As EventArgs) Handles btn2.Click
        If dmx2Value = False Then
            btn2.BackColor = Color.RoyalBlue
            dmx2Value = True
            dmxdata(2) = 255
        ElseIf dmx2Value = True Then
            btn2.BackColor = Color.Transparent
            dmx2Value = False
            dmxdata(2) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn3_Click(sender As Object, e As EventArgs) Handles btn3.Click
        If dmx3Value = False Then
            btn3.BackColor = Color.RoyalBlue
            dmx3Value = True
            dmxdata(3) = 255
        ElseIf dmx3Value = True Then
            btn3.BackColor = Color.Transparent
            dmx3Value = False
            dmxdata(3) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn4_Click(sender As Object, e As EventArgs) Handles btn4.Click
        If dmx4Value = False Then
            btn4.BackColor = Color.RoyalBlue
            dmx4Value = True
            dmxdata(4) = 255
        ElseIf dmx4Value = True Then
            btn4.BackColor = Color.Transparent
            dmx4Value = False
            dmxdata(4) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn5_Click(sender As Object, e As EventArgs) Handles btn5.Click
        If dmx5Value = False Then
            btn5.BackColor = Color.RoyalBlue
            dmx5Value = True
            dmxdata(5) = 255
        ElseIf dmx5Value = True Then
            btn5.BackColor = Color.Transparent
            dmx5Value = False
            dmxdata(5) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn6_Click(sender As Object, e As EventArgs) Handles btn6.Click
        If dmx6Value = False Then
            btn6.BackColor = Color.RoyalBlue
            dmx6Value = True
            dmxdata(6) = 255
        ElseIf dmx6Value = True Then
            btn6.BackColor = Color.Transparent
            dmx6Value = False
            dmxdata(6) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn7_Click(sender As Object, e As EventArgs) Handles btn7.Click
        If dmx7Value = False Then
            btn7.BackColor = Color.RoyalBlue
            dmx7Value = True
            dmxdata(7) = 255
        ElseIf dmx7Value = True Then
            btn7.BackColor = Color.Transparent
            dmx7Value = False
            dmxdata(7) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn8_Click(sender As Object, e As EventArgs) Handles btn8.Click
        If dmx8Value = False Then
            btn8.BackColor = Color.RoyalBlue
            dmx8Value = True
            dmxdata(8) = 255
        ElseIf dmx8Value = True Then
            btn8.BackColor = Color.Transparent
            dmx8Value = False
            dmxdata(8) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn9_Click(sender As Object, e As EventArgs) Handles btn9.Click
        If dmx9Value = False Then
            btn9.BackColor = Color.RoyalBlue
            dmx9Value = True
            dmxdata(9) = 255
        ElseIf dmx9Value = True Then
            btn9.BackColor = Color.Transparent
            dmx9Value = False
            dmxdata(9) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn10_Click(sender As Object, e As EventArgs) Handles btn10.Click
        If dmx10Value = False Then
            btn10.BackColor = Color.RoyalBlue
            dmx10Value = True
            dmxdata(10) = 255
        ElseIf dmx10Value = True Then
            btn10.BackColor = Color.Transparent
            dmx10Value = False
            dmxdata(10) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn11_Click(sender As Object, e As EventArgs) Handles btn11.Click
        If dmx11Value = False Then
            btn11.BackColor = Color.RoyalBlue
            dmx11Value = True
            dmxdata(11) = 255
        ElseIf dmx11Value = True Then
            btn11.BackColor = Color.Transparent
            dmx11Value = False
            dmxdata(11) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn12_Click(sender As Object, e As EventArgs) Handles btn12.Click
        If dmx12Value = False Then
            btn12.BackColor = Color.RoyalBlue
            dmx12Value = True
            dmxdata(12) = 255
        ElseIf dmx12Value = True Then
            btn12.BackColor = Color.Transparent
            dmx12Value = False
            dmxdata(12) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn13_Click(sender As Object, e As EventArgs) Handles btn13.Click
        If dmx13Value = False Then
            btn13.BackColor = Color.RoyalBlue
            dmx13Value = True
            dmxdata(13) = 255
        ElseIf dmx13Value = True Then
            btn13.BackColor = Color.Transparent
            dmx13Value = False
            dmxdata(13) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn14_Click(sender As Object, e As EventArgs) Handles btn14.Click
        If dmx14Value = False Then
            btn14.BackColor = Color.RoyalBlue
            dmx14Value = True
            dmxdata(14) = 255
        ElseIf dmx14Value = True Then
            btn14.BackColor = Color.Transparent
            dmx14Value = False
            dmxdata(14) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn15_Click(sender As Object, e As EventArgs) Handles btn15.Click
        If dmx15Value = False Then
            btn15.BackColor = Color.RoyalBlue
            dmx15Value = True
            dmxdata(15) = 255
        ElseIf dmx15Value = True Then
            btn15.BackColor = Color.Transparent
            dmx15Value = False
            dmxdata(15) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn16_Click(sender As Object, e As EventArgs) Handles btn16.Click
        If dmx16Value = False Then
            btn16.BackColor = Color.RoyalBlue
            dmx16Value = True
            dmxdata(16) = 255
        ElseIf dmx16Value = True Then
            btn16.BackColor = Color.Transparent
            dmx16Value = False
            dmxdata(16) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn17_Click(sender As Object, e As EventArgs) Handles btn17.Click
        If dmx17Value = False Then
            btn17.BackColor = Color.RoyalBlue
            dmx17Value = True
            dmxdata(17) = 255
        ElseIf dmx17Value = True Then
            btn17.BackColor = Color.Transparent
            dmx17Value = False
            dmxdata(17) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn18_Click(sender As Object, e As EventArgs) Handles btn18.Click
        If dmx18Value = False Then
            btn18.BackColor = Color.RoyalBlue
            dmx18Value = True
            dmxdata(18) = 255
        ElseIf dmx18Value = True Then
            btn18.BackColor = Color.Transparent
            dmx18Value = False
            dmxdata(18) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn19_Click(sender As Object, e As EventArgs) Handles btn19.Click
        If dmx19Value = False Then
            btn19.BackColor = Color.RoyalBlue
            dmx19Value = True
            dmxdata(19) = 255
        ElseIf dmx19Value = True Then
            btn19.BackColor = Color.Transparent
            dmx19Value = False
            dmxdata(19) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn20_Click(sender As Object, e As EventArgs) Handles btn20.Click
        If dmx20Value = False Then
            btn20.BackColor = Color.RoyalBlue
            dmx20Value = True
            dmxdata(20) = 255
        ElseIf dmx20Value = True Then
            btn20.BackColor = Color.Transparent
            dmx20Value = False
            dmxdata(20) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn21_Click(sender As Object, e As EventArgs) Handles btn21.Click
        If dmx21Value = False Then
            btn21.BackColor = Color.RoyalBlue
            dmx21Value = True
            dmxdata(21) = 255
        ElseIf dmx21Value = True Then
            btn21.BackColor = Color.Transparent
            dmx21Value = False
            dmxdata(21) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn22_Click(sender As Object, e As EventArgs) Handles btn22.Click
        If dmx22Value = False Then
            btn22.BackColor = Color.RoyalBlue
            dmx22Value = True
            dmxdata(22) = 255
        ElseIf dmx22Value = True Then
            btn22.BackColor = Color.Transparent
            dmx22Value = False
            dmxdata(22) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn23_Click(sender As Object, e As EventArgs) Handles btn23.Click
        If dmx23Value = False Then
            btn23.BackColor = Color.RoyalBlue
            dmx23Value = True
            dmxdata(23) = 255
        ElseIf dmx23Value = True Then
            btn23.BackColor = Color.Transparent
            dmx23Value = False
            dmxdata(23) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn24_Click(sender As Object, e As EventArgs) Handles btn24.Click
        If dmx24Value = False Then
            btn24.BackColor = Color.RoyalBlue
            dmx24Value = True
            dmxdata(24) = 255
        ElseIf dmx24Value = True Then
            btn24.BackColor = Color.Transparent
            dmx24Value = False
            dmxdata(24) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn25_Click(sender As Object, e As EventArgs) Handles btn25.Click
        If dmx25Value = False Then
            btn25.BackColor = Color.RoyalBlue
            dmx25Value = True
            dmxdata(25) = 255
        ElseIf dmx25Value = True Then
            btn25.BackColor = Color.Transparent
            dmx25Value = False
            dmxdata(25) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn26_Click(sender As Object, e As EventArgs) Handles btn26.Click
        If dmx26Value = False Then
            btn26.BackColor = Color.RoyalBlue
            dmx26Value = True
            dmxdata(26) = 255
        ElseIf dmx26Value = True Then
            btn26.BackColor = Color.Transparent
            dmx26Value = False
            dmxdata(26) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn27_Click(sender As Object, e As EventArgs) Handles btn27.Click
        If dmx27Value = False Then
            btn27.BackColor = Color.RoyalBlue
            dmx27Value = True
            dmxdata(27) = 255
        ElseIf dmx27Value = True Then
            btn27.BackColor = Color.Transparent
            dmx27Value = False
            dmxdata(27) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn28_Click(sender As Object, e As EventArgs) Handles btn28.Click
        If dmx28Value = False Then
            btn28.BackColor = Color.RoyalBlue
            dmx28Value = True
            dmxdata(28) = 255
        ElseIf dmx28Value = True Then
            btn28.BackColor = Color.Transparent
            dmx28Value = False
            dmxdata(28) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn29_Click(sender As Object, e As EventArgs) Handles btn29.Click
        If dmx29Value = False Then
            btn29.BackColor = Color.RoyalBlue
            dmx29Value = True
            dmxdata(29) = 255
        ElseIf dmx29Value = True Then
            btn29.BackColor = Color.Transparent
            dmx29Value = False
            dmxdata(29) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn30_Click(sender As Object, e As EventArgs) Handles btn30.Click
        If dmx30Value = False Then
            btn30.BackColor = Color.RoyalBlue
            dmx30Value = True
            dmxdata(30) = 255
        ElseIf dmx30Value = True Then
            btn30.BackColor = Color.Transparent
            dmx30Value = False
            dmxdata(30) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn31_Click(sender As Object, e As EventArgs) Handles btn31.Click
        If dmx31Value = False Then
            btn31.BackColor = Color.RoyalBlue
            dmx31Value = True
            dmxdata(31) = 255
        ElseIf dmx31Value = True Then
            btn31.BackColor = Color.Transparent
            dmx31Value = False
            dmxdata(31) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn32_Click(sender As Object, e As EventArgs) Handles btn32.Click
        If dmx32Value = False Then
            btn32.BackColor = Color.RoyalBlue
            dmx32Value = True
            dmxdata(32) = 255
        ElseIf dmx32Value = True Then
            btn32.BackColor = Color.Transparent
            dmx32Value = False
            dmxdata(32) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn33_Click(sender As Object, e As EventArgs) Handles btn33.Click
        If dmx33Value = False Then
            btn33.BackColor = Color.RoyalBlue
            dmx33Value = True
            dmxdata(33) = 255
        ElseIf dmx33Value = True Then
            btn33.BackColor = Color.Transparent
            dmx33Value = False
            dmxdata(33) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn34_Click(sender As Object, e As EventArgs) Handles btn34.Click
        If dmx34Value = False Then
            btn34.BackColor = Color.RoyalBlue
            dmx34Value = True
            dmxdata(34) = 255
        ElseIf dmx34Value = True Then
            btn34.BackColor = Color.Transparent
            dmx34Value = False
            dmxdata(34) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn35_Click(sender As Object, e As EventArgs) Handles btn35.Click
        If dmx35Value = False Then
            btn35.BackColor = Color.RoyalBlue
            dmx35Value = True
            dmxdata(35) = 255
        ElseIf dmx35Value = True Then
            btn35.BackColor = Color.Transparent
            dmx35Value = False
            dmxdata(35) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn36_Click(sender As Object, e As EventArgs) Handles btn36.Click
        If dmx36Value = False Then
            btn36.BackColor = Color.RoyalBlue
            dmx36Value = True
            dmxdata(36) = 255
        ElseIf dmx36Value = True Then
            btn36.BackColor = Color.Transparent
            dmx36Value = False
            dmxdata(36) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn37_Click(sender As Object, e As EventArgs) Handles btn37.Click
        If dmx37Value = False Then
            btn37.BackColor = Color.RoyalBlue
            dmx37Value = True
            dmxdata(37) = 255
        ElseIf dmx37Value = True Then
            btn37.BackColor = Color.Transparent
            dmx37Value = False
            dmxdata(37) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn38_Click(sender As Object, e As EventArgs) Handles btn38.Click
        If dmx38Value = False Then
            btn38.BackColor = Color.RoyalBlue
            dmx38Value = True
            dmxdata(38) = 255
        ElseIf dmx38Value = True Then
            btn38.BackColor = Color.Transparent
            dmx38Value = False
            dmxdata(38) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn39_Click(sender As Object, e As EventArgs) Handles btn39.Click
        If dmx39Value = False Then
            btn39.BackColor = Color.RoyalBlue
            dmx39Value = True
            dmxdata(39) = 255
        ElseIf dmx39Value = True Then
            btn39.BackColor = Color.Transparent
            dmx39Value = False
            dmxdata(39) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn40_Click(sender As Object, e As EventArgs) Handles btn40.Click
        If dmx40Value = False Then
            btn40.BackColor = Color.RoyalBlue
            dmx40Value = True
            dmxdata(40) = 255
        ElseIf dmx40Value = True Then
            btn40.BackColor = Color.Transparent
            dmx40Value = False
            dmxdata(40) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn41_Click(sender As Object, e As EventArgs) Handles btn41.Click
        If dmx41Value = False Then
            btn41.BackColor = Color.RoyalBlue
            dmx41Value = True
            dmxdata(41) = 255
        ElseIf dmx41Value = True Then
            btn41.BackColor = Color.Transparent
            dmx41Value = False
            dmxdata(41) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn42_Click(sender As Object, e As EventArgs) Handles btn42.Click
        If dmx42Value = False Then
            btn42.BackColor = Color.RoyalBlue
            dmx42Value = True
            dmxdata(42) = 255
        ElseIf dmx42Value = True Then
            btn42.BackColor = Color.Transparent
            dmx42Value = False
            dmxdata(42) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn43_Click(sender As Object, e As EventArgs) Handles btn43.Click
        If dmx43Value = False Then
            btn43.BackColor = Color.RoyalBlue
            dmx43Value = True
            dmxdata(43) = 255
        ElseIf dmx43Value = True Then
            btn43.BackColor = Color.Transparent
            dmx43Value = False
            dmxdata(43) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn44_Click(sender As Object, e As EventArgs) Handles btn44.Click
        If dmx44Value = False Then
            btn44.BackColor = Color.RoyalBlue
            dmx44Value = True
            dmxdata(44) = 255
        ElseIf dmx44Value = True Then
            btn44.BackColor = Color.Transparent
            dmx44Value = False
            dmxdata(44) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn45_Click(sender As Object, e As EventArgs) Handles btn45.Click
        If dmx45Value = False Then
            btn45.BackColor = Color.RoyalBlue
            dmx45Value = True
            dmxdata(45) = 255
        ElseIf dmx45Value = True Then
            btn45.BackColor = Color.Transparent
            dmx45Value = False
            dmxdata(45) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn46_Click(sender As Object, e As EventArgs) Handles btn46.Click
        If dmx46Value = False Then
            btn46.BackColor = Color.RoyalBlue
            dmx46Value = True
            dmxdata(46) = 255
        ElseIf dmx46Value = True Then
            btn46.BackColor = Color.Transparent
            dmx46Value = False
            dmxdata(46) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn47_Click(sender As Object, e As EventArgs) Handles btn47.Click
        If dmx47Value = False Then
            btn47.BackColor = Color.RoyalBlue
            dmx47Value = True
            dmxdata(47) = 255
        ElseIf dmx47Value = True Then
            btn47.BackColor = Color.Transparent
            dmx47Value = False
            dmxdata(47) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btn48_Click(sender As Object, e As EventArgs) Handles btn48.Click
        If dmx48Value = False Then
            btn48.BackColor = Color.RoyalBlue
            dmx48Value = True
            dmxdata(48) = 255
        ElseIf dmx48Value = True Then
            btn48.BackColor = Color.Transparent
            dmx48Value = False
            dmxdata(48) = 0
        End If

        MainClass.sendDMXdata(dmxdata, 0)
    End Sub

    Private Sub btnAll48Off_Click(sender As Object, e As EventArgs) Handles btnAll48Off.Click
        For x = 1 To 48
            dmxdata(x) = 0
            Dim button As Button = TabPage3.Controls("btn" & x.ToString())
            button.BackColor = Color.Transparent
        Next
        MainClass.sendDMXdata(dmxdata, 0)
        setBtnFlag(False)
    End Sub

    Private Function setBtnFlag(mode As Boolean)
        dmx1Value = mode
        dmx2Value = mode
        dmx3Value = mode
        dmx4Value = mode
        dmx5Value = mode
        dmx6Value = mode
        dmx7Value = mode
        dmx8Value = mode
        dmx9Value = mode
        dmx10Value = mode
        dmx11Value = mode
        dmx12Value = mode
        dmx13Value = mode
        dmx14Value = mode
        dmx15Value = mode
        dmx16Value = mode
        dmx17Value = mode
        dmx18Value = mode
        dmx19Value = mode
        dmx20Value = mode
        dmx21Value = mode
        dmx22Value = mode
        dmx23Value = mode
        dmx24Value = mode
        dmx25Value = mode
        dmx26Value = mode
        dmx27Value = mode
        dmx28Value = mode
        dmx29Value = mode
        dmx30Value = mode
        dmx31Value = mode
        dmx32Value = mode
        dmx33Value = mode
        dmx34Value = mode
        dmx35Value = mode
        dmx36Value = mode
        dmx37Value = mode
        dmx38Value = mode
        dmx39Value = mode
        dmx40Value = mode
        dmx41Value = mode
        dmx42Value = mode
        dmx43Value = mode
        dmx44Value = mode
        dmx45Value = mode
        dmx46Value = mode
        dmx47Value = mode
        dmx48Value = mode

    End Function

    Private Sub btnAll48On_Click(sender As Object, e As EventArgs) Handles btnAll48On.Click
        For x = 1 To 48
            dmxdata(x) = 255
            Dim button As Button = TabPage3.Controls("btn" & x.ToString())
            button.BackColor = Color.RoyalBlue
        Next
        MainClass.sendDMXdata(dmxdata, 0)
        setBtnFlag(True)
    End Sub

    Private Async Sub btnMark_Click(sender As Object, e As EventArgs) Handles btnMark.Click
        Dim currentPosition As Double = Math.Round(AxWindowsMediaPlayer1.Ctlcontrols.currentPosition, 1)
        Dim r As String
        Dim s As Integer
        Dim t As DataTable

        'r = Await cls.DoesRecordExist("scenes", "name", "NEW")

        'If r = 0 Then
        ' create new record
        s = Await cls.AddNewSliceCall("NEW", currentPosition, txtContentID.Text)
        'Else
        'Using New Centered_MessageBox(Me)
        'MessageBox.Show("A NEW record already exists. Use this first")
        'End Using
        'End If

        Await populateDG1(txtContentID.Text)

        'dgSliceCalls.Rows.Add(New String() {dgSliceCalls.NewRowIndex, currentPosition, "", "EDIT"})
    End Sub

    Private Sub dgslicecalls_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Try
            SetupDgSliceCalls()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    Private Sub dgfixtures_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Try
            SetupDgfixtures()
            SetupDgChannels()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    Private Sub dgavailabledevices_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Try
            SetupDgavailabledevices()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    Private Sub dgshowlayouts_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Try
            SetupDgshowlayouts()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    Public Sub SetupDgavailabledevices()
        'fill dataTableData with database data
        Dim cls As New dataclass
        Dim t As DataTable

        Dim c0 As New DataGridViewTextBoxColumn
        With c0
            .DataPropertyName = "id"
            .HeaderText = "ID"
            .Width = 20
            .MinimumWidth = 20
            .Visible = False
        End With
        Dim c1 As New DataGridViewTextBoxColumn
        With c1
            .DataPropertyName = "name"
            .HeaderText = "Device Name"
            .Width = 200
            .MinimumWidth = 200
        End With
        Dim c2 As New DataGridViewTextBoxColumn
        With c2
            .DataPropertyName = "quantity"
            .HeaderText = "Qty Available"
            .Width = 100
            .MinimumWidth = 100
        End With
        Dim c3 As New DataGridViewButtonColumn
        With c3
            .DataPropertyName = "action"
            .HeaderText = ""
            .Width = 89
            .MinimumWidth = 89
            .Text = "ADD"
            .UseColumnTextForButtonValue = True
        End With

        dgavailabledevices.Columns.AddRange({c1, c2, c3})
        dgavailabledevices.ColumnHeadersVisible = True
    End Sub

    Public Sub SetupDgshowlayouts()
        'fill dataTableData with database data
        Dim cls As New dataclass
        Dim t As DataTable

        Dim c0 As New DataGridViewTextBoxColumn
        With c0
            .DataPropertyName = "id"
            .HeaderText = "ID"
            .Width = 20
            .MinimumWidth = 20
            .Visible = False
        End With
        Dim c1 As New DataGridViewTextBoxColumn
        With c1
            .DataPropertyName = "devicename"
            .HeaderText = "Device Name"
            .Width = 166
            .MinimumWidth = 166
        End With
        Dim c2 As New DataGridViewTextBoxColumn
        With c2
            .DataPropertyName = "devicealias"
            .HeaderText = "Device Alias Name"
            .Width = 170
            .MinimumWidth = 170
        End With
        Dim c3 As New DataGridViewTextBoxColumn
        With c3
            .DataPropertyName = "devicegroups"
            .HeaderText = "Group Membership"
            .Width = 200
            .MinimumWidth = 200
        End With
        Dim c4 As New DataGridViewComboBoxColumn
        t = cls.LoadComboDefaultRanges(255)
        With c4
            .DataSource = t
            .ValueMember = "item"
            .DisplayMember = "item"
            .DataPropertyName = "dmxid"
            .HeaderText = "DMX ID"
            .Width = 60
            .MinimumWidth = 60
            .FlatStyle = FlatStyle.Flat
            .DisplayStyleForCurrentCellOnly = True
        End With
        'Dim c5 As New DataGridViewComboBoxColumn
        't = cls.LoadComboDefaultRanges(255)
        'With c5
        '    .DataSource = t
        '    .ValueMember = "item"
        '    .DisplayMember = "item"
        '    .DataPropertyName = "mirrorx"
        '    .HeaderText = "Mirror X"
        '    .Width = 60
        '    .MinimumWidth = 60
        '    .FlatStyle = FlatStyle.Flat
        '    .DisplayStyleForCurrentCellOnly = True
        'End With
        'Dim c6 As New DataGridViewComboBoxColumn
        't = cls.LoadComboDefaultRanges(255)
        'With c6
        '    .DataSource = t
        '    .ValueMember = "item"
        '    .DisplayMember = "item"
        '    .DataPropertyName = "mirrory"
        '    .HeaderText = "Mirror Y"
        '    .Width = 60
        '    .MinimumWidth = 60
        '    .FlatStyle = FlatStyle.Flat
        '    .DisplayStyleForCurrentCellOnly = True
        'End With

        dgshowlayouts.Columns.AddRange({c1, c2, c3, c4})
        dgshowlayouts.ColumnHeadersVisible = True
    End Sub

    Public Async Sub SetupDgSliceCalls()
        'fill dataTableData with database data
        Dim cls As New dataclass
        cls.InitiateDb()

        Dim t As DataTable

        Dim c1 As New DataGridViewTextBoxColumn
        With c1
            .DataPropertyName = "_id"
            .HeaderText = "Object ID"
            .Width = 20
            .MinimumWidth = 20
            .Visible = False
        End With

        Dim c2 As New DataGridViewTextBoxColumn
        With c2
            .DataPropertyName = "comment"
            .HeaderText = "Comment"
            .Width = 200
            .MinimumWidth = 200
        End With

        Dim c3 As New DataGridViewTextBoxColumn
        With c3
            .DataPropertyName = "position"
            .HeaderText = "Position"
            .Width = 60
            .MinimumWidth = 60
        End With

        Dim c4 As New DataGridViewComboBoxColumn
        t = Await cls.GetSlices()
        With c4
            .DataSource = t
            .ValueMember = "_id"
            .DisplayMember = "name"
            .DataPropertyName = "slicename"
            .HeaderText = "Slice Name"
            .Width = 200
            .MinimumWidth = 200
            .FlatStyle = FlatStyle.Flat
            .DisplayStyleForCurrentCellOnly = True
        End With

        Dim c5 As New DataGridViewTextBoxColumn
        With c5
            .DataPropertyName = "contentid"
            .HeaderText = "Content ID"
            .Width = 60
            .MinimumWidth = 60
            .Visible = False
        End With

        dgSliceCalls.Columns.AddRange({c1, c2, c3, c4, c5})
        dgSliceCalls.ColumnHeadersVisible = True

    End Sub

    Public Sub SetupDgfixtures()
        'fill dataTableData with database data
        Dim cls As New dataclass
        Dim t As DataTable

        Dim c0 As New DataGridViewTextBoxColumn
        With c0
            .DataPropertyName = "id"
            .HeaderText = "ID"
            .Width = 20
            .MinimumWidth = 20
            .Visible = False
        End With
        Dim c1 As New DataGridViewTextBoxColumn
        With c1
            .DataPropertyName = "name"
            .HeaderText = "Fixture Name"
            .Width = 150
            .MinimumWidth = 150
        End With
        Dim c2 As New DataGridViewTextBoxColumn
        With c2
            .DataPropertyName = "manufacturer"
            .HeaderText = "Manufacturer"
            .Width = 150
            .MinimumWidth = 150
        End With
        Dim c3 As New DataGridViewTextBoxColumn
        With c3
            .DataPropertyName = "model"
            .HeaderText = "Model"
            .Width = 140
            .MinimumWidth = 140
        End With
        Dim c4 As New DataGridViewComboBoxColumn
        t = cls.LoadComboDefaultRanges(255)
        With c4
            .DataSource = t
            .ValueMember = "item"
            .DisplayMember = "item"
            .DataPropertyName = "channels"
            .HeaderText = "Channels"
            .Width = 60
            .MinimumWidth = 60
            .FlatStyle = FlatStyle.Flat
            .DisplayStyleForCurrentCellOnly = True
        End With
        Dim c5 As New DataGridViewComboBoxColumn
        t = cls.LoadComboDefaultRanges(10)
        With c5
            .DataSource = t
            .ValueMember = "item"
            .DisplayMember = "item"
            .DataPropertyName = "quantity"
            .HeaderText = "Quantity"
            .Width = 60
            .MinimumWidth = 60
            .FlatStyle = FlatStyle.Flat
            .DisplayStyleForCurrentCellOnly = True
        End With

        dgFixtures.Columns.AddRange({c1, c2, c3, c4, c5})
        dgFixtures.ColumnHeadersVisible = True

    End Sub

    Public Sub SetupDgChannels()
        'fill dataTableData with database data
        Dim cls As New dataclass
        Dim t As DataTable

        Dim c1 As New DataGridViewComboBoxColumn

        t = cls.LoadComboDefaultRanges(255)
        With c1
            .DataSource = t
            .ValueMember = "item"
            .DisplayMember = "item"
            .DataPropertyName = "startchannel"
            .HeaderText = "Channel"
            .Width = 60
            .MinimumWidth = 60
            .FlatStyle = FlatStyle.Flat
            .DisplayStyleForCurrentCellOnly = True
        End With
        Dim c2 As New DataGridViewComboBoxColumn
        t = cls.LoadComboDefaultRanges(255)
        With c2
            .DataSource = t
            .ValueMember = "item"
            .DisplayMember = "item"
            .DataPropertyName = "dmxvalue"
            .HeaderText = "DMX Start"
            .Width = 66
            .MinimumWidth = 66
            .FlatStyle = FlatStyle.Flat
            .DisplayStyleForCurrentCellOnly = True
        End With
        Dim c3 As New DataGridViewTextBoxColumn
        With c3
            .DataPropertyName = "dmxfunctioncategory"
            .HeaderText = "Category"
            .Width = 140
            .MinimumWidth = 140
        End With
        Dim c4 As New DataGridViewTextBoxColumn
        With c4
            .DataPropertyName = "dmxfunction"
            .HeaderText = "Function"
            .Width = 165
            .MinimumWidth = 165
        End With

        dgChannels.Columns.AddRange({c1, c2, c3, c4})
        dgChannels.ColumnHeadersVisible = True

    End Sub

    Async Sub SetupDgScenes()
        'fill dataTableData with database data
        Dim cls As New dataclass
        Dim t As DataTable

        Dim c1 As New DataGridViewTextBoxColumn
        With c1
            .DataPropertyName = "name"
            .HeaderText = "Scene Name"
            .Width = 260
            .MinimumWidth = 260
        End With
        Dim c2 As New DataGridViewTextBoxColumn
        With c2
            .DataPropertyName = "type"
            .HeaderText = "Scene Type"
            .Width = 120
            .MinimumWidth = 120
        End With
        Dim c3 As New DataGridViewComboBoxColumn
        t = cls.LoadComboDefaultRanges(255)
        With c3
            .DataSource = t
            .ValueMember = "item"
            .DisplayMember = "item"
            .DataPropertyName = "actions"
            .HeaderText = "Scene Actions"
            .Width = 120
            .MinimumWidth = 120
        End With

        dgScenes.Columns.AddRange({c1, c2, c3})
        dgScenes.ColumnHeadersVisible = True

    End Sub

    Private Async Function autoRunSlice(keyvalue As String) As Task
        Dim t As DataTable

        t = Await cls.GetSlice("_id", keyvalue)
        Dim name As String = t.Rows(0)("name")
        Dim dmxstring As String = t.Rows(0)("dmxstring")
        Dim nextslice As String = t.Rows(0)("nextslice")

        ' decode dmxstring back to dmx byte array
        dmxdata = System.Text.Encoding.Default.GetBytes(dmxstring)

        ' send dmx
        MainClass.sendDMXdata(dmxdata, 0)

    End Function


    Private Async Sub wmProgressTimer_Tick(sender As Object, e As EventArgs) Handles wmProgressTimer.Tick
        Dim currentposition As Double = AxWindowsMediaPlayer1.Ctlcontrols.currentPosition
        Dim keyvalue As String

        ' update visual 
        txtCurrentPosition.Text = currentposition

        ' calculate if match
        For Each item As Double In posArray
            If Math.Round(currentposition, 1) = item And item <> 0 Then

                ' match found
                dgSliceCalls.Rows(posArrayCount).DefaultCellStyle.BackColor = Color.LightGreen

                If IsDBNull(dgSliceCalls.Rows(posArrayCount).Cells(4).Value) Then
                    ' no slice defined, move on
                Else
                    ' retrieve slice and run
                    keyvalue = dgSliceCalls.Rows(posArrayCount).Cells(4).Value
                    Await autoRunSlice(keyvalue)
                End If

                ' clear prior row highlight
                If posArrayCount >= 1 And posArrayCount < dgSliceCalls.Rows.Count Then
                    dgSliceCalls.Rows(posArrayCount - 1).DefaultCellStyle.BackColor = Color.White
                End If
                If posArrayCount < dgSliceCalls.Rows.Count Then
                    posArrayCount = posArrayCount + 1
                Else
                    posArrayCount = 0
                End If

            End If
        Next

    End Sub

    Private Sub btnPlay_Click(sender As Object, e As EventArgs) Handles btnPlay.Click
        Dim row As DataGridViewRow
        For Each row In dgSliceCalls.Rows
            posArray.Add(row.Cells(2).Value)
        Next

        AxWindowsMediaPlayer1.Ctlcontrols.stop()
        AxWindowsMediaPlayer1.Ctlcontrols.currentPosition = 0

        txtCurrentPosition.Text = ""
        wmProgressTimer.Stop()
        posArrayCount = 0

        wmProgressTimer.Start()
        AxWindowsMediaPlayer1.Ctlcontrols.play()
    End Sub

    Private Sub btnSceneClear_Click(sender As Object, e As EventArgs) Handles btnSceneClear.Click
        dgSliceCalls.DataSource = Nothing
        dgSliceCalls.Columns.Clear()

    End Sub


    Async Sub btnAddFixture_Click(sender As Object, e As EventArgs) Handles btnAddFixture.Click
        Dim r As String
        Dim l As DataTable

        r = Await cls.DoesRecordExist("fixtures", "name", "NEW")

        If r = 0 Then
            ' create new record
            r = Await cls.AddNewRecord("NEW")
        Else
            Using New Centered_MessageBox(Me)
                MessageBox.Show("A NEW record already exists. Use this first")
            End Using
        End If

        l = Await cls.GetDevices()

        dgFixtures.DataSource = l
        dgFixtures.Refresh()

    End Sub

    Async Sub btnDelFixture_Click(sender As Object, e As EventArgs) Handles btnDelFixture.Click
        Dim r As Integer
        Dim l As DataTable

        Using New Centered_MessageBox(Me)
            If MessageBox.Show("Are you sure you want to delete selected rows?", "Delete Confirmation", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                For Each row As DataGridViewRow In dgFixtures.SelectedRows
                    r = Await cls.DeleteRecord(row.Cells(0).Value)
                Next
                l = Await cls.GetDevices()

                dgFixtures.DataSource = l
                dgFixtures.Refresh()

            Else
                Exit Sub
            End If
        End Using
    End Sub

    Async Sub dgFixtures_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgFixtures.CellEndEdit

        Dim fieldname As String = dgFixtures.Columns(e.ColumnIndex).DataPropertyName
        Dim keyname As String = dgFixtures.Rows(e.RowIndex).Cells(0).Value
        Dim fieldvalue As String = dgFixtures.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
        Dim r As String

        ' null check
        If Not IsDBNull(dgFixtures.Columns(e.ColumnIndex).DataPropertyName) Then
            fieldname = dgFixtures.Columns(e.ColumnIndex).DataPropertyName
        End If
        If Not IsDBNull(dgFixtures.Rows(e.RowIndex).Cells(0).Value) Then
            keyname = dgFixtures.Rows(e.RowIndex).Cells(0).Value
        End If
        If Not IsDBNull(dgFixtures.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) Then
            fieldvalue = dgFixtures.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
        End If

        ' check if record exists
        If e.ColumnIndex = 0 Then
            r = Await cls.DoesRecordExist("fixtures", "name", fieldvalue)
            If r = 1 Then
                Using New Centered_MessageBox(Me)
                    MessageBox.Show("Fixture already exists. Please use a new name")
                End Using
            Else
                ' create new record
                r = Await cls.AddNewRecord(fieldvalue)
            End If
        Else
            ' create new record
            r = Await cls.UpdateRecord(keyname, fieldname, fieldvalue)
        End If

    End Sub


    Public Async Function LoadLocation() As Task
        Dim d1 As New DataTable

        d1 = Await cls.GetLayouts()
        cmboLayouts.DataSource = d1
        cmboLayouts.ValueMember = "name"
        cmboLayouts.DisplayMember = "name"
    End Function

    Private Async Function populateDG2(s As String) As Task
        Dim l As DataTable

        dgChannels.Columns.Clear()
        dgChannels.DataSource = Nothing
        SetupDgChannels()

        l = Await cls.GetDeviceConfiguration("_id", s)

        dgChannels.DataSource = l
        dgChannels.Refresh()

    End Function

    Private Async Function populateDG1(contentid As String) As Task
        Dim t As DataTable

        dgSliceCalls.DataSource = Nothing
        dgSliceCalls.Columns.Clear()
        SetupDgSliceCalls()


        t = Await cls.GetSliceCalls(contentid)

        dgSliceCalls.DataSource = t

    End Function

    Private Sub dgFixtures_RowEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgFixtures.RowEnter
        Dim l As Task

        If dgFixtures.SelectedRows.Count > 0 Then

            Dim i As Integer = dgFixtures.SelectedRows(0).Index
            Dim s As String = dgFixtures.Item(0, i).Value

            l = populateDG2(s)
        End If
    End Sub


    Private Async Sub btnAddConfiguration_Click(sender As Object, e As EventArgs) Handles btnAddConfiguration.Click
        Dim i = dgFixtures.CurrentRow.Index
        Dim keyvalue = dgFixtures.Item(0, i).Value
        Dim r As Integer

        ' add new item to json record
        r = Await cls.AddNewConfigRecord(keyvalue)

        ' load collection


        Dim l As DataTable

        l = Await cls.GetDeviceConfiguration("_id", keyvalue)

        dgChannels.DataSource = l
        dgChannels.Refresh()

        dgChannels.CurrentCell = dgChannels.Rows(dgChannels.RowCount - 1).Cells(0)

    End Sub

    Private Sub btnLasersON_Click(sender As Object, e As EventArgs)
        Timer2.Stop()
        dmxdata(1) = 255 'dmx mode
        dmxdata(2) = 0 ' pan
        dmxdata(3) = 255 ' tilt
        dmxdata(4) = 100 ' vertical speed
        dmxdata(5) = 2 ' group 1 pattern (2-220)
        dmxdata(6) = 255 ' group 1 laser on
        dmxdata(7) = 0 ' group 2 pattern selection (0-28)
        dmxdata(8) = 0 ' group 2 laser on/off
        dmxdata(9) = 0 ' dual group animations
        dmxdata(10) = 0 ' moving x
        dmxdata(11) = 0 ' moving y
        dmxdata(12) = 255 ' x axis rotation
        dmxdata(13) = 0 ' y access rotation
        dmxdata(14) = 0 ' rotation
        dmxdata(15) = 255 ' zoom
        dmxdata(16) = 0 ' sine wave flux
        dmxdata(17) = 0 ' flux speed
        dmxdata(18) = 255 ' drawing
        MainClass.sendDMXdata(dmxdata)
    End Sub

    Private Sub lblCh1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub txtCh1_TextChanged(sender As Object, e As EventArgs) Handles txtCh1.TextChanged
        If txtCh1.Text = "" Then txtCh1.Text = 0
        If txtCh1.Text > 255 Then txtCh1.Text = 255
        If ch1.Value < 256 And txtCh1.Text < 256 Then
            ch1.Value = txtCh1.Text

            ' get device id
            If cmboDevice.Text = "" Then Exit Sub
            Dim r As DataRowView = cmboDevice.SelectedItem
            Dim s As String = r(2)

            ' get and send dmx
            dmxdata(s + 0) = CInt(txtCh1.Text)
            MainClass.sendDMXdata(dmxdata)
        End If
    End Sub

    Private Sub txtCh2_TextChanged(sender As Object, e As EventArgs) Handles txtCh2.TextChanged
        If txtCh2.Text = "" Then txtCh2.Text = 0
        If txtCh2.Text > 255 Then txtCh2.Text = 255
        If ch2.Value < 256 And txtCh2.Text < 256 Then
            ch2.Value = txtCh2.Text

            ' get device id
            If cmboDevice.Text = "" Then Exit Sub
            Dim r As DataRowView = cmboDevice.SelectedItem
            Dim s As String = r(2)

            ' get and send dmx
            dmxdata(s + 1) = CInt(txtCh2.Text)
            MainClass.sendDMXdata(dmxdata)
        End If
    End Sub

    Private Sub txtCh3_TextChanged(sender As Object, e As EventArgs) Handles txtCh3.TextChanged
        If txtCh3.Text = "" Then txtCh3.Text = 0
        If txtCh3.Text > 255 Then txtCh3.Text = 255
        If ch3.Value < 256 And txtCh3.Text < 256 Then
            ch3.Value = txtCh3.Text

            ' get device id
            If cmboDevice.Text = "" Then Exit Sub
            Dim r As DataRowView = cmboDevice.SelectedItem
            Dim s As String = r(2)

            ' get and send dmx
            dmxdata(s + 2) = CInt(txtCh3.Text)
            MainClass.sendDMXdata(dmxdata)
        End If
    End Sub

    Private Sub txtCh4_TextChanged(sender As Object, e As EventArgs) Handles txtCh4.TextChanged
        If txtCh4.Text = "" Then txtCh4.Text = 0
        If txtCh4.Text > 255 Then txtCh4.Text = 255
        If ch4.Value < 256 And txtCh4.Text < 256 Then
            ch4.Value = txtCh4.Text

            ' get device id
            If cmboDevice.Text = "" Then Exit Sub
            Dim r As DataRowView = cmboDevice.SelectedItem
            Dim s As String = r(2)

            ' get and send dmx
            dmxdata(s + 3) = CInt(txtCh4.Text)
            MainClass.sendDMXdata(dmxdata)
        End If
    End Sub

    Private Sub txtCh5_TextChanged(sender As Object, e As EventArgs) Handles txtCh5.TextChanged
        If txtCh5.Text = "" Then txtCh5.Text = 0
        If txtCh5.Text > 255 Then txtCh5.Text = 255
        If ch5.Value < 256 And txtCh5.Text < 256 Then
            ch5.Value = txtCh5.Text

            ' get device id
            If cmboDevice.Text = "" Then Exit Sub
            Dim r As DataRowView = cmboDevice.SelectedItem
            Dim s As String = r(2)

            ' get and send dmx
            dmxdata(s + 4) = CInt(txtCh5.Text)
            MainClass.sendDMXdata(dmxdata)
        End If
    End Sub

    Private Sub txtCh6_TextChanged(sender As Object, e As EventArgs) Handles txtCh6.TextChanged
        If txtCh6.Text = "" Then txtCh6.Text = 0
        If txtCh6.Text > 255 Then txtCh6.Text = 255
        If ch6.Value < 256 And txtCh6.Text < 256 Then
            ch6.Value = txtCh6.Text

            ' get device id
            If cmboDevice.Text = "" Then Exit Sub
            Dim r As DataRowView = cmboDevice.SelectedItem
            Dim s As String = r(2)

            ' get and send dmx
            dmxdata(s + 5) = CInt(txtCh6.Text)
            MainClass.sendDMXdata(dmxdata)
        End If
    End Sub

    Private Sub txtCh7_TextChanged(sender As Object, e As EventArgs) Handles txtCh7.TextChanged
        If txtCh7.Text = "" Then txtCh7.Text = 0
        If txtCh7.Text > 255 Then txtCh7.Text = 255
        If ch7.Value < 256 And txtCh7.Text < 256 Then
            ch7.Value = txtCh7.Text

            ' get device id
            If cmboDevice.Text = "" Then Exit Sub
            Dim r As DataRowView = cmboDevice.SelectedItem
            Dim s As String = r(2)

            ' get and send dmx
            dmxdata(s + 6) = CInt(txtCh7.Text)
            MainClass.sendDMXdata(dmxdata)
        End If
    End Sub

    Private Sub txtCh8_TextChanged(sender As Object, e As EventArgs) Handles txtCh8.TextChanged
        If txtCh8.Text = "" Then txtCh8.Text = 0
        If txtCh8.Text > 255 Then txtCh8.Text = 255
        If ch8.Value < 256 And txtCh8.Text < 256 Then
            ch8.Value = txtCh8.Text

            ' get device id
            If cmboDevice.Text = "" Then Exit Sub
            Dim r As DataRowView = cmboDevice.SelectedItem
            Dim s As String = r(2)

            ' get and send dmx
            dmxdata(s + 7) = CInt(txtCh8.Text)
            MainClass.sendDMXdata(dmxdata)
        End If
    End Sub

    Private Sub txtCh9_TextChanged(sender As Object, e As EventArgs) Handles txtCh9.TextChanged
        If txtCh9.Text = "" Then txtCh9.Text = 0
        If txtCh9.Text > 255 Then txtCh9.Text = 255
        If ch9.Value < 256 And txtCh9.Text < 256 Then
            ch9.Value = txtCh9.Text

            ' get device id
            If cmboDevice.Text = "" Then Exit Sub
            Dim r As DataRowView = cmboDevice.SelectedItem
            Dim s As String = r(2)

            ' get and send dmx
            dmxdata(s + 8) = CInt(txtCh9.Text)
            MainClass.sendDMXdata(dmxdata)
        End If
    End Sub

    Private Sub txtCh10_TextChanged(sender As Object, e As EventArgs) Handles txtCh10.TextChanged
        If txtCh10.Text = "" Then txtCh10.Text = 0
        If txtCh10.Text > 255 Then txtCh10.Text = 255
        If ch10.Value < 256 And txtCh10.Text < 256 Then
            ch10.Value = txtCh10.Text

            ' get device id
            If cmboDevice.Text = "" Then Exit Sub
            Dim r As DataRowView = cmboDevice.SelectedItem
            Dim s As String = r(2)

            ' get and send dmx
            dmxdata(s + 9) = CInt(txtCh10.Text)
            MainClass.sendDMXdata(dmxdata)
        End If
    End Sub

    Private Sub txtCh11_TextChanged(sender As Object, e As EventArgs) Handles txtCh11.TextChanged
        If txtCh11.Text = "" Then txtCh11.Text = 0
        If txtCh11.Text > 255 Then txtCh11.Text = 255
        If ch11.Value < 256 And txtCh11.Text < 256 Then
            ch11.Value = txtCh11.Text

            ' get device id
            If cmboDevice.Text = "" Then Exit Sub
            Dim r As DataRowView = cmboDevice.SelectedItem
            Dim s As String = r(2)

            ' get and send dmx
            dmxdata(s + 10) = CInt(txtCh11.Text)
            MainClass.sendDMXdata(dmxdata)
        End If
    End Sub

    Private Sub txtCh12_TextChanged(sender As Object, e As EventArgs) Handles txtCh12.TextChanged
        If txtCh12.Text = "" Then txtCh12.Text = 0
        If txtCh12.Text > 255 Then txtCh12.Text = 255
        If ch12.Value < 256 And txtCh12.Text < 256 Then
            ch12.Value = txtCh12.Text

            ' get device id
            If cmboDevice.Text = "" Then Exit Sub
            Dim r As DataRowView = cmboDevice.SelectedItem
            Dim s As String = r(2)

            ' get and send dmx
            dmxdata(s + 11) = CInt(txtCh12.Text)
            MainClass.sendDMXdata(dmxdata)
        End If
    End Sub

    Private Sub txtCh13_TextChanged(sender As Object, e As EventArgs) Handles txtCh13.TextChanged
        If txtCh13.Text = "" Then txtCh13.Text = 0
        If txtCh13.Text > 255 Then txtCh13.Text = 255
        If ch13.Value < 256 And txtCh13.Text < 256 Then
            ch13.Value = txtCh13.Text

            ' get device id
            If cmboDevice.Text = "" Then Exit Sub
            Dim r As DataRowView = cmboDevice.SelectedItem
            Dim s As String = r(2)

            ' get and send dmx
            dmxdata(s + 12) = CInt(txtCh13.Text)
            MainClass.sendDMXdata(dmxdata)
        End If
    End Sub

    Private Sub txtCh14_TextChanged(sender As Object, e As EventArgs) Handles txtCh14.TextChanged
        If txtCh14.Text = "" Then txtCh14.Text = 0
        If txtCh14.Text > 255 Then txtCh14.Text = 255
        If ch14.Value < 256 And txtCh14.Text < 256 Then
            ch14.Value = txtCh14.Text

            ' get device id
            If cmboDevice.Text = "" Then Exit Sub
            Dim r As DataRowView = cmboDevice.SelectedItem
            Dim s As String = r(2)

            ' get and send dmx
            dmxdata(s + 13) = CInt(txtCh14.Text)
            MainClass.sendDMXdata(dmxdata)
        End If
    End Sub

    Private Sub txtCh15_TextChanged(sender As Object, e As EventArgs) Handles txtCh15.TextChanged
        If txtCh15.Text = "" Then txtCh15.Text = 0
        If txtCh15.Text > 255 Then txtCh15.Text = 255
        If ch15.Value < 256 And txtCh15.Text < 256 Then
            ch15.Value = txtCh15.Text

            ' get device id
            If cmboDevice.Text = "" Then Exit Sub
            Dim r As DataRowView = cmboDevice.SelectedItem
            Dim s As String = r(2)

            ' get and send dmx
            dmxdata(s + 14) = CInt(txtCh15.Text)
            MainClass.sendDMXdata(dmxdata)
        End If
    End Sub

    Private Sub txtCh16_TextChanged(sender As Object, e As EventArgs) Handles txtCh16.TextChanged
        If txtCh16.Text = "" Then txtCh16.Text = 0
        If txtCh16.Text > 255 Then txtCh16.Text = 255
        If ch16.Value < 256 And txtCh16.Text < 256 Then
            ch16.Value = txtCh16.Text

            ' get device id
            If cmboDevice.Text = "" Then Exit Sub
            Dim r As DataRowView = cmboDevice.SelectedItem
            Dim s As String = r(2)

            ' get and send dmx
            dmxdata(s + 15) = CInt(txtCh16.Text)
            MainClass.sendDMXdata(dmxdata)
        End If
    End Sub

    Private Sub txtCh17_TextChanged(sender As Object, e As EventArgs) Handles txtCh17.TextChanged
        If txtCh17.Text = "" Then txtCh17.Text = 0
        If txtCh17.Text > 255 Then txtCh17.Text = 255
        If ch17.Value < 256 And txtCh17.Text < 256 Then
            ch17.Value = txtCh17.Text

            ' get device id
            If cmboDevice.Text = "" Then Exit Sub
            Dim r As DataRowView = cmboDevice.SelectedItem
            Dim s As String = r(2)

            ' get and send dmx
            dmxdata(s + 16) = CInt(txtCh17.Text)
            MainClass.sendDMXdata(dmxdata)
        End If
    End Sub

    Private Sub txtCh18_TextChanged(sender As Object, e As EventArgs) Handles txtCh18.TextChanged
        If txtCh18.Text = "" Then txtCh18.Text = 0
        If txtCh18.Text > 255 Then txtCh18.Text = 255
        If ch18.Value < 256 And txtCh18.Text < 256 Then
            ch18.Value = txtCh18.Text

            ' get device id
            If cmboDevice.Text = "" Then Exit Sub
            Dim r As DataRowView = cmboDevice.SelectedItem
            Dim s As String = r(2)

            ' get and send dmx
            dmxdata(s + 17) = CInt(txtCh18.Text)
            MainClass.sendDMXdata(dmxdata)
        End If
    End Sub

    Private Sub ch1_Scroll(sender As Object, e As EventArgs) Handles ch1.Scroll
        Dim trkBar As TrackBar = CType(sender, TrackBar)
        'Set timer value based on the Selection
        txtCh1.Text = trkBar.Value
    End Sub
    Private Sub ch2_Scroll(sender As Object, e As EventArgs) Handles ch2.Scroll
        Dim trkBar As TrackBar = CType(sender, TrackBar)
        'Set timer value based on the Selection
        txtCh2.Text = trkBar.Value
    End Sub
    Private Sub ch3_Scroll(sender As Object, e As EventArgs) Handles ch3.Scroll
        Dim trkBar As TrackBar = CType(sender, TrackBar)
        'Set timer value based on the Selection
        txtCh3.Text = trkBar.Value
    End Sub
    Private Sub ch4_Scroll(sender As Object, e As EventArgs) Handles ch4.Scroll
        Dim trkBar As TrackBar = CType(sender, TrackBar)
        'Set timer value based on the Selection
        txtCh4.Text = trkBar.Value
    End Sub
    Private Sub ch5_Scroll(sender As Object, e As EventArgs) Handles ch5.Scroll
        Dim trkBar As TrackBar = CType(sender, TrackBar)
        'Set timer value based on the Selection
        txtCh5.Text = trkBar.Value
    End Sub
    Private Sub ch6_Scroll(sender As Object, e As EventArgs) Handles ch6.Scroll
        Dim trkBar As TrackBar = CType(sender, TrackBar)
        'Set timer value based on the Selection
        txtCh6.Text = trkBar.Value
    End Sub
    Private Sub ch7_Scroll(sender As Object, e As EventArgs) Handles ch7.Scroll
        Dim trkBar As TrackBar = CType(sender, TrackBar)
        'Set timer value based on the Selection
        txtCh7.Text = trkBar.Value
    End Sub
    Private Sub ch8_Scroll(sender As Object, e As EventArgs) Handles ch8.Scroll
        Dim trkBar As TrackBar = CType(sender, TrackBar)
        'Set timer value based on the Selection
        txtCh8.Text = trkBar.Value
    End Sub
    Private Sub ch9_Scroll(sender As Object, e As EventArgs) Handles ch9.Scroll
        Dim trkBar As TrackBar = CType(sender, TrackBar)
        'Set timer value based on the Selection
        txtCh9.Text = trkBar.Value
    End Sub
    Private Sub ch10_Scroll(sender As Object, e As EventArgs) Handles ch10.Scroll
        Dim trkBar As TrackBar = CType(sender, TrackBar)
        'Set timer value based on the Selection
        txtCh10.Text = trkBar.Value
    End Sub
    Private Sub ch11_Scroll(sender As Object, e As EventArgs) Handles ch11.Scroll
        Dim trkBar As TrackBar = CType(sender, TrackBar)
        'Set timer value based on the Selection
        txtCh11.Text = trkBar.Value
    End Sub
    Private Sub ch12_Scroll(sender As Object, e As EventArgs) Handles ch12.Scroll
        Dim trkBar As TrackBar = CType(sender, TrackBar)
        'Set timer value based on the Selection
        txtCh12.Text = trkBar.Value
    End Sub
    Private Sub ch13_Scroll(sender As Object, e As EventArgs) Handles ch13.Scroll
        Dim trkBar As TrackBar = CType(sender, TrackBar)
        'Set timer value based on the Selection
        txtCh13.Text = trkBar.Value
    End Sub
    Private Sub ch14_Scroll(sender As Object, e As EventArgs) Handles ch14.Scroll
        Dim trkBar As TrackBar = CType(sender, TrackBar)
        'Set timer value based on the Selection
        txtCh14.Text = trkBar.Value
    End Sub
    Private Sub ch15_Scroll(sender As Object, e As EventArgs) Handles ch15.Scroll
        Dim trkBar As TrackBar = CType(sender, TrackBar)
        'Set timer value based on the Selection
        txtCh15.Text = trkBar.Value
    End Sub
    Private Sub ch16_Scroll(sender As Object, e As EventArgs) Handles ch16.Scroll
        Dim trkBar As TrackBar = CType(sender, TrackBar)
        'Set timer value based on the Selection
        txtCh16.Text = trkBar.Value
    End Sub
    Private Sub ch17_Scroll(sender As Object, e As EventArgs) Handles ch17.Scroll
        Dim trkBar As TrackBar = CType(sender, TrackBar)
        'Set timer value based on the Selection
        txtCh17.Text = trkBar.Value
    End Sub
    Private Sub ch18_Scroll(sender As Object, e As EventArgs) Handles ch18.Scroll
        Dim trkBar As TrackBar = CType(sender, TrackBar)
        'Set timer value based on the Selection
        txtCh18.Text = trkBar.Value
    End Sub

    Private Sub txtCh1_MouseEnter(sender As Object, e As EventArgs)
        txtCh1.SelectAll()
    End Sub

    Private Async Sub cmboDevice_SelectedValueChanged(sender As Object, e As EventArgs) Handles cmboDevice.SelectedValueChanged
        Dim DmxAddress As Integer
        Dim s = "channel 1"
        Dim l As New DataTable


        If cmboDevice.Text = "" Then Exit Sub
        Dim row As DataRowView = cmboDevice.SelectedItem
        Dim fieldvalue As String = row(0).ToString

        l = Await cls.GetDeviceConfiguration("name", fieldvalue)

        ' query datatable for distinct channel # and name
        l.Columns.Remove("dmxvalue")
        l.Columns.Remove("dmxfunction")

        Dim t As DataTable = l.DefaultView.ToTable(True, "startchannel", "dmxfunctioncategory")

        ' Loop and display.
        For x = 1 To 18
            'TabControl1.TabPages("TabPage5").Controls("lblch" & x).Text = ""
            Panel3.Controls("lblch" & x).Text = ""
        Next

        For Each r As DataRowView In t.DefaultView
            'TabControl1.TabPages("TabPage5").Controls("lblch" & r(0)).Text = r(1)
            Panel3.Controls("lblch" & r(0)).Text = r(1)
        Next

        ' store max channels of device
        txtChannelCount.Text = t.Rows.Count - 1

        'Select Case fieldvalue
        '    Case "Laser"
        '        lblCh1.Text = "test"
        '        DmxAddress = 1
        '    Case "Strobe"
        '        lblCh1.Text = "function laser 2"
        '        DmxAddress = 33
        '    Case Else
        '        lblCh1.Text = ""
        'End Select

        device = DmxAddress - 1

    End Sub

    Private Sub dgScenes_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Try
            SetupDgScenes()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    Private Async Sub btnAddScene_Click(sender As Object, e As EventArgs) Handles btnAddScene.Click
        Dim r As String
        Dim l As DataTable

        r = Await cls.DoesRecordExist("scenes", "name", "NEW")

        If r = 0 Then
            ' create new record
            r = Await cls.AddNewSceneRecord("NEW")
        Else
            Using New Centered_MessageBox(Me)
                MessageBox.Show("A NEW record already exists. Use this first")
            End Using
        End If

        l = Await cls.GetScenes()

        dgScenes.DataSource = l
        dgScenes.Refresh()
    End Sub

    Private Sub btnLasers1_Click(sender As Object, e As EventArgs) Handles btnLasers1.Click
        Timer2.Stop()
        dmxdata(1) = 255
        dmxdata(2) = 9
        dmxdata(3) = 36
        dmxdata(4) = 0
        dmxdata(5) = 73
        dmxdata(6) = 0
        dmxdata(7) = 0
        dmxdata(8) = 0
        dmxdata(9) = 0
        dmxdata(10) = 13
        dmxdata(11) = 0
        dmxdata(12) = 0
        dmxdata(13) = 133
        dmxdata(14) = 32
        dmxdata(15) = 0
        dmxdata(16) = 0
        dmxdata(17) = 0
        dmxdata(18) = 0

        dmxdata(33) = 255
        dmxdata(34) = 163
        dmxdata(35) = 28
        dmxdata(36) = 0
        dmxdata(37) = 73
        dmxdata(38) = 0
        dmxdata(39) = 0
        dmxdata(40) = 0
        dmxdata(41) = 0
        dmxdata(42) = 13
        dmxdata(43) = 0
        dmxdata(44) = 0
        dmxdata(45) = 129
        dmxdata(46) = 34
        dmxdata(47) = 0
        dmxdata(48) = 0
        dmxdata(49) = 0
        dmxdata(50) = 0

        MainClass.sendDMXdata(dmxdata)
    End Sub

    Private Sub btnLasers2_Click(sender As Object, e As EventArgs) Handles btnLasers2.Click
        dmxdata(1) = 255
        dmxdata(2) = 15
        dmxdata(3) = 32
        dmxdata(4) = 0
        dmxdata(5) = 57
        dmxdata(6) = 0
        dmxdata(7) = 0
        dmxdata(8) = 0
        dmxdata(9) = 0
        dmxdata(10) = 0
        dmxdata(11) = 0
        dmxdata(12) = 147
        dmxdata(13) = 0
        dmxdata(14) = 31
        dmxdata(15) = 0
        dmxdata(16) = 0
        dmxdata(17) = 0
        dmxdata(18) = 0

        dmxdata(33) = 255
        dmxdata(34) = 79
        dmxdata(35) = 232
        dmxdata(36) = 0
        dmxdata(37) = 57
        dmxdata(38) = 0
        dmxdata(39) = 0
        dmxdata(40) = 0
        dmxdata(41) = 0
        dmxdata(42) = 0
        dmxdata(43) = 129
        dmxdata(44) = 147
        dmxdata(45) = 0
        dmxdata(46) = 31
        dmxdata(47) = 0
        dmxdata(48) = 0
        dmxdata(49) = 0
        dmxdata(50) = 0

        MainClass.sendDMXdata(dmxdata)
    End Sub

    Private Sub btnLasers3_Click(sender As Object, e As EventArgs) Handles btnLasers3.Click
        dmxdata(1) = 255
        dmxdata(2) = 11
        dmxdata(3) = 32
        dmxdata(4) = 0
        dmxdata(5) = 87
        dmxdata(6) = 0
        dmxdata(7) = 0
        dmxdata(8) = 0
        dmxdata(9) = 0
        dmxdata(10) = 0
        dmxdata(11) = 0
        dmxdata(12) = 0
        dmxdata(13) = 0
        dmxdata(14) = 160
        dmxdata(15) = 131
        dmxdata(16) = 0
        dmxdata(17) = 0
        dmxdata(18) = 0

        dmxdata(33) = 255
        dmxdata(34) = 0
        dmxdata(35) = 25
        dmxdata(36) = 0
        dmxdata(37) = 87
        dmxdata(38) = 0
        dmxdata(39) = 0
        dmxdata(40) = 0
        dmxdata(41) = 0
        dmxdata(42) = 0
        dmxdata(43) = 0
        dmxdata(44) = 0
        dmxdata(45) = 0
        dmxdata(46) = 160
        dmxdata(47) = 131
        dmxdata(48) = 0
        dmxdata(49) = 0
        dmxdata(50) = 0

        MainClass.sendDMXdata(dmxdata)
    End Sub

    Private Sub btnLasers4_Click(sender As Object, e As EventArgs) Handles btnLasers4.Click
        dmxdata(1) = 255
        dmxdata(2) = 180
        dmxdata(3) = 30
        dmxdata(4) = 0
        dmxdata(5) = 56
        dmxdata(6) = 0
        dmxdata(7) = 0
        dmxdata(8) = 0
        dmxdata(9) = 0
        dmxdata(10) = 0
        dmxdata(11) = 0
        dmxdata(12) = 0
        dmxdata(13) = 0
        dmxdata(14) = 0
        dmxdata(15) = 1
        dmxdata(16) = 0
        dmxdata(17) = 0
        dmxdata(18) = 0

        dmxdata(33) = 255
        dmxdata(34) = 168
        dmxdata(35) = 33
        dmxdata(36) = 0
        dmxdata(37) = 56
        dmxdata(38) = 0
        dmxdata(39) = 0
        dmxdata(40) = 0
        dmxdata(41) = 0
        dmxdata(42) = 0
        dmxdata(43) = 0
        dmxdata(44) = 0
        dmxdata(45) = 0
        dmxdata(46) = 0
        dmxdata(47) = 1
        dmxdata(48) = 0
        dmxdata(49) = 0
        dmxdata(50) = 0

        MainClass.sendDMXdata(dmxdata)
    End Sub

    Private Sub btn_LasersOFF_Click(sender As Object, e As EventArgs) Handles btn_LasersOFF.Click
        dmxdata(1) = 0
        dmxdata(2) = 0
        dmxdata(3) = 0
        dmxdata(4) = 0
        dmxdata(5) = 0
        dmxdata(6) = 0
        dmxdata(7) = 0
        dmxdata(8) = 0
        dmxdata(9) = 0
        dmxdata(10) = 0
        dmxdata(11) = 0
        dmxdata(12) = 0
        dmxdata(13) = 0
        dmxdata(14) = 0
        dmxdata(15) = 0
        dmxdata(16) = 0
        dmxdata(17) = 0
        dmxdata(18) = 0

        dmxdata(33) = 0
        dmxdata(34) = 0
        dmxdata(35) = 0
        dmxdata(36) = 0
        dmxdata(37) = 0
        dmxdata(38) = 0
        dmxdata(39) = 0
        dmxdata(40) = 0
        dmxdata(41) = 0
        dmxdata(42) = 0
        dmxdata(43) = 0
        dmxdata(44) = 0
        dmxdata(45) = 0
        dmxdata(46) = 0
        dmxdata(47) = 0
        dmxdata(48) = 0
        dmxdata(49) = 0
        dmxdata(50) = 0

        MainClass.sendDMXdata(dmxdata)
    End Sub

    Private Sub btnLasers5_Click(sender As Object, e As EventArgs) Handles btnLasers5.Click
        dmxdata(6) = 0
        dmxdata(38) = 0

        MainClass.sendDMXdata(dmxdata)
    End Sub

    Private Sub txtCh1_MouseClick(sender As Object, e As MouseEventArgs) Handles txtCh1.MouseClick
        txtCh1.SelectAll()
    End Sub

    Private Sub txtCh2_MouseClick(sender As Object, e As MouseEventArgs) Handles txtCh2.MouseClick
        txtCh2.SelectAll()
    End Sub

    Private Sub txtCh3_MouseClick(sender As Object, e As MouseEventArgs) Handles txtCh3.MouseClick
        txtCh3.SelectAll()
    End Sub

    Private Sub txtCh4_MouseClick(sender As Object, e As MouseEventArgs) Handles txtCh4.MouseClick
        txtCh4.SelectAll()
    End Sub

    Private Sub txtCh5_MouseClick(sender As Object, e As MouseEventArgs) Handles txtCh5.MouseClick
        txtCh5.SelectAll()
    End Sub

    Private Sub txtCh6_MouseClick(sender As Object, e As MouseEventArgs) Handles txtCh6.MouseClick
        txtCh6.SelectAll()
    End Sub

    Private Sub txtCh7_MouseClick(sender As Object, e As MouseEventArgs) Handles txtCh7.MouseClick
        txtCh7.SelectAll()
    End Sub

    Private Sub txtCh8_MouseClick(sender As Object, e As MouseEventArgs) Handles txtCh8.MouseClick
        txtCh8.SelectAll()
    End Sub

    Private Sub txtCh9_MouseClick(sender As Object, e As MouseEventArgs) Handles txtCh9.MouseClick
        txtCh9.SelectAll()
    End Sub

    Private Sub txtCh10_MouseClick(sender As Object, e As MouseEventArgs) Handles txtCh10.MouseClick
        txtCh10.SelectAll()
    End Sub

    Private Sub txtCh11_MouseClick(sender As Object, e As MouseEventArgs) Handles txtCh11.MouseClick
        txtCh11.SelectAll()
    End Sub

    Private Sub txtCh12_MouseClick(sender As Object, e As MouseEventArgs) Handles txtCh12.MouseClick
        txtCh12.SelectAll()
    End Sub

    Private Sub txtCh13_MouseClick(sender As Object, e As MouseEventArgs) Handles txtCh13.MouseClick
        txtCh13.SelectAll()
    End Sub

    Private Sub txtCh14_MouseClick(sender As Object, e As MouseEventArgs) Handles txtCh14.MouseClick
        txtCh14.SelectAll()
    End Sub

    Private Sub txtCh15_MouseClick(sender As Object, e As MouseEventArgs) Handles txtCh15.MouseClick
        txtCh15.SelectAll()
    End Sub

    Private Sub txtCh16_MouseClick(sender As Object, e As MouseEventArgs) Handles txtCh16.MouseClick
        txtCh16.SelectAll()
    End Sub

    Private Sub txtCh17_MouseClick(sender As Object, e As MouseEventArgs) Handles txtCh17.MouseClick
        txtCh17.SelectAll()
    End Sub

    Private Sub txtCh18_MouseClick(sender As Object, e As MouseEventArgs) Handles txtCh18.MouseClick
        txtCh18.SelectAll()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        dmxdata(6) = 255
        dmxdata(38) = 255

        MainClass.sendDMXdata(dmxdata)
    End Sub

    Private Sub btnLasers5_MouseHover(sender As Object, e As EventArgs)
        dmxdata(6) = 0
        dmxdata(38) = 0

        MainClass.sendDMXdata(dmxdata)
    End Sub

    Private Sub Button2_MouseHover(sender As Object, e As EventArgs) Handles Button2.MouseHover
        dmxdata(6) = 255
        dmxdata(38) = 255

        MainClass.sendDMXdata(dmxdata)
    End Sub

    Private Sub Button2_MouseLeave(sender As Object, e As EventArgs) Handles Button2.MouseLeave
        dmxdata(6) = 0
        dmxdata(38) = 0

        MainClass.sendDMXdata(dmxdata)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        dmxdata(6) = 255
        dmxdata(38) = 255

        MainClass.sendDMXdata(dmxdata)
    End Sub

    Async Sub btnSaveSlice_Click(sender As Object, e As EventArgs) Handles btnSaveSlice.Click
        Dim keyname As String = cmboScenes.SelectedText
        Dim r As Integer

        If keyname = "" Then Exit Sub
        r = Await cls.UpdateScene(keyname, dmxdata)

    End Sub

    Private Sub dgChannels_DataError(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgChannels.DataError


        MessageBox.Show("Error:  " & e.Context.ToString())
        'e.Cancel = True

        If (e.Context = DataGridViewDataErrorContexts.Commit) _
            Then
            MessageBox.Show("Commit error")
        End If
        If (e.Context = DataGridViewDataErrorContexts.CurrentCellChange) Then
            MessageBox.Show("Cell change")
        End If
        If (e.Context = DataGridViewDataErrorContexts.Parsing) Then
            MessageBox.Show("parsing error")
        End If
        If (e.Context = DataGridViewDataErrorContexts.LeaveControl) Then
            MessageBox.Show("leave control error")
        End If

        If (TypeOf (e.Exception) Is ConstraintException) Then
            Dim view As DataGridView = CType(sender, DataGridView)
            view.Rows(e.RowIndex).ErrorText = "an error"
            view.Rows(e.RowIndex).Cells(e.ColumnIndex) _
                .ErrorText = "an error"
            MsgBox("error")
            e.ThrowException = False

        End If

    End Sub


    Private Async Sub btn_DeleteConfig_Click(sender As Object, e As EventArgs) Handles btn_DeleteConfig.Click
        Dim i = dgFixtures.CurrentRow.Index
        Dim j = dgChannels.CurrentRow.Index
        Dim keyvalue = dgFixtures.Item(0, i).Value
        Dim l As DataTable
        Dim r As Integer

        Using New Centered_MessageBox(Me)
            If MessageBox.Show("Are you sure you want to delete selected configuration rows?", "Delete Confirmation", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                For Each row As DataGridViewRow In dgChannels.SelectedRows
                    r = Await cls.DeleteConfigRecord(keyvalue, j)
                Next

                l = Await cls.GetDeviceConfiguration("_id", keyvalue)

                dgChannels.DataSource = l
                dgChannels.Refresh()

            Else
                Exit Sub
            End If
        End Using
    End Sub

    Private Async Sub dgChannels_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgChannels.CellEndEdit
        Dim keyname As String = dgFixtures.SelectedRows(0).Cells(0).Value
        Dim dgindex As Integer = e.RowIndex
        Dim fieldname As String = ""
        Dim fieldvalue As String = ""
        Dim r As Integer

        ' null check
        If Not IsDBNull(dgChannels.Columns(e.ColumnIndex).DataPropertyName) Then
            fieldname = dgChannels.Columns(e.ColumnIndex).DataPropertyName
        End If
        If Not IsDBNull(dgChannels.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) Then
            fieldvalue = dgChannels.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
        End If

        ' update record
        r = Await cls.UpdateConfigRecord(keyname, dgindex, fieldname, fieldvalue)

    End Sub

    Private Async Sub dgavailabledevices_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgavailabledevices.CellContentClick
        Dim senderGrid = DirectCast(sender, DataGridView)

        If TypeOf senderGrid.Columns(e.ColumnIndex) Is DataGridViewButtonColumn AndAlso
           e.RowIndex >= 0 Then
            'TODO - Button Clicked - Execute Code Here
            Dim row As Integer = e.RowIndex ' 0 based

            ' validate current layout name
            Dim drv As DataRowView = cmboLayouts.SelectedItem
            Dim s As String = drv(5)
            Dim t As String = drv(1)

            If s = "" Then
                Using New Centered_MessageBox(Me)
                    MessageBox.Show("Please provide a new layout name or load existing layout first")
                End Using
                Exit Sub
            End If

            ' edit the layout collection to add device
            Dim keyvalue = dgavailabledevices.Item(2, row).Value
            Dim r As Integer

            r = Await cls.AddDeviceToLayout(keyvalue, s)

            ' reload layout datagrid
            Await populateDG4(t)

            ' decrement available device quantity by 1 for specific layout

        End If
    End Sub

    Private Sub btnNewLayout_Click(sender As Object, e As EventArgs) Handles btnNewLayout.Click
        Dim frm As New frmNewObject("layout")
        frm.Show()
    End Sub

    Private Sub cmboLayouts_SelectedValueChanged(sender As Object, e As EventArgs) Handles cmboLayouts.SelectedValueChanged
        Dim t As Task


        'load dmx fixture library
        If cmboLayouts.Text <> "" Then
            Dim drv As DataRowView = cmboLayouts.SelectedItem
            Dim s As String = drv(1)
            t = populateDG4(s)
            txtLayoutName.Text = s
            My.Settings.defaultLayout = s
            My.Settings.Save()
        End If


        'd1 = Await cls.GetLayouts()
        'dgshowlayouts.DataSource = d1

        'Dim col As DataColumnCollection = d1.Columns
        'If col.Contains("_id") Then
        '    dgshowlayouts.Columns("_id").Visible = False
        'End If

    End Sub

    Private Async Function populateDG4(s As String) As Task
        Dim l As DataTable

        dgshowlayouts.Columns.Clear()
        dgshowlayouts.DataSource = Nothing
        SetupDgshowlayouts()

        l = Await cls.GetLayoutConfiguration("name", s)

        dgshowlayouts.DataSource = l
        dgshowlayouts.Refresh()

        If txtLayoutName.Text <> "" Then
            cmboLayouts.SelectedItem = s
        End If

    End Function

    Private Sub TabPage7_Click(sender As Object, e As EventArgs) Handles TabPage7.Click

    End Sub

    Private Async Sub btnDelDeviceFromLayout_Click(sender As Object, e As EventArgs) Handles btnDelDeviceFromLayout.Click

        Dim l As Task
        Dim r As Integer

        ' error checking
        If cmboLayouts.Text = "" Then Exit Sub

        Dim drv As DataRowView = cmboLayouts.SelectedItem
        Dim s As String = drv(5)
        Dim t As String = drv(1)

        Using New Centered_MessageBox(Me)
            If MessageBox.Show("Are you sure you want to delete selected device rows?", "Delete Confirmation", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                For Each row As DataGridViewRow In dgshowlayouts.SelectedRows
                    r = Await cls.DeleteLayoutDevice(s, row.Index)
                Next

                l = populateDG4(t)

            Else
                Exit Sub
            End If
        End Using
    End Sub

    Private Async Sub dgshowlayouts_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgshowlayouts.CellEndEdit
        Dim dgindex As Integer = e.RowIndex
        Dim fieldname As String = ""
        Dim fieldvalue As String = ""
        Dim r As Integer


        ' error checking
        If cmboLayouts.Text = "" Then Exit Sub

        Dim drv As DataRowView = cmboLayouts.SelectedItem
        Dim s As String = drv(5)
        Dim t As String = drv(1)

        ' null check
        If Not IsDBNull(dgshowlayouts.Columns(e.ColumnIndex).DataPropertyName) Then
            fieldname = dgshowlayouts.Columns(e.ColumnIndex).DataPropertyName
        End If
        If Not IsDBNull(dgshowlayouts.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) Then
            fieldvalue = dgshowlayouts.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
        End If

        ' update record
        r = Await cls.UpdateLayoutConfigRecord(s, dgindex, fieldname, fieldvalue)
    End Sub

    Private Sub cmboSliceSelect_SelectedValueChanged(sender As Object, e As EventArgs) Handles cmboSliceSelect.SelectedValueChanged
        ' clear slice name if existing slice selected to avoid user confusion
        txtSliceName.Text = ""
    End Sub

    Private Async Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim slicename As String
        Dim sliceid As String
        Dim duration As String
        Dim fade As String
        Dim nextslice As String
        Dim r As String

        ' user validation
        If cmboSliceSelect.Text = "" And txtSliceName.Text = "" Then
            Using New Centered_MessageBox(Me)
                MessageBox.Show("Please select an existing slice or enter a new slice name")
            End Using
            Exit Sub
        End If

        ' get values
        duration = cmboSliceDuration.SelectedValue
        fade = cmboSliceFade.Text
        nextslice = cmboNextSlice.Text

        ' get slice name
        If cmboSliceSelect.Text <> "" Then
            ' update existing slice
            Dim drv As DataRowView = cmboSliceSelect.SelectedItem
            slicename = drv(1)
            sliceid = drv(0)
        Else
            ' save new slice
            slicename = txtSliceName.Text
        End If

        r = Await cls.DoesRecordExist("slices", "name", slicename)


        If r = "000000000000000000000000" Then
            ' create new record
            r = Await cls.saveNewSlice(txtSliceName.Text, duration, fade, nextslice, dmxdata)
            Await loadSliceCombo()
        Else
            r = Await cls.updateSlice(r, slicename, duration, fade, nextslice, dmxdata)
        End If

    End Sub

    Private Async Function loadSliceCombo() As Task
        Dim t As DataTable
        t = Await cls.GetSlices()
        cmboSliceSelect.DataSource = t
        cmboSliceSelect.ValueMember = "_id"
        cmboSliceSelect.DisplayMember = "name"
        cmboSliceSelect.Text = ""
    End Function



    Private Async Sub btnRunSlice_Click(sender As Object, e As EventArgs) Handles btnRunSlice.Click
        Dim d As DataTable
        Dim a As Integer

        ' load slice info
        Dim r1 As DataRowView = cmboSliceSelect.SelectedItem
        Dim slicename As String = r1(2)
        Dim sliceid As String = r1(0)
        Dim channelcount As String = txtChannelCount.Text

        ' load current selected device info
        Dim r2 As DataRowView = cmboDevice.SelectedItem
        Dim startchannel As String = r2(2)


        d = Await cls.GetSlice("_id", sliceid)

        Dim name As String = d.Rows(0)("name")
        Dim id As String = d.Rows(0)("_id")
        Dim duration As String = d.Rows(0)("duration")
        Dim fade As String = d.Rows(0)("fade")
        Dim dmxstring As String = d.Rows(0)("dmxstring")
        Dim nextslice As String = d.Rows(0)("nextslice")

        ' decode dmxstring back to dmx byte array
        dmxdata = System.Text.Encoding.Default.GetBytes(dmxstring)

        ' set mixer board to match slice values for selected device
        For a = 0 To channelcount - 1
            Dim pb As TrackBar = Controls.Find("Ch" & a + 1, True).FirstOrDefault()

            If dmxdata(startchannel + a) = Nothing Then
                pb.Value = 0
            Else
                pb.Value = dmxdata(startchannel + a)
            End If

            Dim tc As TextBox = Controls.Find("txtCh" & a + 1, True).FirstOrDefault()
            tc.Text = pb.Value

        Next

        ' send dmx
        MainClass.sendDMXdata(dmxdata, 0)

        ' populate mixer per device
    End Sub

    Private Sub btnDeviceOff_Click(sender As Object, e As EventArgs) Handles btnDeviceOff.Click
        ResetSliders()
    End Sub

    Private Sub txtSliceName_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtSliceName.KeyPress
        cmboSliceSelect.SelectedValue = ""
    End Sub

    Async Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        Dim d1, d2 As New DataTable

        If TabControl1.SelectedTab Is TabPage2 Then

        End If


        'fixtures tab
        If TabControl1.SelectedTab Is TabPage4 Then
            'load dmx fixture library
            d1 = Await cls.GetDevices()
            dgFixtures.DataSource = d1
            dgFixtures.Columns("_id").Visible = False
            If dgFixtures.RowCount > 0 Then
                dgFixtures.Rows(0).Selected = True
                Dim s As String = dgFixtures.Item(0, 0).Value
                Await populateDG2(s)
            End If
        End If

        ' board/mixer tab 
        If TabControl1.SelectedTab Is TabPage5 Then

            ' validation
            If txtLayoutName.Text = "" Then
                Using New Centered_MessageBox(Me)
                    MessageBox.Show("Please select a device layout from the layouts tab")
                End Using
                Exit Sub
            End If

            'load scene library
            d1 = Await cls.GetScenes()
            cmboScenes.DataSource = d1
            cmboScenes.ValueMember = "name"
            cmboScenes.DisplayMember = "name"

            ' load fixtures
            If txtLayoutName.Text = "" Then
                ' error validation
            Else
                d2 = Await cls.GetLayoutConfiguration("name", txtLayoutName.Text)
                cmboDevice.DataSource = d2
                cmboDevice.ValueMember = "devicealias"
                cmboDevice.DisplayMember = "deviceref"
            End If

            ' populate slice panel controls
            Dim t As DataTable
            t = cls.LoadSliceCombos("duration")
            cmboSliceDuration.DataSource = t
            cmboSliceDuration.ValueMember = "item"
            cmboSliceDuration.DisplayMember = "alias"
            cmboSliceDuration.Text = ""

            t = cls.LoadSliceCombos("fade")
            cmboSliceFade.DataSource = t
            cmboSliceFade.ValueMember = "item"
            cmboSliceFade.DisplayMember = "alias"
            cmboSliceFade.Text = ""

            Await loadSliceCombo()
        End If

        ' scenes tab
        If TabControl1.SelectedTab Is TabPage6 Then
            'load scene library
            d1 = Await cls.GetScenes()
            dgScenes.DataSource = d1

            cmboScenes.DataSource = d1
            cmboScenes.ValueMember = "name"
            cmboScenes.DisplayMember = "name"
        End If

        ' layouts tab
        If TabControl1.SelectedTab Is TabPage7 Then

            'load layout combo
            Await LoadLocation()

            'load dmx fixture library
            d1 = Await cls.GetDevices()
            dgavailabledevices.DataSource = d1
            dgavailabledevices.Columns("_id").Visible = False
            If dgavailabledevices.RowCount > 0 Then
                dgavailabledevices.Rows(0).Selected = True
            End If
            ' populate drop down layout combo control

        End If
    End Sub

    Private Async Sub btnDeleteCalls_Click(sender As Object, e As EventArgs) Handles btnDeleteCalls.Click
        'Dim i = dgSliceCalls.CurrentRow.Index
        Dim keyvalue As String '= dgSliceCalls.Item(0, i).Value
        Dim r As Integer

        Using New Centered_MessageBox(Me)
            If MessageBox.Show("Are you sure you want to delete selected slice calls?", "Delete Confirmation", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                For Each row As DataGridViewRow In dgSliceCalls.SelectedRows
                    keyvalue = dgSliceCalls.Rows(row.Index).Cells(0).Value
                    r = Await cls.DeleteSliceCall(keyvalue)
                Next

                Await populateDG1(txtContentID.Text)

            Else
                Exit Sub
            End If
        End Using
    End Sub

    Private Async Sub dgSliceCalls_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgSliceCalls.CellEndEdit
        Dim keyvalue As String = dgSliceCalls.Rows(e.RowIndex).Cells(0).Value
        Dim dgindex As Integer = e.RowIndex
        Dim fieldname As String = ""
        Dim fieldvalue As String = ""
        Dim r As Integer

        ' null check
        If Not IsDBNull(dgSliceCalls.Columns(e.ColumnIndex).DataPropertyName) Then
            fieldname = dgSliceCalls.Columns(e.ColumnIndex).DataPropertyName
        End If
        If Not IsDBNull(dgSliceCalls.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) Then
            fieldvalue = dgSliceCalls.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
        End If

        ' update record
        r = Await cls.UpdateSliceCallRecord(keyvalue, dgindex, fieldname, fieldvalue)
    End Sub

    Private Function ResetSliders()
        Dim r As DataRowView = cmboDevice.SelectedItem
        Dim startchannel As String = r(2)
        Dim channelcount As String = txtChannelCount.Text
        Dim a As Integer

        For a = 1 To channelcount
            Dim pb As TrackBar = Controls.Find("Ch" & a, True).FirstOrDefault()
            pb.Value = 0

            Dim tc As TextBox = Controls.Find("txtCh" & a, True).FirstOrDefault()
            tc.Text = ""
            tc.Text = "0"

        Next
    End Function


    Private Sub btnAllDevicesOff_Click(sender As Object, e As EventArgs) Handles btnAllDevicesOff.Click
        ResetSliders()

        For x = 0 To UBound(dmxdata)
            dmxdata(x) = 0
        Next
    End Sub
End Class
