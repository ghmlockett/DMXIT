
Option Explicit On
Imports System.Threading
Imports System
Imports System.Timers


Public Class USBDMXPro
    Public deviceInitialized As Boolean ' DMX device is initialized
    Public DeviceType = "DMXUSBpro" ' DMX interface
    Public PortNumber As Integer = 3 ' port DMX interface is connected to
    Public Direction As Integer ' crossfader driection
    Public currTime As Long ' in milliseconds
    Public startTime As Single ' realtime clock
    Public fadeLength As Long ' fade time in seconds
    Public prevScene As String
    Public blackout As Boolean
    Public numChannels As Integer = 8 ' number of fader channels
    Public numSubmasters As Integer ' number of submasters
    Public submasterScene(300) As Double ' store the scene number the submaster is associated with
    Public submasterLabel(300) As String
    Public submasterLevel(300) As Integer ' the current level of this submaster
    Public submasterInd(300) As Integer ' independent submaster flag
    Public submasterLTP(300) As Integer ' LTP/HTP submaster flag
    Public submasterScenes(300, 512) As Integer ' store the scene levels for each submaster
    Public currSubTab As Integer ' which notebook tab  of submasters is being displayed
    Public Unsaved As Boolean ' changed show has not been save yet
    Public UnsavedPatch As Boolean ' changed patch has not been saved yet
    Public numChanChanged As Boolean
    Public patchdata(0 To 511) As Integer ' the softpatch
    Public dmxdata(0 To 511) As Byte  ' DMX output array (levels 0 to 255)
    Public dmxdataStore(0 To 511) As Byte ' DMX output array (levels 0 to 255)
    Public dmxdataNew(0 To 511) As Byte ' DMX output array (levels 0 to 255)
    Public currentShowFile As String
    Public currentPatchName As String
    Public currentPatchFile As String
    Public hackCount As Integer

    Private Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)

    Dim Firmware_VersionL As Byte
    Dim Firmware_VersionH As Byte  '
    Dim serial As String
    Dim BreakTime As Byte  ' 9..127 in 10.67us units
    Dim MarkAfterBreak As Byte  ' 1..127 in 10.67us units
    Dim FrameRate As Byte  ' frames sent per second, 1..40
    Const MaxPorts = 16 ' max number of serial ports to handle

    Dim packetbuffer As String
    ' eg after SearchPorts, if com_ports(4) = true this com port 4 actually exists
    Dim Com_Ports(0 To MaxPorts - 1) As Boolean
    Dim Widgets_Found(0 To MaxPorts - 1) As Boolean
    ' misc packet constants - don't work as cost for some reason
    'Const SOM = &H7E ' Start Of Message Marker
    'Const EOM = &HE7 ' End Of Message Marker
    Public Const GET_WIDGET_CFG_REPLY = 3 ' returned config
    Public Const DMX_INPUT = 5  ' auto sent by widget upon arrival of a dmx packet
    Dim SOM As String
    Dim EOM As String

    Public MSComm1 As New MSCommLib.MSComm

    'Private WithEvents Timer1 As New System.Windows.Forms.Timer()
    Private WithEvents Timer1 As New Timers.Timer()

    Private fade_start_time As Date
    Private fade_stop_time As Date
    Private fade_elapsed_time As TimeSpan

    Public Sub Main()
        startDMX()
    End Sub

    Public Function runDMXfade()
        Dim fadeInterval As Long


        ' set fadeLength before calling
        ' defaults to 10 seconds if not set
        If fadeLength = 0 Then
            fadeLength = 10
        Else
            fadeLength = fadeLength - 2
        End If

        ' stop timer1
        Timer1.Stop()

        fadeInterval = 300
        fade_elapsed_time = Nothing

        fadeLength = fadeLength / (fadeInterval / 1000)


        ' save new value reference
        Array.Copy(dmxdata, dmxdataNew, dmxdata.Length)

        ' start timer
        'Timer1 = New System.Timers.Timer(fadeInterval)
        Timer1.Interval = fadeInterval
        Timer1.Start()

        fade_start_time = Now

    End Function

    Private Sub TimerEventProcessor(myObject As Object,
                                           ByVal myEventArgs As EventArgs)
        Dim elapsedSeconds As Decimal = 0.01
        Dim remainingSeconds As Decimal
        Dim currPos As Decimal

        fade_elapsed_time = Now - fade_start_time

        elapsedSeconds = fade_elapsed_time.TotalSeconds
        remainingSeconds = fadeLength - elapsedSeconds

        If elapsedSeconds > fadeLength Then
            ' expired - reset
            Timer1.Stop()
        Else
            ' in flight - process
            If fadeLength = 0 Then
                currPos = 1
            Else
                currPos = elapsedSeconds / fadeLength ' fraction
            End If

            calcFade(currPos)

            ' send dmx
            sendDMXpro()

        End If


    End Sub

    Public Function calcFade(currPos As Decimal)
        Dim i, j As Integer
        Dim chanDif As Long
        Dim chanStepValue As Long


        If currPos = 0 Then
            currPos = 0.05
        End If

        For i = 0 To numChannels - 1
            j = i + 1
            chanDif = CInt(dmxdataStore(j)) - dmxdataNew(j)
            chanStepValue = chanDif * currPos

            Select Case chanStepValue
                Case > 0

                    If dmxdataStore(j) - chanStepValue > 255 Then
                        dmxdata(j) = 255
                    Else
                        dmxdata(j) = dmxdataStore(j) - chanStepValue
                    End If

                Case < 0

                    chanStepValue = chanStepValue * -1

                    If dmxdataStore(j) + chanStepValue < 0 Then
                        dmxdata(j) = 0
                    Else
                        dmxdata(j) = dmxdataStore(j) + chanStepValue
                    End If

            End Select

        Next i

        calcFade = 0

    End Function

    Public Sub sendDMXdata(a() As Byte, Optional FadeSeconds As Long = 0)

        ' copy array for history
        Array.Copy(a, dmxdata, a.Length)

        ' route to fade routine if specified
        If FadeSeconds > 0 Then
            fadeLength = FadeSeconds ' set class variable
            runDMXfade()
        Else
            runDMX()
        End If

    End Sub

    Public Sub runDMX()
        Timer1.Stop()
        sendDMXpro()
    End Sub

    Public Function initPro()
        ' initialize serial port connected to Enttec USB DMX Pro widget
        Dim msg As String

        MSComm1.CommPort = PortNumber

        On Error Resume Next
        MSComm1.PortOpen = True

        If MSComm1.PortOpen = False Then
            MSComm1.CommPort = PortNumber
            MSComm1.CommPort = False
        End If
        If MSComm1.PortOpen = False Then
            'MsgBox("Failed to open USB DMX Pro on comm port: " & Str(PortNumber) & " - Check Settings, Device")
            Return 1
            Exit Function
        Else
            'Call Send_CFG_Request_Packet ' see if a pro responds..
            'Sleep (100)
            'if we recieve a cfg reply then we know we have a pro widget
            'If frmMain.MSComm1.InBufferCount > 0 Then result = Recieve_Packet(frmMain.MSComm1.Input)

            'start the interface sending DMX msgnumber 6 and 1 byte of data
            msg = Chr(126) & Chr(6) & Chr(1) & Chr(0) & Chr(0) & Chr(231)
            MSComm1.Output = msg
            deviceInitialized = True
            'MsgBox("Output on Enttec USBDMX Pro port: " & PortNumber & " firmware: " & Firmware_VersionH & "." & Firmware_VersionL)
            Return 0

            ' Hook up the Elapsed event for the timer.
            AddHandler Timer1.Elapsed, AddressOf TimerEventProcessor

            updateLevels()
        End If

    End Function

    Sub updateLevels()
        ' update the DMX output after calculating current levels
        Dim i As Integer
        Dim chn As Integer
        Dim Ind As Integer
        Dim sm As Integer
        Dim smChan As Integer
        Dim tmpSMchan As Integer
        Dim smlevel As Single
        Dim master As Single
        Dim xfade As Single
        Dim afader As Single
        Dim bfader As Single
        Dim mix As Integer
        Dim override As Integer
        Dim chan As Single
        Dim percentLevel As Integer
        Dim DMXlevel As Integer
        Dim indChan As Boolean
        Dim dmx As Integer
        'frmDebug.Cls
        If blackout Then
            ' zero all outputs
            For i = 0 To numChannels - 1
                dmxdata(i + 1) = 0
                'frmOutputMonitor.update
            Next i
        Else
            'send out current levels
            '    master = (255 - frmRun.sliderGrandMaster.Value) / 255
            '    xfade = frmRun.sliderXfader.Value / 255

            '    For chn = 0 To numChannels - 1
            '        ' indicate which scene is active by changing background colors
            '        Ind = chn + numChannels
            '        ' display color highlight on active scene
            '        If frmRun.sliderXfader.Value < 5 Then ' display color highlight on active scene
            '            lblLevel(chn).BackColor = vbYellow
            '            lblLevel(Ind).BackColor = vbWhite
            '        End If
            '        If frmRun.sliderXfader.Value > 250 Then
            '            lblLevel(chn).BackColor = vbWhite
            '            lblLevel(Ind).BackColor = vbYellow
            '        End If
            '        ' calculate level based on sceneA, sceneB, crossfader, and master
            '        afader = (100 - SliderChan(chn).Value) * 2.55
            '        bfader = (100 - SliderChan(chn + numChannels).Value) * 2.55
            '        mix = Int((xfade * bfader) + ((1 - xfade) * afader)) ' calculate level crossfaded

            '        ' apply submasters
            '        indChan = False ' there is no independent submaster controlling this channel
            '        smChan = 0 ' the highest level for this channel, for all submasters
            '        For sm = 0 To numSubmasters - 1
            '            smlevel = (submasterLevel(sm)) / 255 ' submaster level multiplyer
            '            tmpSMchan = smlevel * submasterScenes(sm, chn)
            '            If smChan < tmpSMchan Then smChan = tmpSMchan ' HTP
            '            ' apply independent submasters to override everything else
            '            If submasterInd(sm) = 1 Then ' if ind is true, override channel level
            '                If submasterScenes(sm, chn) > 0 Then ' level must be above zero to override
            '                    indChan = True
            '                    override = (smlevel * submasterScenes(sm, chn)) * 2.55
            '                End If
            '            End If
            '        Next sm

            '        smChan = smChan * 2.55 ' scale to dmx value
            '        'frmDebug.Print smChan;
            '        If indChan Then  '  override -  independent
            '            chan = override
            '        Else ' dont override, see if level governed by submaster is higher than mix
            '            If smChan > mix Then mix = smChan ' if submaster level is higher, take precedence
            '            chan = master * mix
            '        End If

            '        DMXlevel = Int(chan)

            '        ' apply soft patch
            '        ' patchdata index is the DMX channel, the value is the fader to which it is assigned
            '        ' search patchdata array for matching fader number, assign level to dmxdata if match
            '        For dmx = 1 To 512
            '            If patchdata(dmx) = (chn + 1) Then dmxdata(dmx) = DMXlevel
            '        Next dmx
            '        'frmOutputMonitor.update
            '    Next chn
        End If

        'dmxdata(1) = 255
        ''dmxdata(3) = 255

        'sendDMXpro()

    End Sub

    Public Sub closeDMX()
        ' close the dmx device
        closePro
    End Sub

    Public Function startDMX()
        Dim a As Integer
        Dim status As Integer
        ' initialize the dmx device
        status = initPro()

        If status = 1 Then Return 1 Else Return 0

        ' clear the universe - or better, load first scene
        For a = 0 To 511
            dmxdata(a) = 0
        Next a

        ' clear history array
        For a = 0 To 511
            dmxdataStore(a) = 0
        Next a
        '    sendDMXdata ' don't blackout
    End Function

    Public Sub sendDMXpro()
        ' send DMX data to Enttec USB DMX Pro widget
        Dim msg As String
        Dim levels As String
        Dim startcode As String
        Dim dataLen As Integer
        Dim dataLenLSB As Integer
        Dim dataLenMSB As Integer
        Dim i As Integer
        startcode = Chr(0)
        'create msg string from public dmxdata array
        levels = ""
        'frmOutputMonitor.lblOut.Caption = ""


        For i = 1 To UBound(dmxdata) - 1
            levels = levels & Chr(dmxdata(i))
            ' display levels
            'OutputMonitor.lblOut.Caption = OutputMonitor.lblOut.Caption & " " & Str(dmxdata(i))
        Next
        dataLen = Len(levels) + 1 ' for startcode, total data length should = 512
        'OutputMonitor.lblOut.Caption = Str(dataLen) & ":" '& levels
        'calculate MSB and LSB for length of DMX data
        If dataLen >= 256 Then
            dataLenLSB = dataLen - 256
            dataLenMSB = 1
            If dataLenLSB >= 256 Then
                dataLenLSB = dataLenLSB - 256
                dataLenMSB = 2
            End If
        Else
            dataLenLSB = dataLen
            dataLenMSB = 0
        End If
        ' sending DMX msgnumber 6 and dataLen bytes of data
        msg = Chr(126) & Chr(6) & Chr(dataLenLSB) & Chr(dataLenMSB) & startcode & levels & Chr(231)
        If deviceInitialized = True Then
            Try
                MSComm1.Output = msg
            Catch ex As Exception
                Debug.Print(ex.Message)
            End Try
        End If

        ' copy array for history
        Array.Copy(dmxdata, dmxdataStore, dmxdata.Length)

        'If frmOutputMonitor.mnuDeviceOutput.Checked Then
        'frmOutputMonitor.lblOut.Caption = Str(dataLen) & ":" & hexStr(msg)  ' display hex message
        'End If
    End Sub

    Public Sub closePro()
        ' close serial port connected to Enttec USB DMX Pro widget
        If deviceInitialized = True Then
            MSComm1.PortOpen = False
        End If
    End Sub

End Class



