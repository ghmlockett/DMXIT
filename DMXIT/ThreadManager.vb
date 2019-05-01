Imports System.Timers
Imports System.Linq

Public Class ThreadManager

    Public buf As New DataTable
    Private cls As dataclass
    Private drv As New USBDMXPro

    Dim timer1 As Timer = New Timer(300)


    Public Function Initiate(clsvar1 As dataclass, clsvar2 As USBDMXPro)
        AddHandler timer1.Elapsed, New ElapsedEventHandler(AddressOf TimerElapsed)
        timer1.Start()

        ' initiate datatable buffer
        buf.Columns.Add("name")
        buf.Columns.Add("sliceid")
        buf.Columns.Add("duration", GetType(Long))
        buf.Columns.Add("lastduration", GetType(Long))
        buf.Columns.Add("fade")
        buf.Columns.Add("status")
        buf.Columns.Add("starttime", GetType(Long))
        buf.Columns.Add("nextslice")
        buf.Columns.Add("dmxstring")


        cls = clsvar1
        drv = clsvar2

    End Function

    Public Sub clrBuffer()
        buf.Clear()
    End Sub

    Public Sub runSliceChain(slicename As String)

        Dim r As DataRow = buf.NewRow()
        r("name") = slicename
        r("dmxstring") = ""
        r("duration") = 0
        r("lastduration") = 0
        r("fade") = ""
        r("starttime") = 0
        r("nextslice") = ""
        r("status") = "new"
        buf.Rows.Add(r)

    End Sub

    Public Async Sub TimerElapsed(ByVal sender As Object, ByVal e As ElapsedEventArgs)
        Dim x As Integer = 0
        Dim name As String
        Dim nextslice As String
        Dim duration As Integer
        Dim fade As String
        Dim sliceid As String
        Dim dmxstring As String
        Dim lastduration As String

        Dim starttime As Long = 0
        Dim status As String = ""
        Dim dmxdata(0 To 511) As Byte
        Dim d As DataTable = Nothing
        Dim t As New DataTable
        Dim stamp As String
        Dim ts As Long


        ' setup temp datatable
        t.Columns.Add("name")
        t.Columns.Add("sliceid")
        t.Columns.Add("duration", GetType(Long))
        t.Columns.Add("lastduration", GetType(Long))
        t.Columns.Add("fade")
        t.Columns.Add("starttime", GetType(Long))
        t.Columns.Add("dmxstring")
        t.Columns.Add("status")
        t.Columns.Add("nextslice")


        ' get timer start time stamp 
        timer1.Stop()
        stamp = Format(Now, "HH:mm:ss")
        ts = TimeSpan.ParseExact(stamp, {"hh\:mm\:ss"}, System.Globalization.CultureInfo.InvariantCulture).TotalMilliseconds


        ' STAGE
        ' read runBufferArray (slicename, status=new, duration, fade, dmxvalue, starttime)
        For Each row As DataRow In buf.Rows

            status = buf.Rows(x).Item("status")
            name = buf.Rows(x).Item("name")

            Select Case status
                Case "new"
                    ' look up any missing values
                    d = Await cls.GetSlice("name", name)

                    ' overwrite
                    duration = d.Rows(0)("duration")
                    buf.Rows(x).Item("duration") = duration

                    nextslice = d.Rows(0)("nextslice")
                    buf.Rows(x).Item("nextslice") = nextslice

                    sliceid = d.Rows(0)("_id")
                    buf.Rows(x).Item("sliceid") = sliceid

                    fade = d.Rows(0)("fade")
                    buf.Rows(x).Item("fade") = fade

                    dmxstring = d.Rows(0)("dmxstring")
                    buf.Rows(x).Item("dmxstring") = dmxstring

                    lastduration = duration
                    buf.Rows(x).Item("lastduration") = duration

                    ' write startime if time missing
                    If starttime = 0 Then buf.Rows(x).Item("starttime") = ts

                Case "pending"
                    ' load back values
                    name = row("name")
                    status = row("status")
                    starttime = row("starttime")
                    dmxstring = row("dmxstring")
                    fade = row("fade")
                    duration = row("duration")
                    lastduration = row("lastduration")
                    nextslice = row("nextslice")
                Case Else
                    Exit Sub
            End Select


            ' convert dmxstring back to dmxdata array
            dmxdata = System.Text.Encoding.Default.GetBytes(dmxstring) ' decode dmxstring back to dmx byte array

            ' send command if no duration
            Dim loadnextslice As Boolean = False
            If status = "new" Or duration = 0 Then
                drv.sendDMXdata(dmxdata, 0)
                buf.Rows(x).Item("status") = "delete" ' mark for delete
                loadnextslice = True
            Else
                If starttime + lastduration <= ts Then
                    drv.sendDMXdata(dmxdata, 0)
                    buf.Rows(x).Item("status") = "delete" ' mark for delete
                    loadnextslice = True
                End If
            End If

            ' get nextslice parameters from DB for slicename
            If loadnextslice = True And nextslice <> "" Then
                Dim r1 As DataRow = t.NewRow
                d = Await cls.GetSlice("name", nextslice)
                r1("name") = d.Rows(0)("name")
                r1("sliceid") = d.Rows(0)("_id")
                r1("duration") = d.Rows(0)("duration")
                r1("lastduration") = duration ' the calling slice value, not new slice!
                r1("fade") = d.Rows(0)("fade")
                r1("dmxstring") = d.Rows(0)("dmxstring")
                r1("starttime") = ts
                r1("status") = "pending"
                r1("nextslice") = d.Rows(0)("nextslice")
                t.Rows.Add(r1)
            End If

            ' row index trace
            x = x + 1
        Next row

        If x > 0 Then
            ' commit deleted records
            Dim rows() As DataRow
            rows = buf.Select("status='delete'")
            For Each row As DataRow In rows
                buf.Rows.Remove(row)
            Next

            ' add rows from datarowcollection
            For Each r As DataRow In t.Rows
                buf.ImportRow(r)
            Next
        End If

        timer1.Start()

    End Sub

End Class

