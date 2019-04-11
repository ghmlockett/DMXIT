Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.IO

Public Class frmNewObject
    Dim objtype As String = ""
    Private cls As New dataclass

    Public Sub New(ByVal param1 As String)
        InitializeComponent()
        objtype = param1

        Me.Text = "NEW " & UCase(objtype)
        lblObject.Text = objtype
    End Sub

    Private Sub btnCreate_Click(sender As Object, e As EventArgs) Handles btnCreate.Click

        ' connect to mongo db
        If cls.InitiateDb() <> 0 Then
            MsgBox("MongoDB initialization error")
            Exit Sub
        End If

        If txtObject.Text = "" Then
            Using New Centered_MessageBox(Me)
                MessageBox.Show("Please enter a value in the text box")
            End Using
            Exit Sub
        End If

        CreateNewObject(objtype)

    End Sub

    Private Async Sub CreateNewObject(objtype As String)
        Dim r As Integer
        Dim l As DataTable
        Dim colname As String = ""


        Select Case objtype
            Case "layout"
                colname = "layouts"
            Case "scene"
                colname = "scenes"
        End Select


        r = Await cls.DoesRecordExist(colname, "name", txtObject.Text)

        If r = 0 Then
            ' create new record
            r = Await cls.AddNewLayoutRecord(txtObject.Text)
        Else
            Using New Centered_MessageBox(Me)
                MessageBox.Show("Layout name already exists.  Please enter new layout name")
            End Using
        End If

        Await frmMain.LoadLocation()

        ' close form
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        txtObject.Text = ""
        Me.Close()
    End Sub
End Class