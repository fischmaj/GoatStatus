Imports System.IO
Public NotInheritable Class About
    Private Sub About_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim fname0 As String = "path.txt"
        Dim xVer As String = ""
        Try
            Using fn0 As StreamReader = New StreamReader(fname0)
                xVer = fn0.ReadLine()
                xVer = fn0.ReadLine()
                xVer = fn0.ReadLine()
                xVer = fn0.ReadLine()
                xVer = fn0.ReadLine()
                fn0.Close()
            End Using
        Catch ex As Exception
            MsgBox("Path File Error")
        End Try

        Me.LabelVersion.Text = "Version " + xVer
    End Sub
    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OKButton.Click
        Me.Close()
    End Sub
End Class
