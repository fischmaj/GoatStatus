Imports System.IO
Public Class News
    Dim lPath As String = ""
    Dim oPath As String = ""
    Dim bPath As String = ""
    Dim iPath As String = ""
    Dim nVer As String = ""
    Dim nNews As Integer
    Dim aNews(9, 3) As String
    Dim bolUpdate As Boolean = False
    Dim nUpdate As Integer = 0
    Dim nDate As Date
    Private Sub News_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim fname0 As String = "path.txt"
        Try
            Using fn0 As StreamReader = New StreamReader(fname0)
                lPath = fn0.ReadLine()
                oPath = fn0.ReadLine()
                bPath = fn0.ReadLine()
                iPath = fn0.ReadLine()
                nVer = fn0.ReadLine()
                nVer = fn0.ReadLine()
            End Using
        Catch ex As Exception
            MsgBox("Path File Error")
        End Try
        Me.lblDate.Text = Microsoft.VisualBasic.DateAndTime.Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now)
        Me.verLabel.Text = "Version " + nVer
        subGetNews()
        subLoadNews()
        tbNewTitle.Focus()
    End Sub
    Private Function fnFNBR(ByVal fn As String, ByVal path1 As String, ByVal path2 As String) As Boolean
        Try
            FileCopy(path1 & fn, path2 & fn)
            fnFNBR = True
        Catch ex As Exception
            fnFNBR = False
        End Try
    End Function
    Private Sub subGetNews()
        Dim fname0 As String = lPath + "news.txt"
        Try
            Using fn0 As StreamReader = New StreamReader(fname0)
                nNews = CDec(fn0.ReadLine)
                nDate = fn0.ReadLine()
                For x = 1 To nNews
                    aNews(x, 1) = fn0.ReadLine()
                    aNews(x, 2) = fn0.ReadLine()
                Next x
                fn0.Close()
            End Using
        Catch ex As Exception
            If fnFNBR("news.txt", bPath, lPath) Then
                subGetNews()
            Else
                MsgBox("admin:News File Read Error")
            End If
        End Try
    End Sub
    Private Sub subPutNews()
        Dim fname0 As String = lPath + "news.txt"
        Try
            Using fn0 As StreamWriter = New StreamWriter(fname0)
                fn0.WriteLine(nNews)
                fn0.WriteLine(nDate)
                For x = 1 To nNews
                    fn0.WriteLine(aNews(x, 1))
                    fn0.WriteLine(aNews(x, 2))
                Next x
                fn0.Close()
            End Using
        Catch ex As Exception
            If fnFNBR("news.txt", bPath, lPath) Then
                subPutNews()
            Else
                MsgBox("admin:News File Write Error")
            End If
        End Try
    End Sub
    Private Sub subLoadNews()
        If nNews >= 1 Then
            tbCN1.Text = aNews(1, 1)
            btnChg1.Enabled = True
            btnDel1.Enabled = True
        Else
            tbCN1.Text = "Empty"
            btnChg1.Enabled = False
            btnDel1.Enabled = False
        End If
        If nNews >= 2 Then
            tbCN2.Text = aNews(2, 1)
            btnChg2.Enabled = True
            btnDel2.Enabled = True
        Else
            tbCN2.Text = "Empty"
            btnChg2.Enabled = False
            btnDel2.Enabled = False
        End If
        If nNews >= 3 Then
            tbCN3.Text = aNews(3, 1)
            btnChg3.Enabled = True
            btnDel3.Enabled = True
        Else
            tbCN3.Text = "Empty"
            btnChg3.Enabled = False
            btnDel3.Enabled = False
        End If
        If nNews >= 4 Then
            tbCN4.Text = aNews(4, 1)
            btnChg4.Enabled = True
            btnDel4.Enabled = True
        Else
            tbCN4.Text = "Empty"
            btnChg4.Enabled = False
            btnDel4.Enabled = False
        End If
        If nNews >= 5 Then
            tbCN5.Text = aNews(5, 1)
            btnChg5.Enabled = True
            btnDel5.Enabled = True
        Else
            tbCN5.Text = "Empty"
            btnChg5.Enabled = False
            btnDel5.Enabled = False
        End If
        If nNews >= 6 Then
            tbCN6.Text = aNews(6, 1)
            btnChg6.Enabled = True
            btnDel6.Enabled = True
        Else
            tbCN6.Text = "Empty"
            btnChg6.Enabled = False
            btnDel6.Enabled = False
        End If
        If nNews >= 7 Then
            tbCN7.Text = aNews(7, 1)
            btnChg7.Enabled = True
            btnDel7.Enabled = True
        Else
            tbCN7.Text = "Empty"
            btnChg7.Enabled = False
            btnDel7.Enabled = False
        End If
        If nNews >= 8 Then
            tbCN8.Text = aNews(8, 1)
            btnChg8.Enabled = True
            btnDel8.Enabled = True
        Else
            tbCN8.Text = "Empty"
            btnChg8.Enabled = False
            btnDel8.Enabled = False
        End If
        If nNews >= 9 Then
            tbCN9.Text = aNews(9, 1)
            btnChg9.Enabled = True
            btnDel9.Enabled = True
        Else
            tbCN9.Text = "Empty"
            btnChg9.Enabled = False
            btnDel9.Enabled = False
        End If
        If nNews = 10 Then
            tbCN1.Text = aNews(10, 1)
            btnChg10.Enabled = True
            btnDel10.Enabled = True
        Else
            tbCN10.Text = "Empty"
            btnChg10.Enabled = False
            btnDel10.Enabled = False
        End If
    End Sub
    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        Dim nItem As Integer
        If Not bolUpdate Then
            nNews = nNews + 1
            nItem = nNews
        Else
            nItem = nUpdate
            bolUpdate = False
            btnClr.Visible = False
            btnAdd.Text = "Add News"
        End If
        aNews(nItem, 1) = tbNewTitle.Text
        aNews(nItem, 2) = tbNewItem.Text
        tbNewItem.Text = ""
        tbNewTitle.Text = ""
        subLoadNews()
    End Sub
    Private Sub subShiftNews()
        Dim x As Integer = 1
        Dim y As Integer
        While x < nNews
            If aNews(x, 1) = "" Then
                For y = x To nNews - 1
                    aNews(y, 1) = aNews(y + 1, 1)
                    aNews(y, 2) = aNews(y + 1, 2)
                Next y
            End If
            x = x + 1
        End While
    End Sub
    Private Sub btnDel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel1.Click
        aNews(1, 1) = ""
        aNews(1, 2) = ""
        subShiftNews()
        nNews = nNews - 1
        subLoadNews()
        subClr()
    End Sub
    Private Sub btnDel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel2.Click
        aNews(2, 1) = ""
        aNews(2, 2) = ""
        subShiftNews()
        nNews = nNews - 1
        subLoadNews()
        subClr()
    End Sub
    Private Sub btnDel3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel3.Click
        aNews(3, 1) = ""
        aNews(3, 2) = ""
        subShiftNews()
        nNews = nNews - 1
        subLoadNews()
        subClr()
    End Sub
    Private Sub btnDel4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel4.Click
        aNews(4, 1) = ""
        aNews(4, 2) = ""
        subShiftNews()
        nNews = nNews - 1
        subLoadNews()
        subClr()
    End Sub
    Private Sub btnDel5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel5.Click
        aNews(5, 1) = ""
        aNews(5, 2) = ""
        subShiftNews()
        nNews = nNews - 1
        subLoadNews()
        subClr()
    End Sub
    Private Sub btnDel6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel6.Click
        aNews(6, 1) = ""
        aNews(6, 2) = ""
        subShiftNews()
        nNews = nNews - 1
        subLoadNews()
        subClr()
    End Sub
    Private Sub btnDel7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel7.Click
        aNews(7, 1) = ""
        aNews(7, 2) = ""
        subShiftNews()
        nNews = nNews - 1
        subLoadNews()
        subClr()
    End Sub
    Private Sub btnDel8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel8.Click
        aNews(8, 1) = ""
        aNews(8, 2) = ""
        subShiftNews()
        nNews = nNews - 1
        subLoadNews()
        subClr()
    End Sub
    Private Sub btnDel9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel9.Click
        aNews(9, 1) = ""
        aNews(9, 2) = ""
        subShiftNews()
        nNews = nNews - 1
        subLoadNews()
        subClr()
    End Sub
    Private Sub btnDel10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel10.Click
        aNews(10, 1) = ""
        aNews(10, 2) = ""
        subShiftNews()
        nNews = nNews - 1
        subLoadNews()
        subClr()
    End Sub
    Private Sub btnChg1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChg1.Click
        bolUpdate = True
        nUpdate = 1
        btnAdd.Text = "Update News"
        btnClr.Visible = True
        tbNewTitle.Text = aNews(1, 1)
        tbNewItem.Text = aNews(1, 2)
    End Sub
    Private Sub btnChg2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChg2.Click
        bolUpdate = True
        nUpdate = 2
        btnAdd.Text = "Update News"
        btnClr.Visible = True
        tbNewTitle.Text = aNews(2, 1)
        tbNewItem.Text = aNews(2, 2)
    End Sub
    Private Sub btnChg3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChg3.Click
        bolUpdate = True
        nUpdate = 3
        btnAdd.Text = "Update News"
        btnClr.Visible = True
        tbNewTitle.Text = aNews(3, 1)
        tbNewItem.Text = aNews(3, 2)
    End Sub
    Private Sub btnChg4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChg4.Click
        bolUpdate = True
        nUpdate = 4
        btnAdd.Text = "Update News"
        btnClr.Visible = True
        tbNewTitle.Text = aNews(4, 1)
        tbNewItem.Text = aNews(4, 2)
    End Sub
    Private Sub btnChg5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChg5.Click
        bolUpdate = True
        nUpdate = 5
        btnAdd.Text = "Update News"
        btnClr.Visible = True
        tbNewTitle.Text = aNews(5, 1)
        tbNewItem.Text = aNews(5, 2)
    End Sub
    Private Sub btnChg6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChg6.Click
        bolUpdate = True
        nUpdate = 6
        btnAdd.Text = "Update News"
        btnClr.Visible = True
        tbNewTitle.Text = aNews(6, 1)
        tbNewItem.Text = aNews(6, 2)
    End Sub
    Private Sub btnChg7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChg7.Click
        bolUpdate = True
        nUpdate = 7
        btnAdd.Text = "Update News"
        btnClr.Visible = True
        tbNewTitle.Text = aNews(7, 1)
        tbNewItem.Text = aNews(7, 2)
    End Sub
    Private Sub btnChg8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChg8.Click
        bolUpdate = True
        nUpdate = 8
        btnAdd.Text = "Update News"
        btnClr.Visible = True
        tbNewTitle.Text = aNews(8, 1)
        tbNewItem.Text = aNews(8, 2)
    End Sub
    Private Sub btnChg9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChg9.Click
        bolUpdate = True
        nUpdate = 9
        btnAdd.Text = "Update News"
        btnClr.Visible = True
        tbNewTitle.Text = aNews(9, 1)
        tbNewItem.Text = aNews(9, 2)
    End Sub
    Private Sub btnChg10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChg10.Click
        bolUpdate = True
        nUpdate = 10
        btnAdd.Text = "Update News"
        btnClr.Visible = True
        tbNewTitle.Text = aNews(10, 1)
        tbNewItem.Text = aNews(10, 2)
    End Sub
    Private Sub btnClr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClr.Click
        subClr()
    End Sub
    Private Sub subClr()
        tbNewItem.Text = ""
        tbNewTitle.Text = ""
        bolUpdate = False
        btnClr.Visible = False
        btnAdd.Text = "Add News"
    End Sub
    Private Sub subSave()
        Dim x As Integer
        FileCopy(lPath & "main1.txt", lPath & "home.html")
        FileOpen(1, lPath & "home.html", OpenMode.Append)
        For x = 1 To nNews
            PrintLine(1, SPC(12), "<h4>" & aNews(x, 1) & "</h4>")
            PrintLine(1, SPC(12), "<p>" & aNews(x, 2) & "</p>")
        Next
        PrintLine(1, SPC(12), "<p align=""center"" style=""color: #91664f; font-size: 12px;"">News Last Updated: " & nDate.ToString("d MMMM yyyy hh:mm") & "</p>")
        FileOpen(2, lPath & "main2.txt", OpenMode.Input)
        While Not EOF(2)
            PrintLine(1, LineInput(2))
        End While
        FileClose(1)
        FileClose(2)
        FileCopy(lPath & "home.html", iPath & "home.htm")
    End Sub
    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        nDate = Now()
        subSave()
        subPutNews()
        Me.Close()
    End Sub
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        nDate = Now()
        subSave()
        subPutNews()
    End Sub
End Class