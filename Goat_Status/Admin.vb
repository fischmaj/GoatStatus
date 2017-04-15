Imports System.IO
Public Class Admin
    Dim lPath As String = ""
    Dim oPath As String = ""
    Dim bPath As String = ""
    Dim xVer As String = ""
    Dim sups(50) As String
    Dim sofs(50) As String
    Dim hang(30, 4) As String
    Dim alts(30, 3) As String
    Dim sr(7) As String
    Dim birds(30, 4) As String
    Dim arTab0(17, 6) As Decimal
    Dim arTab1(17, 6) As Decimal
    Dim nSup As Integer
    Dim nSOF As Integer
    Dim nAlts As Integer
    Dim nSR As Integer
    Dim nHang As Integer
    Dim nBird As Integer
    Private Sub Admin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim fname0 As String = "path.txt"
        Try
            Using fn0 As StreamReader = New StreamReader(fname0)
                lPath = fn0.ReadLine()
                oPath = fn0.ReadLine()
                bPath = fn0.ReadLine()
                xVer = fn0.ReadLine()
            End Using
        Catch ex As Exception
            MsgBox("Path File Error")
        End Try
        subLoadCombos()
        subShowSups()
        subShowSOFs()
        subShowStats()
        subShowAlts()
        subShowBirds()
        subShowTab0()
        subShowTab1()
        subShowSRs()
    End Sub
    Private Sub subLoadCombos()
        subGetAlt()
        subGetHang()
        subGetSOF()
        subGetSup()
        subGetBirds()
        subGetTab0()
        subGetTab1()
        subGetSR()
    End Sub
    Private Function fnFNBR(ByVal fn As String, ByVal path1 As String, ByVal path2 As String) As Boolean
        Try
            FileCopy(path1 & fn, path2 & fn)
            fnFNBR = True
        Catch ex As Exception
            fnFNBR = False
        End Try
    End Function
    Private Sub subGetSOF()
        Dim fname0 As String = lPath + "sof.txt"
        Try
            Using fn0 As StreamReader = New StreamReader(fname0)
                nSOF = CDec(fn0.ReadLine)
                For x = 1 To nSOF
                    sofs(x) = fn0.ReadLine()
                Next x
                fn0.Close()
            End Using
        Catch ex As Exception
            If fnFNBR("sof.txt", bPath, lPath) Then
                subGetSOF()
            Else
                MsgBox("admin:SoF File Read Error")
            End If
        End Try
    End Sub
    Private Sub subGetSup()
        Dim fname0 As String = lPath + "sup.txt"
        Try
            Using fn0 As StreamReader = New StreamReader(fname0)
                nSup = CDec(fn0.ReadLine)
                For x = 1 To nSup
                    sups(x) = fn0.ReadLine()
                Next x
                fn0.Close()
            End Using
        Catch ex As Exception
            If fnFNBR("sup.txt", bPath, lPath) Then
                subGetSup()
            Else
                MsgBox("admin:Sup File Read Error")
            End If
        End Try
    End Sub
    Private Sub subGetHang()
        Dim fname0 As String = lPath + "hangover.txt"
        Try
            Using fn0 As StreamReader = New StreamReader(fname0)
                nHang = CDec(fn0.ReadLine)
                For x = 1 To nHang
                    Dim iline As String = fn0.ReadLine()
                    Dim cols() As String = iline.Split(",")
                    hang(x, 0) = cols(0)
                    hang(x, 1) = cols(1)
                    hang(x, 2) = cols(2)
                    hang(x, 3) = cols(3)
                    hang(x, 4) = cols(4)
                Next x
                fn0.Close()
            End Using
        Catch ex As Exception
            If fnFNBR("hangover.txt", bPath, lPath) Then
                subGetHang()
            Else
                MsgBox("admin:Hangover File Read Error")
            End If
        End Try
    End Sub
    Private Sub subGetSR()
        Dim fname0 As String = lPath + "sr.txt"
        Try
            Using fn0 As StreamReader = New StreamReader(fname0)
                nSR = CDec(fn0.ReadLine)
                For x = 1 To nSR
                    Dim iline As String = fn0.ReadLine()
                    sr(x) = iline
                Next x
                fn0.Close()
            End Using
        Catch ex As Exception
            If fnFNBR("sr.txt", bPath, lPath) Then
                subGetSR()
            Else
                MsgBox("SR File Read Error")
            End If
        End Try
    End Sub
    Private Sub subGetAlt()
        Dim fname0 As String = lPath + "alternate.txt"
        Try
            Using fn0 As StreamReader = New StreamReader(fname0)
                nAlts = CDec(fn0.ReadLine)
                For x = 1 To nAlts
                    Dim iline As String = fn0.ReadLine()
                    Dim cols() As String = iline.Split(",")
                    alts(x, 1) = cols(0)
                    alts(x, 2) = cols(1)
                    alts(x, 3) = cols(2)
                Next x
                fn0.Close()
            End Using
        Catch ex As Exception
            If fnFNBR("alternate.txt", bPath, lPath) Then
                subGetAlt()
            Else
                MsgBox("admin:Alternate File Read Error")
            End If
        End Try
    End Sub
    Private Sub subGetBirds()
        Dim fname0 As String = lPath + "birds.txt"
        Try
            Using fn0 As StreamReader = New StreamReader(fname0)
                nBird = CDec(fn0.ReadLine)
                For x = 1 To nBird
                    Dim iline As String = fn0.ReadLine()
                    Dim cols() As String = iline.Split(",")
                    birds(x, 0) = cols(0)
                    birds(x, 1) = cols(1)
                    birds(x, 2) = cols(2)
                    birds(x, 3) = cols(3)
                    birds(x, 4) = cols(4)
                Next x
                fn0.Close()
            End Using
        Catch ex As Exception
            If fnFNBR("birds.txt", bPath, lPath) Then
                subGetBirds()
            Else
                MsgBox("admin:Bird Status File Read Error")
            End If
        End Try
    End Sub
    Private Sub subGetTab0()
        Dim fname0 As String = lPath & "tab0.txt"
        Try
            Using fn0 As StreamReader = New StreamReader(fname0)
                For x0 = 1 To 17
                    Dim iline As String = fn0.ReadLine()
                    Dim cols() As String = iline.Split(",")
                    For i = 1 To 6
                        arTab0(x0, i) = Val(cols(i - 1))
                    Next
                Next x0
                fn0.Close()
            End Using
        Catch ex As Exception
            If fnFNBR("tab0.txt", bPath, lPath) Then
                subGetTab0()
            Else
                MsgBox("admin:Tab 0 File Read Error")
            End If
        End Try
    End Sub
    Private Sub subGetTab1()
        Dim fname0 As String = lPath & "tab1.txt"
        Try
            Using fn0 As StreamReader = New StreamReader(fname0)
                For x1 = 1 To 17
                    Dim iline As String = fn0.ReadLine()
                    Dim cols() As String = iline.Split(",")
                    For i = 1 To 6
                        arTab1(x1, i) = Val(cols(i - 1))
                    Next
                Next x1
                fn0.Close()
            End Using
        Catch ex As Exception
            If fnFNBR("tab1.txt", bPath, lPath) Then
                subGetTab0()
            Else
                MsgBox("admin:Tab 1 File Read Error")
            End If
        End Try
    End Sub
    Private Sub subPutSOF()
        Dim fname0 As String = lPath + "sof.txt"
        Try
            Using fn0 As StreamWriter = New StreamWriter(fname0)
                fn0.WriteLine(nSOF)
                For x = 1 To nSOF
                    fn0.WriteLine(sofs(x))
                Next x
                fn0.Close()
            End Using
        Catch ex As Exception
            If fnFNBR("sof.txt", bPath, lPath) Then
                subPutSOF()
            Else
                MsgBox("admin:SoF File Write Error")
            End If
        End Try
    End Sub
    Private Sub subPutSup()
        Dim fname0 As String = lPath + "sup.txt"
        Try
            Using fn0 As StreamWriter = New StreamWriter(fname0)
                fn0.WriteLine(nSup)
                For x = 1 To nSup
                    fn0.WriteLine(sups(x))
                Next x
                fn0.Close()
            End Using
        Catch ex As Exception
            If fnFNBR("sup.txt", bPath, lPath) Then
                subPutSup()
            Else
                MsgBox("admin:Sup File Write Error")
            End If
        End Try
    End Sub
    Private Sub subPutSR()
        Dim fname0 As String = lPath + "sr.txt"
        Try
            Using fn0 As StreamWriter = New StreamWriter(fname0)
                fn0.WriteLine(nSR)
                For x = 1 To nSR
                    fn0.WriteLine(sr(x))
                Next x
                fn0.Close()
            End Using
        Catch ex As Exception
            If fnFNBR("sr.txt", bPath, lPath) Then
                subPutSR()
            Else
                MsgBox("admin:SR File Write Error")
            End If
        End Try
    End Sub
    Private Sub subPutHang()
        Dim fname0 As String = lPath + "hangover.txt"
        Try
            Using fn0 As StreamWriter = New StreamWriter(fname0)
                fn0.WriteLine(nHang)
                For x = 1 To nHang
                    fn0.Write(hang(x, 0) & ",")
                    fn0.Write(hang(x, 1) & ",")
                    fn0.Write(hang(x, 2) & ",")
                    fn0.Write(hang(x, 3) & ",")
                    fn0.WriteLine(hang(x, 4))
                Next x
                fn0.Close()
            End Using
        Catch ex As Exception
            If fnFNBR("hangover.txt", bPath, lPath) Then
                subPutHang()
            Else
                MsgBox("admin:Hangover File Write Error")
            End If
        End Try
    End Sub
    Private Sub subPutAlt()
        Dim fname0 As String = lPath + "alternate.txt"
        Try
            Using fn0 As StreamWriter = New StreamWriter(fname0)
                fn0.WriteLine(nAlts)
                For x = 1 To nAlts
                    fn0.Write(alts(x, 1) & ",")
                    fn0.Write(alts(x, 2) & ",")
                    fn0.WriteLine(alts(x, 3))
                Next x
                fn0.Close()
            End Using
        Catch ex As Exception
            If fnFNBR("alternate.txt", bPath, lPath) Then
                subPutAlt()
            Else
                MsgBox("admin:Alternate File Write Error")
            End If
        End Try
    End Sub
    Private Sub subPutBirds()
        Dim fname0 As String = lPath + "birds.txt"
        Try
            Using fn0 As StreamWriter = New StreamWriter(fname0)
                fn0.WriteLine(nBird)
                For x = 1 To nBird
                    fn0.Write(birds(x, 0) & ",")
                    fn0.Write(birds(x, 1) & ",")
                    fn0.Write(birds(x, 2) & ",")
                    fn0.Write(birds(x, 3) & ",")
                    fn0.WriteLine(birds(x, 4))
                Next x
                fn0.Close()
            End Using
        Catch ex As Exception
            If fnFNBR("birds.txt", bPath, lPath) Then
                subPutBirds()
            Else
                MsgBox("admin:Birds File Write Error")
            End If
        End Try
    End Sub
    Private Sub subPutTab0()
        Dim fname0 As String = lPath & "tab0.txt"
        Try
            Using fn0 As StreamWriter = New StreamWriter(fname0)
                For x0 = 1 To 17
                    For i = 1 To 5
                        fn0.Write(arTab0(x0, i) & ",")
                    Next
                    fn0.WriteLine(arTab0(x0, 6) & ",")
                Next x0
                fn0.Close()
            End Using
        Catch ex As Exception
            If fnFNBR("tab0.txt", bPath, lPath) Then
                subPutTab0()
            Else
                MsgBox("admin:Tab 0 File Write Error")
            End If
        End Try
    End Sub
    Private Sub subPutTab1()
        Dim fname0 As String = lPath & "tab1.txt"
        Try
            Using fn0 As StreamWriter = New StreamWriter(fname0)
                For x1 = 1 To 17
                    For i = 1 To 5
                        fn0.Write(arTab1(x1, i) & ",")
                    Next
                    fn0.WriteLine(arTab1(x1, 6) & ",")
                Next x1
                fn0.Close()
            End Using
        Catch ex As Exception
            If fnFNBR("tab1.txt", bPath, lPath) Then
                subPutTab0()
            Else
                MsgBox("admin:Tab 1 File Write Error")
            End If
        End Try
    End Sub
    Private Sub subShowSups()
        Dim l As Integer
        For l = 1 To nSup
            clbSups.Items.Add(sups(l))
        Next
    End Sub
    Private Sub subShowSOFs()
        Dim l As Integer
        For l = 1 To nSOF
            clbSof.Items.Add(sofs(l))
        Next
    End Sub
    Private Sub subShowStats()
        Dim l As Integer
        For l = 1 To nHang
            clbStatus.Items.Add(hang(l, 0) + ", " + hang(l, 1) + ", " + _
                                hang(l, 2) + ", " + hang(l, 3) + ", " + _
                                hang(l, 4))
        Next
    End Sub
    Private Sub subShowAlts()
        Dim l As Integer
        For l = 1 To nAlts
            clbAlt.Items.Add(alts(l, 1) + ", " + alts(l, 2) + ", " + alts(l, 3))
        Next
    End Sub
    Private Sub subShowBirds()
        Dim l As Integer
        For l = 1 To nBird
            clbBirds.Items.Add(birds(l, 0) + ", " + birds(l, 1) + ", " + _
                                birds(l, 2) + ", " + birds(l, 3) + ", " + _
                                birds(l, 4))
        Next
    End Sub
    Private Sub subShowSRs()
        tbSR1.Text = sr(1)
        tbSR2.Text = sr(2)
        tbSR3.Text = sr(3)
        tbSR4.Text = sr(4)
        tbSR5.Text = sr(5)
    End Sub
    Private Sub subShowTab0()
        tbTab011.Text = arTab0(1, 1)
        tbTab012.Text = arTab0(1, 2)
        tbTab013.Text = arTab0(1, 3)
        tbTab014.Text = arTab0(1, 4)
        tbTab015.Text = arTab0(1, 5)
        tbTab016.Text = arTab0(1, 6)
        tbTab021.Text = arTab0(2, 1)
        tbTab022.Text = arTab0(2, 2)
        tbTab023.Text = arTab0(2, 3)
        tbTab024.Text = arTab0(2, 4)
        tbTab025.Text = arTab0(2, 5)
        tbTab026.Text = arTab0(2, 6)
        tbTab031.Text = arTab0(3, 1)
        tbTab032.Text = arTab0(3, 2)
        tbTab033.Text = arTab0(3, 3)
        tbTab034.Text = arTab0(3, 4)
        tbTab035.Text = arTab0(3, 5)
        tbTab036.Text = arTab0(3, 6)
        tbTab041.Text = arTab0(4, 1)
        tbTab042.Text = arTab0(4, 2)
        tbTab043.Text = arTab0(4, 3)
        tbTab044.Text = arTab0(4, 4)
        tbTab045.Text = arTab0(4, 5)
        tbTab046.Text = arTab0(4, 6)
        tbTab051.Text = arTab0(5, 1)
        tbTab052.Text = arTab0(5, 2)
        tbTab053.Text = arTab0(5, 3)
        tbTab054.Text = arTab0(5, 4)
        tbTab055.Text = arTab0(5, 5)
        tbTab056.Text = arTab0(5, 6)
        tbTab061.Text = arTab0(6, 1)
        tbTab062.Text = arTab0(6, 2)
        tbTab063.Text = arTab0(6, 3)
        tbTab064.Text = arTab0(6, 4)
        tbTab065.Text = arTab0(6, 5)
        tbTab066.Text = arTab0(6, 6)
        tbTab071.Text = arTab0(7, 1)
        tbTab072.Text = arTab0(7, 2)
        tbTab073.Text = arTab0(7, 3)
        tbTab074.Text = arTab0(7, 4)
        tbTab075.Text = arTab0(7, 5)
        tbTab076.Text = arTab0(7, 6)
        tbTab081.Text = arTab0(8, 1)
        tbTab082.Text = arTab0(8, 2)
        tbTab083.Text = arTab0(8, 3)
        tbTab084.Text = arTab0(8, 4)
        tbTab085.Text = arTab0(8, 5)
        tbTab086.Text = arTab0(8, 6)
        tbTab091.Text = arTab0(9, 1)
        tbTab092.Text = arTab0(9, 2)
        tbTab093.Text = arTab0(9, 3)
        tbTab094.Text = arTab0(9, 4)
        tbTab095.Text = arTab0(9, 5)
        tbTab096.Text = arTab0(9, 6)
        tbTab0101.Text = arTab0(10, 1)
        tbTab0102.Text = arTab0(10, 2)
        tbTab0103.Text = arTab0(10, 3)
        tbTab0104.Text = arTab0(10, 4)
        tbTab0105.Text = arTab0(10, 5)
        tbTab0106.Text = arTab0(10, 6)
        tbTab0111.Text = arTab0(11, 1)
        tbTab0112.Text = arTab0(11, 2)
        tbTab0113.Text = arTab0(11, 3)
        tbTab0114.Text = arTab0(11, 4)
        tbTab0115.Text = arTab0(11, 5)
        tbTab0116.Text = arTab0(11, 6)
        tbTab0121.Text = arTab0(12, 1)
        tbTab0122.Text = arTab0(12, 2)
        tbTab0123.Text = arTab0(12, 3)
        tbTab0124.Text = arTab0(12, 4)
        tbTab0125.Text = arTab0(12, 5)
        tbTab0126.Text = arTab0(12, 6)
        tbTab0131.Text = arTab0(13, 1)
        tbTab0132.Text = arTab0(13, 2)
        tbTab0133.Text = arTab0(13, 3)
        tbTab0134.Text = arTab0(13, 4)
        tbTab0135.Text = arTab0(13, 5)
        tbTab0136.Text = arTab0(13, 6)
        tbTab0141.Text = arTab0(14, 1)
        tbTab0142.Text = arTab0(14, 2)
        tbTab0143.Text = arTab0(14, 3)
        tbTab0144.Text = arTab0(14, 4)
        tbTab0145.Text = arTab0(14, 5)
        tbTab0146.Text = arTab0(14, 6)
        tbTab0151.Text = arTab0(15, 1)
        tbTab0152.Text = arTab0(15, 2)
        tbTab0153.Text = arTab0(15, 3)
        tbTab0154.Text = arTab0(15, 4)
        tbTab0155.Text = arTab0(15, 5)
        tbTab0156.Text = arTab0(15, 6)
        tbTab0161.Text = arTab0(16, 1)
        tbTab0162.Text = arTab0(16, 2)
        tbTab0163.Text = arTab0(16, 3)
        tbTab0164.Text = arTab0(16, 4)
        tbTab0165.Text = arTab0(16, 5)
        tbTab0166.Text = arTab0(16, 6)
        tbTab0171.Text = arTab0(17, 1)
        tbTab0172.Text = arTab0(17, 2)
        tbTab0173.Text = arTab0(17, 3)
        tbTab0174.Text = arTab0(17, 4)
        tbTab0175.Text = arTab0(17, 5)
        tbTab0176.Text = arTab0(17, 6)
    End Sub
    Private Sub subSaveTab0()
        arTab0(1, 1) = tbTab011.Text
        arTab0(1, 2) = tbTab012.Text
        arTab0(1, 3) = tbTab013.Text
        arTab0(1, 4) = tbTab014.Text
        arTab0(1, 5) = tbTab015.Text
        arTab0(1, 6) = tbTab016.Text
        arTab0(2, 1) = tbTab021.Text
        arTab0(2, 2) = tbTab022.Text
        arTab0(2, 3) = tbTab023.Text
        arTab0(2, 4) = tbTab024.Text
        arTab0(2, 5) = tbTab025.Text
        arTab0(2, 6) = tbTab026.Text
        arTab0(3, 1) = tbTab031.Text
        arTab0(3, 2) = tbTab032.Text
        arTab0(3, 3) = tbTab033.Text
        arTab0(3, 4) = tbTab034.Text
        arTab0(3, 5) = tbTab035.Text
        arTab0(3, 6) = tbTab036.Text
        arTab0(4, 1) = tbTab041.Text
        arTab0(4, 2) = tbTab042.Text
        arTab0(4, 3) = tbTab043.Text
        arTab0(4, 4) = tbTab044.Text
        arTab0(4, 5) = tbTab045.Text
        arTab0(4, 6) = tbTab046.Text
        arTab0(5, 1) = tbTab051.Text
        arTab0(5, 2) = tbTab052.Text
        arTab0(5, 3) = tbTab053.Text
        arTab0(5, 4) = tbTab054.Text
        arTab0(5, 5) = tbTab055.Text
        arTab0(5, 6) = tbTab056.Text
        arTab0(6, 1) = tbTab061.Text
        arTab0(6, 2) = tbTab062.Text
        arTab0(6, 3) = tbTab063.Text
        arTab0(6, 4) = tbTab064.Text
        arTab0(6, 5) = tbTab065.Text
        arTab0(6, 6) = tbTab066.Text
        arTab0(7, 1) = tbTab071.Text
        arTab0(7, 2) = tbTab072.Text
        arTab0(7, 3) = tbTab073.Text
        arTab0(7, 4) = tbTab074.Text
        arTab0(7, 5) = tbTab075.Text
        arTab0(7, 6) = tbTab076.Text
        arTab0(8, 1) = tbTab081.Text
        arTab0(8, 2) = tbTab082.Text
        arTab0(8, 3) = tbTab083.Text
        arTab0(8, 4) = tbTab084.Text
        arTab0(8, 5) = tbTab085.Text
        arTab0(8, 6) = tbTab086.Text
        arTab0(9, 1) = tbTab091.Text
        arTab0(9, 2) = tbTab092.Text
        arTab0(9, 3) = tbTab093.Text
        arTab0(9, 4) = tbTab094.Text
        arTab0(9, 5) = tbTab095.Text
        arTab0(9, 6) = tbTab096.Text
        arTab0(10, 1) = tbTab0101.Text
        arTab0(10, 2) = tbTab0102.Text
        arTab0(10, 3) = tbTab0103.Text
        arTab0(10, 4) = tbTab0104.Text
        arTab0(10, 5) = tbTab0105.Text
        arTab0(10, 6) = tbTab0106.Text
        arTab0(11, 1) = tbTab0111.Text
        arTab0(11, 2) = tbTab0112.Text
        arTab0(11, 3) = tbTab0113.Text
        arTab0(11, 4) = tbTab0114.Text
        arTab0(11, 5) = tbTab0115.Text
        arTab0(11, 6) = tbTab0116.Text
        arTab0(12, 1) = tbTab0121.Text
        arTab0(12, 2) = tbTab0122.Text
        arTab0(12, 3) = tbTab0123.Text
        arTab0(12, 4) = tbTab0124.Text
        arTab0(12, 5) = tbTab0125.Text
        arTab0(12, 6) = tbTab0126.Text
        arTab0(13, 1) = tbTab0131.Text
        arTab0(13, 2) = tbTab0132.Text
        arTab0(13, 3) = tbTab0133.Text
        arTab0(13, 4) = tbTab0134.Text
        arTab0(13, 5) = tbTab0135.Text
        arTab0(13, 6) = tbTab0136.Text
        arTab0(14, 1) = tbTab0141.Text
        arTab0(14, 2) = tbTab0142.Text
        arTab0(14, 3) = tbTab0143.Text
        arTab0(14, 4) = tbTab0144.Text
        arTab0(14, 5) = tbTab0145.Text
        arTab0(14, 6) = tbTab0146.Text
        arTab0(15, 1) = tbTab0151.Text
        arTab0(15, 2) = tbTab0152.Text
        arTab0(15, 3) = tbTab0153.Text
        arTab0(15, 4) = tbTab0154.Text
        arTab0(15, 5) = tbTab0155.Text
        arTab0(15, 6) = tbTab0156.Text
        arTab0(16, 1) = tbTab0161.Text
        arTab0(16, 2) = tbTab0162.Text
        arTab0(16, 3) = tbTab0163.Text
        arTab0(16, 4) = tbTab0164.Text
        arTab0(16, 5) = tbTab0165.Text
        arTab0(16, 6) = tbTab0166.Text
        arTab0(17, 1) = tbTab0171.Text
        arTab0(17, 2) = tbTab0172.Text
        arTab0(17, 3) = tbTab0173.Text
        arTab0(17, 4) = tbTab0174.Text
        arTab0(17, 5) = tbTab0175.Text
        arTab0(17, 6) = tbTab0176.Text
        subPutTab0()
    End Sub
    Private Sub subShowTab1()
        tbTab111.Text = arTab1(1, 1)
        tbTab112.Text = arTab1(1, 2)
        tbTab113.Text = arTab1(1, 3)
        tbTab114.Text = arTab1(1, 4)
        tbTab115.Text = arTab1(1, 5)
        tbTab116.Text = arTab1(1, 6)
        tbTab121.Text = arTab1(2, 1)
        tbTab122.Text = arTab1(2, 2)
        tbTab123.Text = arTab1(2, 3)
        tbTab124.Text = arTab1(2, 4)
        tbTab125.Text = arTab1(2, 5)
        tbTab126.Text = arTab1(2, 6)
        tbTab131.Text = arTab1(3, 1)
        tbTab132.Text = arTab1(3, 2)
        tbTab133.Text = arTab1(3, 3)
        tbTab134.Text = arTab1(3, 4)
        tbTab135.Text = arTab1(3, 5)
        tbTab136.Text = arTab1(3, 6)
        tbTab141.Text = arTab1(4, 1)
        tbTab142.Text = arTab1(4, 2)
        tbTab143.Text = arTab1(4, 3)
        tbTab144.Text = arTab1(4, 4)
        tbTab145.Text = arTab1(4, 5)
        tbTab146.Text = arTab1(4, 6)
        tbTab151.Text = arTab1(5, 1)
        tbTab152.Text = arTab1(5, 2)
        tbTab153.Text = arTab1(5, 3)
        tbTab154.Text = arTab1(5, 4)
        tbTab155.Text = arTab1(5, 5)
        tbTab156.Text = arTab1(5, 6)
        tbTab161.Text = arTab1(6, 1)
        tbTab162.Text = arTab1(6, 2)
        tbTab163.Text = arTab1(6, 3)
        tbTab164.Text = arTab1(6, 4)
        tbTab165.Text = arTab1(6, 5)
        tbTab166.Text = arTab1(6, 6)
        tbTab171.Text = arTab1(7, 1)
        tbTab172.Text = arTab1(7, 2)
        tbTab173.Text = arTab1(7, 3)
        tbTab174.Text = arTab1(7, 4)
        tbTab175.Text = arTab1(7, 5)
        tbTab176.Text = arTab1(7, 6)
        tbTab181.Text = arTab1(8, 1)
        tbTab182.Text = arTab1(8, 2)
        tbTab183.Text = arTab1(8, 3)
        tbTab184.Text = arTab1(8, 4)
        tbTab185.Text = arTab1(8, 5)
        tbTab186.Text = arTab1(8, 6)
        tbTab191.Text = arTab1(9, 1)
        tbTab192.Text = arTab1(9, 2)
        tbTab193.Text = arTab1(9, 3)
        tbTab194.Text = arTab1(9, 4)
        tbTab195.Text = arTab1(9, 5)
        tbTab196.Text = arTab1(9, 6)
        tbTab1101.Text = arTab1(10, 1)
        tbTab1102.Text = arTab1(10, 2)
        tbTab1103.Text = arTab1(10, 3)
        tbTab1104.Text = arTab1(10, 4)
        tbTab1105.Text = arTab1(10, 5)
        tbTab1106.Text = arTab1(10, 6)
        tbTab1111.Text = arTab1(11, 1)
        tbTab1112.Text = arTab1(11, 2)
        tbTab1113.Text = arTab1(11, 3)
        tbTab1114.Text = arTab1(11, 4)
        tbTab1115.Text = arTab1(11, 5)
        tbTab1116.Text = arTab1(11, 6)
        tbTab1121.Text = arTab1(12, 1)
        tbTab1122.Text = arTab1(12, 2)
        tbTab1123.Text = arTab1(12, 3)
        tbTab1124.Text = arTab1(12, 4)
        tbTab1125.Text = arTab1(12, 5)
        tbTab1126.Text = arTab1(12, 6)
        tbTab1131.Text = arTab1(13, 1)
        tbTab1132.Text = arTab1(13, 2)
        tbTab1133.Text = arTab1(13, 3)
        tbTab1134.Text = arTab1(13, 4)
        tbTab1135.Text = arTab1(13, 5)
        tbTab1136.Text = arTab1(13, 6)
        tbTab1141.Text = arTab1(14, 1)
        tbTab1142.Text = arTab1(14, 2)
        tbTab1143.Text = arTab1(14, 3)
        tbTab1144.Text = arTab1(14, 4)
        tbTab1145.Text = arTab1(14, 5)
        tbTab1146.Text = arTab1(14, 6)
        tbTab1151.Text = arTab1(15, 1)
        tbTab1152.Text = arTab1(15, 2)
        tbTab1153.Text = arTab1(15, 3)
        tbTab1154.Text = arTab1(15, 4)
        tbTab1155.Text = arTab1(15, 5)
        tbTab1156.Text = arTab1(15, 6)
        tbTab1161.Text = arTab1(16, 1)
        tbTab1162.Text = arTab1(16, 2)
        tbTab1163.Text = arTab1(16, 3)
        tbTab1164.Text = arTab1(16, 4)
        tbTab1165.Text = arTab1(16, 5)
        tbTab1166.Text = arTab1(16, 6)
        tbTab1171.Text = arTab1(17, 1)
        tbTab1172.Text = arTab1(17, 2)
        tbTab1173.Text = arTab1(17, 3)
        tbTab1174.Text = arTab1(17, 4)
        tbTab1175.Text = arTab1(17, 5)
        tbTab1176.Text = arTab1(17, 6)
    End Sub
    Private Sub subSaveTab1()

        arTab1(1, 1) = tbTab111.Text
        arTab1(1, 2) = tbTab112.Text
        arTab1(1, 3) = tbTab113.Text
        arTab1(1, 4) = tbTab114.Text
        arTab1(1, 5) = tbTab115.Text
        arTab1(1, 6) = tbTab116.Text
        arTab1(2, 1) = tbTab121.Text
        arTab1(2, 2) = tbTab122.Text
        arTab1(2, 3) = tbTab123.Text
        arTab1(2, 4) = tbTab124.Text
        arTab1(2, 5) = tbTab125.Text
        arTab1(2, 6) = tbTab126.Text
        arTab1(3, 1) = tbTab131.Text
        arTab1(3, 2) = tbTab132.Text
        arTab1(3, 3) = tbTab133.Text
        arTab1(3, 4) = tbTab134.Text
        arTab1(3, 5) = tbTab135.Text
        arTab1(3, 6) = tbTab136.Text
        arTab1(4, 1) = tbTab141.Text
        arTab1(4, 2) = tbTab142.Text
        arTab1(4, 3) = tbTab143.Text
        arTab1(4, 4) = tbTab144.Text
        arTab1(4, 5) = tbTab145.Text
        arTab1(4, 6) = tbTab146.Text
        arTab1(5, 1) = tbTab151.Text
        arTab1(5, 2) = tbTab152.Text
        arTab1(5, 3) = tbTab153.Text
        arTab1(5, 4) = tbTab154.Text
        arTab1(5, 5) = tbTab155.Text
        arTab1(5, 6) = tbTab156.Text
        arTab1(6, 1) = tbTab161.Text
        arTab1(6, 2) = tbTab162.Text
        arTab1(6, 3) = tbTab163.Text
        arTab1(6, 4) = tbTab164.Text
        arTab1(6, 5) = tbTab165.Text
        arTab1(6, 6) = tbTab166.Text
        arTab1(7, 1) = tbTab171.Text
        arTab1(7, 2) = tbTab172.Text
        arTab1(7, 3) = tbTab173.Text
        arTab1(7, 4) = tbTab174.Text
        arTab1(7, 5) = tbTab175.Text
        arTab1(7, 6) = tbTab176.Text
        arTab1(8, 1) = tbTab181.Text
        arTab1(8, 2) = tbTab182.Text
        arTab1(8, 3) = tbTab183.Text
        arTab1(8, 4) = tbTab184.Text
        arTab1(8, 5) = tbTab185.Text
        arTab1(8, 6) = tbTab186.Text
        arTab1(9, 1) = tbTab191.Text
        arTab1(9, 2) = tbTab192.Text
        arTab1(9, 3) = tbTab193.Text
        arTab1(9, 4) = tbTab194.Text
        arTab1(9, 5) = tbTab195.Text
        arTab1(9, 6) = tbTab196.Text
        arTab1(10, 1) = tbTab1101.Text
        arTab1(10, 2) = tbTab1102.Text
        arTab1(10, 3) = tbTab1103.Text
        arTab1(10, 4) = tbTab1104.Text
        arTab1(10, 5) = tbTab1105.Text
        arTab1(10, 6) = tbTab1106.Text
        arTab1(11, 1) = tbTab1111.Text
        arTab1(11, 2) = tbTab1112.Text
        arTab1(11, 3) = tbTab1113.Text
        arTab1(11, 4) = tbTab1114.Text
        arTab1(11, 5) = tbTab1115.Text
        arTab1(11, 6) = tbTab1116.Text
        arTab1(12, 1) = tbTab1121.Text
        arTab1(12, 2) = tbTab1122.Text
        arTab1(12, 3) = tbTab1123.Text
        arTab1(12, 4) = tbTab1124.Text
        arTab1(12, 5) = tbTab1125.Text
        arTab1(12, 6) = tbTab1126.Text
        arTab1(13, 1) = tbTab1131.Text
        arTab1(13, 2) = tbTab1132.Text
        arTab1(13, 3) = tbTab1133.Text
        arTab1(13, 4) = tbTab1134.Text
        arTab1(13, 5) = tbTab1135.Text
        arTab1(13, 6) = tbTab1136.Text
        arTab1(14, 1) = tbTab1141.Text
        arTab1(14, 2) = tbTab1142.Text
        arTab1(14, 3) = tbTab1143.Text
        arTab1(14, 4) = tbTab1144.Text
        arTab1(14, 5) = tbTab1145.Text
        arTab1(14, 6) = tbTab1146.Text
        arTab1(15, 1) = tbTab1151.Text
        arTab1(15, 2) = tbTab1152.Text
        arTab1(15, 3) = tbTab1153.Text
        arTab1(15, 4) = tbTab1154.Text
        arTab1(15, 5) = tbTab1155.Text
        arTab1(15, 6) = tbTab1156.Text
        arTab1(16, 1) = tbTab1161.Text
        arTab1(16, 2) = tbTab1162.Text
        arTab1(16, 3) = tbTab1163.Text
        arTab1(16, 4) = tbTab1164.Text
        arTab1(16, 5) = tbTab1165.Text
        arTab1(16, 6) = tbTab1166.Text
        arTab1(17, 1) = tbTab1171.Text
        arTab1(17, 2) = tbTab1172.Text
        arTab1(17, 3) = tbTab1173.Text
        arTab1(17, 4) = tbTab1174.Text
        arTab1(17, 5) = tbTab1175.Text
        arTab1(17, 6) = tbTab1176.Text
        subPutTab1()
    End Sub
    Private Sub subSaveSRs()
        sr(1) = tbSR1.Text
        sr(2) = tbSR2.Text
        sr(3) = tbSR3.Text
        sr(4) = tbSR4.Text
        sr(5) = tbSR5.Text
        subPutSR()
    End Sub
    Private Sub subShift(ByVal x As Integer, ByVal r As String)
        Dim y As Integer = x + 1
        If r = "Sup" Then
            If x < nSup Then
                sups(x) = sups(y)
                subShift(y, "Sup")
            End If
        ElseIf r = "Sof" Then
            If x < nSOF Then
                sofs(x) = sofs(y)
                subShift(y, "Sof")
            End If
        ElseIf r = "Alt" Then
            If x < nAlts Then
                alts(x, 1) = alts(y, 1)
                alts(x, 2) = alts(y, 2)
                alts(x, 3) = alts(y, 3)
                subShift(y, "Alt")
            End If
        ElseIf r = "Stat" Then
            If x < nHang Then
                hang(x, 0) = hang(y, 0)
                hang(x, 1) = hang(y, 1)
                hang(x, 2) = hang(y, 2)
                hang(x, 3) = hang(y, 3)
                hang(x, 4) = hang(y, 4)
                subShift(y, "Stat")
            End If
        ElseIf r = "Birds" Then
            If x < nBird Then
                birds(x, 0) = birds(y, 0)
                birds(x, 1) = birds(y, 1)
                birds(x, 2) = birds(y, 2)
                birds(x, 3) = birds(y, 3)
                birds(x, 4) = birds(y, 4)
                subShift(y, "Birds")
            End If
        End If
    End Sub
    Private Sub subSortSup()
        Dim temps As String
        Dim switch As Boolean
        Dim y As Integer
        Do
            switch = False
            For x = 1 To nSup - 1
                y = x + 1
                If sups(x) > sups(y) Then
                    temps = sups(x)
                    sups(x) = sups(y)
                    sups(y) = temps
                    switch = True
                End If
            Next
        Loop Until Not switch
    End Sub
    Private Sub subSortSof()
        Dim temps As String
        Dim switch As Boolean
        Dim y As Integer
        Do
            switch = False
            For x = 1 To nSOF - 1
                y = x + 1
                If sofs(x) > sofs(y) Then
                    temps = sofs(x)
                    sofs(x) = sofs(y)
                    sofs(y) = temps
                    switch = True
                End If
            Next
        Loop Until Not switch
    End Sub
    Private Sub subSortStat()
        Dim temps As String
        Dim switch As Boolean
        Dim y As Integer
        Do
            switch = False
            For x = 1 To nHang - 1
                y = x + 1
                If Val(hang(x, 0)) > Val(hang(y, 0)) Then
                    temps = hang(x, 0)
                    hang(x, 0) = hang(y, 0)
                    hang(y, 0) = temps
                    temps = hang(x, 1)
                    hang(x, 1) = hang(y, 1)
                    hang(y, 1) = temps
                    temps = hang(x, 2)
                    hang(x, 2) = hang(y, 2)
                    hang(y, 2) = temps
                    temps = hang(x, 3)
                    hang(x, 3) = hang(y, 3)
                    hang(y, 3) = temps
                    temps = hang(x, 4)
                    hang(x, 4) = hang(y, 4)
                    hang(y, 4) = temps
                    switch = True
                End If
            Next
        Loop Until Not switch
        For x = 1 To nHang
            hang(x, 0) = x * 2
        Next
    End Sub
    Private Sub subSortAlt()
        Dim temps As String
        Dim switch As Boolean
        Dim y As Integer
        Do
            switch = False
            For x = 1 To nAlts - 1
                y = x + 1
                If Val(alts(x, 3)) > Val(alts(y, 3)) Then
                    temps = alts(x, 1)
                    alts(x, 1) = alts(y, 1)
                    alts(y, 1) = temps
                    temps = alts(x, 2)
                    alts(x, 2) = alts(y, 2)
                    alts(y, 2) = temps
                    temps = alts(x, 3)
                    alts(x, 3) = alts(y, 3)
                    alts(y, 3) = temps
                    switch = True
                End If
            Next
        Loop Until Not switch
    End Sub
    Private Sub subSortBirds()
        Dim temps As String
        Dim switch As Boolean
        Dim y As Integer
        Do
            switch = False
            For x = 1 To nBird - 1
                y = x + 1
                If Val(birds(x, 0)) > Val(birds(y, 0)) Then
                    temps = birds(x, 0)
                    birds(x, 0) = birds(y, 0)
                    birds(y, 0) = temps
                    temps = birds(x, 1)
                    birds(x, 1) = birds(y, 1)
                    birds(y, 1) = temps
                    temps = birds(x, 2)
                    birds(x, 2) = birds(y, 2)
                    birds(y, 2) = temps
                    temps = birds(x, 3)
                    birds(x, 3) = birds(y, 3)
                    birds(y, 3) = temps
                    temps = birds(x, 4)
                    birds(x, 4) = birds(y, 4)
                    birds(y, 4) = temps
                    switch = True
                End If
            Next
        Loop Until Not switch
        For x = 1 To nBird
            birds(x, 0) = x * 2
        Next
    End Sub
    Private Sub btnDelSup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelSup.Click
        Dim x As Integer
        For x = clbSups.Items.Count - 1 To 0 Step -1
            If clbSups.GetSelected(x) = True Then
                subShift(x + 1, "Sup")
                nSup = nSup - 1
            End If
        Next
        clbSups.Items.Clear()
        subSortSup()
        subShowSups()
        subPutSup()
    End Sub
    Private Sub btnDelSof_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelSof.Click
        Dim x As Integer
        For x = clbSof.Items.Count - 1 To 0 Step -1
            If clbSof.GetSelected(x) = True Then
                subShift(x + 1, "Sof")
                nSOF = nSOF - 1
            End If
        Next
        clbSof.Items.Clear()
        subSortSof()
        subShowSOFs()
        subPutSOF()
    End Sub
    Private Sub btnDelStat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelStat.Click
        Dim x As Integer
        For x = clbStatus.Items.Count - 1 To 0 Step -1
            If clbStatus.GetSelected(x) = True Then
                subShift(x + 1, "Stat")
                nHang = nHang - 1
            End If
        Next
        clbStatus.Items.Clear()
        subSortStat()
        subShowStats()
        subPutHang()
    End Sub
    Private Sub btnDelAlt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelAlt.Click
        Dim x As Integer
        For x = clbAlt.Items.Count - 1 To 0 Step -1
            If clbAlt.GetSelected(x) = True Then
                subShift(x + 1, "Alt")
                nAlts = nAlts - 1
            End If
        Next
        clbAlt.Items.Clear()
        subSortAlt()
        subShowAlts()
        subPutAlt()
    End Sub
    Private Sub btnDelBird_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelBird.Click
        Dim x As Integer
        For x = clbBirds.Items.Count - 1 To 0 Step -1
            If clbBirds.GetSelected(x) = True Then
                subShift(x + 1, "Birds")
                nBird = nBird - 1
            End If
        Next
        clbBirds.Items.Clear()
        subSortBirds()
        subShowBirds()
        subPutBirds()
    End Sub
    Private Sub btnAddSup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddSup.Click
        nSup = nSup + 1
        sups(nSup) = tbSupsName.Text
        clbSups.Items.Clear()
        subSortSup()
        subShowSups()
        subPutSup()
        tbSupsName.Text = ""
    End Sub
    Private Sub btnAddSof_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddSof.Click
        nSOF = nSOF + 1
        sofs(nSOF) = tbSofsName.Text
        clbSof.Items.Clear()
        subSortSof()
        subShowSOFs()
        subPutSOF()
        tbSofsName.Text = ""
    End Sub
    Private Sub btnAddStat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddStat.Click
        nHang = nHang + 1
        hang(nHang, 0) = tbOrder.Text
        hang(nHang, 1) = tbStatName.Text
        hang(nHang, 2) = tbStatD1.Text
        hang(nHang, 3) = tbStatD2.Text
        hang(nHang, 4) = cbColor.Text
        clbStatus.Items.Clear()
        subSortStat()
        subShowStats()
        subPutHang()
        tbStatName.Text = ""
        tbStatD1.Text = ""
        tbStatD2.Text = ""
        tbOrder.Text = ""
        cbColor.Text = "Green"
    End Sub
    Private Sub btnAddalt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddAlt.Click
        nAlts = nAlts + 1
        alts(nAlts, 1) = tbAltName.Text
        alts(nAlts, 2) = tbAltIcao.Text
        alts(nAlts, 3) = tbAltFuel.Text
        clbAlt.Items.Clear()
        subSortAlt()
        subShowAlts()
        subPutAlt()
        tbAltName.Text = ""
        tbAltIcao.Text = ""
        tbAltFuel.Text = ""
    End Sub
    Private Sub btnAddBird_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddBird.Click
        nBird = nBird + 1
        birds(nBird, 0) = tbBirdOrd.Text
        birds(nBird, 1) = tbBirdStat.Text
        birds(nBird, 2) = tbBird1.Text
        birds(nBird, 3) = tbBird2.Text
        birds(nBird, 4) = cbBirdColor.Text
        clbBirds.Items.Clear()
        subSortBirds()
        subShowBirds()
        subPutBirds()
        tbBirdStat.Text = ""
        tbBird1.Text = ""
        tbBird2.Text = ""
        tbBirdOrd.Text = ""
        cbBirdColor.Text = "Green"
    End Sub
    Private Sub btnSaveSRs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveSRs.Click
        subSaveSRs()
    End Sub
    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Status.subBackupFiles()
        Me.Close()
    End Sub
    Private Sub Admin_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Status.subGetFields()
    End Sub
    Private Sub btnSaveTab0_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveTab0.Click
        subSaveTab0()
    End Sub
    Private Sub btnSaveTab1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveTab1.Click
        subSaveTab1()
    End Sub
End Class