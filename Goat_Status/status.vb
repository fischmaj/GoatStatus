' Status Board Selector
' Version 1 - July 2008
' Version 1.1 - August 2008 
' Version 1.2 - October 2008
' Version 1.3 - November 2008
' Version 1.3.1 - 24 November 2008
' Version 1.3.2 - 8 December 2008
' Version 1.4 - 2 April 2009
' Version 1.4.1 - 6 April 2009
' Version 1.4.2 - 20 July 2009
' Version 1.4.3 - 16 September 2009
' Version 1.5 - 15 October 2009
' Version 1.6 - 29 October 2009
' Version 1.7.1 - 16 May 2012
' Version 1.7.2 - 14 July 2012
Imports System.Xml
Imports System.IO
Public Class Status
    Dim lPath As String = ""
    Dim oPath As String = ""
    Dim bPath As String = ""
    Dim iPath As String = ""
    Dim xVer As String = ""
    Dim stats(30, 4) As String
    Dim alts(30, 3) As String
    Dim birds(10, 4) As String
    Dim its(3) As String
    Dim MOA(3) As String
    Dim sr(6) As String
    Dim gblStatus As Integer = 0
    Dim gblViagra As Integer = 0
    Dim arTab0(17, 6) As Decimal
    Dim arTab1(17, 6) As Decimal
    Dim degF As String = "°F"
    Dim degC As String = "°C"
    Dim varSOF As String = ""
    Dim varSUP As String = ""
    Dim varStatus As String = ""
    Dim varStat1 As String = ""
    Dim varStat2 As String = ""
    Dim varStatCol As String = ""
    Dim varHangRwy As String = ""
    Dim varWetDry As String = ""
    Dim varAlternate As String = ""
    Dim varAltFuel As String = ""
    Dim varAHC As String = ""
    Dim varA1B As String = ""
    Dim varA1E As String = ""
    Dim varA2B As String = ""
    Dim varA2E As String = ""
    Dim varA3B As String = ""
    Dim varA3E As String = ""
    Dim varBirds As String = ""
    Dim varBird1 As String = ""
    Dim varBird2 As String = ""
    Dim varBirdColor As String = ""
    Dim varITS As String = ""
    Dim varMoaSouth As String = ""
    Dim varMoaTweet As String = ""
    Dim varMoaTalon As String = ""
    Dim varSR1 As String = ""
    Dim varSR2 As String = ""
    Dim varSR3 As String = ""
    Dim varSR4 As String = ""
    Dim varSR5 As String = ""
    Dim varNavRND As String = ""
    Dim varNavILS As String = ""
    Dim varNavDME As String = ""
    Dim varNavSAT As String = ""
    Dim varNavSKF As String = ""
    Dim varNavSSF As String = ""
    Dim varOps1 As String = ""
    Dim varOps2 As String = ""
    Dim varOps3 As String = ""
    Dim varOps4 As String = ""
    Dim varOps5 As String = ""
    Dim varOps6 As String = ""
    Dim varOps7 As String = ""
    Dim varQoD As String = ""
    Dim varQoDShow As String = ""
    Dim varFCIFB As String = ""
    Dim varFCIFC As String = ""
    Dim varPIF As String = ""
    Dim varFBNew As String = ""
    Dim varFCNew As String = ""
    Dim varPNew As String = ""
    Dim varBF As String = ""
    Dim varSRF As String = ""
    Dim varTM As String = ""
    Dim varTemp As String = ""
    Dim varPA As String = ""
    Dim varTorque As String = ""
    Dim varToD As String = ""
    Dim varAbD As String = ""
    Dim varAbW As String = ""
    Dim varLdD As String = ""
    Dim varLdW As String = ""
    Dim varManTorq As String = ""
    Dim varManToD As String = ""
    Dim varManAbD As String = ""
    Dim varManAbW As String = ""
    Dim varManLdD As String = ""
    Dim varManLdW As String = ""
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim fname0 As String = "path.txt"
        Try
            Using fn0 As StreamReader = New StreamReader(fname0)
                lPath = fn0.ReadLine()
                oPath = fn0.ReadLine()
                bPath = fn0.ReadLine()
                iPath = fn0.ReadLine()
                xVer = fn0.ReadLine()
                fn0.Close()
            End Using
        Catch ex As Exception
            MsgBox("Path File Error")
        End Try
        Me.verLabel.Text = "Version " + xVer
        subGetFields()
        Me.lblDate.Text = Microsoft.VisualBasic.DateAndTime.Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now)
    End Sub
    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Application.Exit()
    End Sub
    Public Sub subLoadCombos()
        subGetAlt()
        subGetBirds()
        subGetHang()
        subGetSOF()
        subGetSup()
        subGetSR()
        subGetTab0()
        subGetTab1()
        its(0) = "Normal"
        its(1) = "Caution"
        its(2) = "Danger"
        MOA(0) = "Unrestricted"
        MOA(1) = "Restricted"
        MOA(2) = "Closed"
        Me.cbITS.Items.Clear()
        Me.cbMOASouth.Items.Clear()
        Me.cbMOATweet.Items.Clear()
        For x = 0 To 2
            cbITS.Items.Add(its(x))
            cbMOASouth.Items.Add(MOA(x))
            cbMOATweet.Items.Add(MOA(x))
        Next
    End Sub
    Private Sub subGetSOF()
        Dim nSof As Integer
        Me.cbSOF.Items.Clear()
        Dim fname0 As String = lPath + "sof.txt"
        Try
            Using fn0 As StreamReader = New StreamReader(fname0)
                nSof = CDec(fn0.ReadLine)
                For x = 1 To nSof
                    Me.cbSOF.Items.Add(fn0.ReadLine())
                Next x
                fn0.Close()
            End Using
        Catch ex As Exception
            If fnFNBR("sof.txt", bPath, lPath) Then
                subGetSOF()
            Else
                MsgBox("SoF File Read Error")
            End If
        End Try
    End Sub
    Private Sub subGetSup()
        Dim nSup As Integer
        Me.cbSUP.Items.Clear()
        Dim fname0 As String = lPath + "sup.txt"
        Try
            Using fn0 As StreamReader = New StreamReader(fname0)
                nSup = CDec(fn0.ReadLine)
                For x = 1 To nSup
                    Me.cbSUP.Items.Add(fn0.ReadLine())
                Next x
                fn0.Close()
            End Using
        Catch ex As Exception
            If fnFNBR("sup.txt", bPath, lPath) Then
                subGetSup()
            Else
                MsgBox("Sup File Read Error")
            End If
        End Try
    End Sub
    Private Sub subGetSR()
        Dim nSR As Integer
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
            lblSr1.Text = sr(1)
            lblSr2.Text = sr(2)
            lblSr3.Text = sr(3)
            lblSr4.Text = sr(4)
            lblSr5.Text = sr(5)
        Catch ex As Exception
            If fnFNBR("sr.txt", bPath, lPath) Then
                subGetSR()
            Else
                MsgBox("SR File Read Error")
            End If
        End Try
    End Sub
    Private Sub subGetHang()
        Dim nHang As Integer
        Me.cbHngStat.Items.Clear()
        Dim fname0 As String = lPath + "hangover.txt"
        Try
            Using fn0 As StreamReader = New StreamReader(fname0)
                nHang = CDec(fn0.ReadLine)
                For x = 1 To nHang
                    Dim iline As String = fn0.ReadLine()
                    Dim cols() As String = iline.Split(",")
                    stats(x, 0) = cols(0)
                    stats(x, 1) = cols(1)
                    stats(x, 2) = cols(2)
                    stats(x, 3) = cols(3)
                    stats(x, 4) = cols(4)
                    Me.cbHngStat.Items.Add(stats(x, 1))
                Next x
                fn0.Close()
            End Using
        Catch ex As Exception
            If fnFNBR("hangover.txt", bPath, lPath) Then
                subGetHang()
            Else
                MsgBox("Hangover File Read Error")
            End If
        End Try
    End Sub
    Private Sub subGetAlt()
        Dim nHang As Integer
        Me.cbHangAlt.Items.Clear()
        Dim fname0 As String = lPath + "alternate.txt"
        Try
            Using fn0 As StreamReader = New StreamReader(fname0)
                nHang = CDec(fn0.ReadLine)
                For x = 1 To nHang
                    Dim iline As String = fn0.ReadLine()
                    Dim cols() As String = iline.Split(",")
                    alts(x, 1) = cols(0)
                    alts(x, 2) = cols(1)
                    alts(x, 3) = cols(2)
                    Me.cbHangAlt.Items.Add(alts(x, 1))
                Next x
                fn0.Close()
            End Using
        Catch ex As Exception
            If fnFNBR("alternate.txt", bPath, lPath) Then
                subGetAlt()
            Else
                MsgBox("Alternate File Read Error")
            End If
        End Try
    End Sub
    Private Sub subGetBirds()
        Dim nBird As Integer
        Me.cbBirds.Items.Clear()
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
                    Me.cbBirds.Items.Add(birds(x, 1))
                Next x
                fn0.Close()
            End Using
        Catch ex As Exception
            If fnFNBR("birds.txt", bPath, lPath) Then
                subGetBirds()
            Else
                MsgBox("Bird Status File Read Error")
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
                MsgBox("Tab 0 File Read Error")
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
                subGetTab1()
            Else
                MsgBox("Tab 1 File Read Error")
            End If
        End Try
    End Sub
    Public Sub subBackupFiles()
        If Not fnFNBR("alternate.txt", lPath, bPath) Then
            MsgBox("Alternate Backup Error")
        End If
        If Not fnFNBR("birds.txt", lPath, bPath) Then
            MsgBox("Birds Backup Error")
        End If
        If Not fnFNBR("hangover.txt", lPath, bPath) Then
            MsgBox("Hangover Backup Error")
        End If
        If Not fnFNBR("path.txt", lPath, bPath) Then
            MsgBox("Path Backup Error")
        End If
        If Not fnFNBR("sof.txt", lPath, bPath) Then
            MsgBox("SoF Backup Error")
        End If
        If Not fnFNBR("status.txt", lPath, bPath) Then
            MsgBox("Status Backup Error")
        End If
        If Not fnFNBR("sr.txt", lPath, bPath) Then
            MsgBox("SR Backup Error")
        End If
        If Not fnFNBR("status2.txt", lPath, bPath) Then
            MsgBox("Status 2 Backup Error")
        End If
        If Not fnFNBR("status4a.txt", lPath, bPath) Then
            MsgBox("Status 4a Backup Error")
        End If
        If Not fnFNBR("status4b.txt", lPath, bPath) Then
            MsgBox("Status 4b Backup Error")
        End If
        If Not fnFNBR("status.xml", lPath, bPath) Then
            MsgBox("Status XML Backup Error")
        End If
        If Not fnFNBR("sup.txt", lPath, bPath) Then
            MsgBox("Sup Backup Error")
        End If
        If Not fnFNBR("tab0.txt", lPath, bPath) Then
            MsgBox("Tab 0 Backup Error")
        End If
        If Not fnFNBR("tab1.txt", lPath, bPath) Then
            MsgBox("Tab 1 Backup Error")
        End If
    End Sub
    Private Function fnFNBR(ByVal fn As String, ByVal path1 As String, ByVal path2 As String) As Boolean
        Try
            FileCopy(path1 & fn, path2 & fn)
            fnFNBR = True
        Catch ex As Exception
            fnFNBR = False
        End Try
    End Function
    Private Sub subResetFields()
        subLoadCombos()
        Me.rbHang14R.Checked = False
        Me.rbHang32L.Checked = False
        Me.lblDryWet.Text = "Dry"
        Me.tbNSFuel.Text = ""
        Me.rbAHCNo.Checked = False
        Me.rbAHCYes.Checked = False
        Me.rbNavRndUp.Checked = False
        Me.rbNavRndDown.Checked = False
        Me.rbNavIlsUp.Checked = False
        Me.rbNavIlsDown.Checked = False
        Me.rbNavDmeUp.Checked = False
        Me.rbNavDmeDown.Checked = False
        Me.rbNavSatUp.Checked = False
        Me.rbNavSatDown.Checked = False
        Me.rbNavSkfUp.Checked = False
        Me.rbNavSkfDown.Checked = False
        Me.rbNavSsfUp.Checked = False
        Me.rbNavSsfDown.Checked = False
        Me.rbSr1Open.Checked = False
        Me.rbSr1Closed.Checked = False
        Me.rbSr2Open.Checked = False
        Me.rbSr2Closed.Checked = False
        Me.rbSr3Open.Checked = False
        Me.rbSr3Closed.Checked = False
        Me.rbSr4Open.Checked = False
        Me.rbSr4Closed.Checked = False
        Me.rbSr5Open.Checked = False
        Me.rbSr5Closed.Checked = False
        Me.tbBravo.Text = ""
        Me.cbBNew.Checked = False
        Me.tbCharlie.Text = ""
        Me.cbCNew.Checked = False
        Me.tbPif.Text = ""
        Me.cbPNew.Checked = False
        Me.rbBfYes.Checked = False
        Me.rbBfNo.Checked = False
        Me.rbSRFYes.Checked = False
        Me.rbSRFNo.Checked = False
        Me.tbSupOne.Text = ""
        Me.tbSupTwo.Text = ""
        Me.tbSupThree.Text = ""
        Me.tbSupFour.Text = ""
        Me.tbSupFive.Text = ""
        Me.tbSupSix.Text = ""
        Me.tbSupSeven.Text = ""
        Me.tbQoD.Text = ""
        Me.rbQoDHide.Checked = False
        Me.rbQoDShow.Checked = False
        Me.tbTOPA.Text = 750
        Me.tbTOTemp.Text = 75
        Me.lblFC.Text = degF
        Me.lblTabErr.Visible = False
        Me.rbTMFor.Checked = False
        Me.rbTMMan.Checked = False
        Me.rbTMTab.Checked = False
        Me.tbManAbD.Text = 0
        Me.tbManAbW.Text = 0
        Me.tbManTOD.Text = 0
        Me.tbManLdD.Text = 0
        Me.tbManLdW.Text = 0
        Me.rbBeerOn.Checked = False
        Me.rbBeerOff.Checked = True
    End Sub
    Public Sub subGetFields()
        Dim o As String
        Dim settings As New XmlReaderSettings
        settings.IgnoreWhitespace = True
        settings.IgnoreComments = True
        Try
            Dim xmlIn As XmlReader = XmlReader.Create(lPath & "status.xml", settings)
            subResetFields()
            xmlIn.ReadStartElement("Status")
            varSOF = xmlIn.ReadElementString("SOF")
            For i = 0 To cbSOF.Items.Count - 1
                If cbSOF.Items(i) = varSOF Then
                    cbSOF.Text = cbSOF.Items(i)
                End If
            Next
            varSUP = xmlIn.ReadElementString("SUP")
            For i = 0 To cbSUP.Items.Count - 1
                If cbSUP.Items(i) = varSUP Then
                    cbSUP.Text = cbSUP.Items(i)
                End If
            Next
            varStatus = xmlIn.ReadElementString("Hang")
            For i = 0 To cbHngStat.Items.Count - 1
                If cbHngStat.Items(i) = varStatus Then
                    cbHngStat.Text = cbHngStat.Items(i)
                End If
            Next
            varHangRwy = xmlIn.ReadElementString("Rwy")
            If varHangRwy = "14R" Then
                rbHang14R.Checked = True
            Else
                rbHang32L.Checked = True
            End If
            varWetDry = xmlIn.ReadElementString("WetDry")
            lblDryWet.Text = varWetDry
            varAlternate = xmlIn.ReadElementString("Alt")
            For i = 0 To cbHangAlt.Items.Count - 1
                If cbHangAlt.Items(i) = varAlternate Then
                    cbHangAlt.Text = cbHangAlt.Items(i)
                End If
            Next
            varAltFuel = xmlIn.ReadElementString("AltF")
            If varAltFuel = "0" Or varAltFuel = "" Then
                tbNSFuel.Text = ""
            Else
                tbNSFuel.Text = varAltFuel
            End If
            varAHC = xmlIn.ReadElementString("AHC")
            If varAHC = "Yes" Then
                rbAHCYes.Checked = True
            Else
                rbAHCNo.Checked = True
            End If
            varA1B = xmlIn.ReadElementString("A1B")
            varA1E = xmlIn.ReadElementString("A1E")
            varA2B = xmlIn.ReadElementString("A2B")
            varA2E = xmlIn.ReadElementString("A2E")
            varA3B = xmlIn.ReadElementString("A3B")
            varA3E = xmlIn.ReadElementString("A3E")
            tbAHC1Be.Text = varA1B
            tbAHC1En.Text = varA1E
            tbAHC2Be.Text = varA2B
            tbAHC2En.Text = varA2E
            tbAHC3Be.Text = varA3B
            tbAHC3En.Text = varA3E
            varBirds = xmlIn.ReadElementString("Birds")
            For i = 0 To cbBirds.Items.Count - 1
                If cbBirds.Items(i) = varBirds Then
                    cbBirds.Text = cbBirds.Items(i)
                    varBirds = birds(i + 1, 1)
                    varBird1 = birds(i + 1, 2)
                    varBird2 = birds(i + 1, 3)
                    varBirdColor = birds(i + 1, 4)
                End If
            Next
            varITS = xmlIn.ReadElementString("Its")
            For i = 0 To cbITS.Items.Count - 1
                If cbITS.Items(i) = varITS Then
                    cbITS.Text = cbITS.Items(i)
                End If
            Next
            varMoaSouth = xmlIn.ReadElementString("MoaSouth")
            For i = 0 To cbMOASouth.Items.Count - 1
                If cbMOASouth.Items(i) = varMoaSouth Then
                    cbMOASouth.Text = cbMOASouth.Items(i)
                End If
            Next
            varMoaTweet = xmlIn.ReadElementString("MoaTweet")
            For i = 0 To cbMOATweet.Items.Count - 1
                If cbMOATweet.Items(i) = varMoaTweet Then
                    cbMOATweet.Text = cbMOATweet.Items(i)
                End If
            Next
            varMoaTalon = xmlIn.ReadElementString("MoaTalon")
            varSR1 = xmlIn.ReadElementString("SR1")
            If varSR1 = "Open" Then
                rbSr1Open.Checked = True
            Else
                rbSr1Closed.Checked = True
            End If
            varSR2 = xmlIn.ReadElementString("SR2")
            If varSR2 = "Open" Then
                rbSr2Open.Checked = True
            Else
                rbSr2Closed.Checked = True
            End If
            varSR3 = xmlIn.ReadElementString("SR3")
            If varSR3 = "Open" Then
                rbSr3Open.Checked = True
            Else
                rbSr3Closed.Checked = True
            End If
            varSR4 = xmlIn.ReadElementString("SR4")
            If varSR4 = "Open" Then
                rbSr4Open.Checked = True
            Else
                rbSr4Closed.Checked = True
            End If
            varSR5 = xmlIn.ReadElementString("SR5")
            If varSR5 = "Open" Then
                rbSr5Open.Checked = True
            Else
                rbSr5Closed.Checked = True
            End If
            varNavRND = xmlIn.ReadElementString("NavRnd")
            If varNavRND = "Up" Then
                rbNavRndUp.Checked = True
            Else
                rbNavRndDown.Checked = True
            End If
            varNavILS = xmlIn.ReadElementString("NavIls")
            If varNavILS = "Up" Then
                rbNavIlsUp.Checked = True
            Else
                rbNavIlsDown.Checked = True
            End If
            varNavDME = xmlIn.ReadElementString("NavDme")
            If varNavDME = "Up" Then
                rbNavDmeUp.Checked = True
            Else
                rbNavDmeDown.Checked = True
            End If
            varNavSAT = xmlIn.ReadElementString("NavSat")
            If varNavSAT = "Up" Then
                rbNavSatUp.Checked = True
            Else
                rbNavSatDown.Checked = True
            End If
            varNavSKF = xmlIn.ReadElementString("NavSkf")
            If varNavSKF = "Up" Then
                rbNavSkfUp.Checked = True
            Else
                rbNavSkfDown.Checked = True
            End If
            varNavSSF = xmlIn.ReadElementString("NavSsf")
            If varNavSSF = "Up" Then
                rbNavSsfUp.Checked = True
            Else
                rbNavSsfDown.Checked = True
            End If
            varOps1 = xmlIn.ReadElementString("SupOne")
            varOps2 = xmlIn.ReadElementString("SupTwo")
            varOps3 = xmlIn.ReadElementString("SupThree")
            varOps4 = xmlIn.ReadElementString("SupFour")
            varOps5 = xmlIn.ReadElementString("SupFive")
            varOps6 = xmlIn.ReadElementString("SupSix")
            varOps7 = xmlIn.ReadElementString("SupSeven")
            tbSupOne.Text = varOps1
            tbSupTwo.Text = varOps2
            tbSupThree.Text = varOps3
            tbSupFour.Text = varOps4
            tbSupFive.Text = varOps5
            tbSupSix.Text = varOps6
            tbSupSeven.Text = varOps7
            varQoD = xmlIn.ReadElementString("QoD")
            tbQoD.Text = varQoD
            varQoDShow = xmlIn.ReadElementString("QoDShow")
            If varQoDShow = "Show" Then
                rbQoDShow.Checked = True
            Else
                rbQoDHide.Checked = True
            End If
            varFCIFB = xmlIn.ReadElementString("Bravo")
            tbBravo.Text = varFCIFB
            varFBNew = xmlIn.ReadElementString("BrNew")
            If varFBNew = "Yes" Then
                cbBNew.Checked = True
            End If
            varFCIFC = xmlIn.ReadElementString("Charlie")
            tbCharlie.Text = varFCIFC
            varFCNew = xmlIn.ReadElementString("ChNew")
            If varFCNew = "Yes" Then
                cbCNew.Checked = True
            End If
            varPIF = xmlIn.ReadElementString("PIF")
            tbPif.Text = varPIF
            varPNew = xmlIn.ReadElementString("PIFNew")
            If varPNew = "Yes" Then
                cbPNew.Checked = True
            End If
            varBF = xmlIn.ReadElementString("BF")
            If varBF = "Yes" Then
                rbBfYes.Checked = True
            Else
                rbBfNo.Checked = True
            End If
            varSRF = xmlIn.ReadElementString("SRF")
            If varSRF = "Yes" Then
                rbSRFYes.Checked = True
            Else
                rbSRFNo.Checked = True
            End If
            varTemp = xmlIn.ReadElementString("ToldTemp")
            varPA = xmlIn.ReadElementString("ToldPA")
            tbTOTemp.Text = varTemp
            tbTOPA.Text = varPA
            varTM = xmlIn.ReadElementString("TM")
            If varTM = "Formula" Then
                rbTMFor.Checked = True
            ElseIf varTM = "Manual" Then
                rbTMMan.Checked = True
            Else
                rbTMTab.Checked = True
            End If
            varTorque = xmlIn.ReadElementString("TMTor")
            tbManTOD.Text = xmlIn.ReadElementString("TMToD")
            tbManAbD.Text = xmlIn.ReadElementString("TMAbD")
            tbManAbW.Text = xmlIn.ReadElementString("TMAbW")
            tbManLdD.Text = xmlIn.ReadElementString("TMLdD")
            tbManLdW.Text = xmlIn.ReadElementString("TMLdW")
            subUpdateTOLD()
            o = xmlIn.ReadElementString("Now")
            xmlIn.ReadEndElement()
            xmlIn.Close()
        Catch ex As Exception
            If fnFNBR("status.xml", bPath, lPath) Then
                subGetFields()
            Else
                MsgBox("Status XML File Read Error")
            End If
        End Try
        gblStatus = 0
        subFixButton(gblStatus)
    End Sub
    Private Sub subTOLDFor()
        Dim tempf As Integer = Val(varTemp)
        Dim pa As Integer = Val(Me.tbTOPA.Text)
        Me.lblForTOD.Text = Int((1292 + (0.4896 * tempf) + (0.041 * tempf ^ 2)) + ((750 - pa) / 10) + 0.5)
        Me.lblForAbD.Text = Int((128.44 + (-0.1421 * tempf) + (-0.0003 * tempf ^ 2)) + (((750 - pa) * 0.8) / 250) + 0.5)
        Me.lblForAbW.Text = Int((85.254 + (-0.0295 * tempf) + (-0.001 * tempf ^ 2)) + (((750 - pa) / 10) * 0.8) / 250 + 0.5)
        Me.lblForLdD.Text = Int((2372 + (5.4768 * tempf) + (-0.0034 * tempf ^ 2)) + ((750 - pa) / 10) + 0.5)
        Me.lblForLdW.Text = Int((3435.8 + (6.4293 * tempf) + (0.0085 * tempf ^ 2)) + ((750 - pa) / 10) + 0.5)
    End Sub
    Private Sub subTOLDTab()
        Dim tempf As Integer = Val(varTemp)
        Dim ratPA As Decimal = Val(Me.tbTOPA.Text) / 1000
        Dim ratTemp As Decimal = tempf / 5 - (tempf \ 5)
        Dim row1 As Integer = (tempf - 15) \ 5
        Dim row2 As Integer = row1 + 1
        Dim vTO, vAbD, vAbW, vLdD, vLdW, v1, v2 As Integer
        If (tempf \ 5) = (tempf / 5) Then
            vTO = arTab0(row1, 2) + (arTab1(row1, 2) - arTab0(row1, 2)) * ratPA
            vAbD = arTab0(row1, 3) + (arTab1(row1, 3) - arTab0(row1, 3)) * ratPA
            vAbW = arTab0(row1, 4) + (arTab1(row1, 4) - arTab0(row1, 4)) * ratPA
            vLdD = arTab0(row1, 5) + (arTab1(row1, 5) - arTab0(row1, 5)) * ratPA
            vLdW = arTab0(row1, 6) + (arTab1(row1, 6) - arTab0(row1, 6)) * ratPA
        Else
            v1 = arTab0(row1, 2) + (arTab1(row1, 2) - arTab0(row1, 2)) * ratPA
            v2 = arTab0(row2, 2) + (arTab1(row2, 2) - arTab0(row2, 2)) * ratPA
            vTO = v1 + (v2 - v1) * ratTemp
            v1 = arTab0(row1, 3) + (arTab1(row1, 3) - arTab0(row1, 3)) * ratPA
            v2 = arTab0(row2, 3) + (arTab1(row2, 3) - arTab0(row2, 3)) * ratPA
            vAbD = v1 + (v2 - v1) * ratTemp
            v1 = arTab0(row1, 4) + (arTab1(row1, 4) - arTab0(row1, 4)) * ratPA
            v2 = arTab0(row2, 4) + (arTab1(row2, 4) - arTab0(row2, 4)) * ratPA
            vAbW = v1 + (v2 - v1) * ratTemp
            v1 = arTab0(row1, 5) + (arTab1(row1, 5) - arTab0(row1, 5)) * ratPA
            v2 = arTab0(row2, 5) + (arTab1(row2, 5) - arTab0(row2, 5)) * ratPA
            vLdD = v1 + (v2 - v1) * ratTemp
            v1 = arTab0(row1, 6) + (arTab1(row1, 6) - arTab0(row1, 6)) * ratPA
            v2 = arTab0(row2, 6) + (arTab1(row2, 6) - arTab0(row2, 6)) * ratPA
            vLdW = v1 + (v2 - v1) * ratTemp
        End If
        Me.lblTabTOD.Text = vTO
        Me.lblTabAbD.Text = vAbD
        Me.lblTabAbW.Text = vAbW
        Me.lblTabLdD.Text = vLdD
        Me.lblTabLdW.Text = vLdW
    End Sub
    Private Sub subUpdateTOLD()
        If rbTMFor.Checked Then
            varToD = lblForTOD.Text
            varAbD = lblForAbD.Text
            varAbW = lblForAbW.Text
            varLdD = lblForLdD.Text
            varLdW = lblForLdW.Text
        ElseIf rbTMTab.Checked Then
            varToD = lblTabTOD.Text
            varAbD = lblTabAbD.Text
            varAbW = lblTabAbW.Text
            varLdD = lblTabLdD.Text
            varLdW = lblTabLdW.Text
        Else
            varToD = tbManTOD.Text
            varAbD = tbManAbD.Text
            varAbW = tbManAbW.Text
            varLdD = tbManLdD.Text
            varLdW = tbManLdW.Text
        End If
    End Sub
    Private Sub lblFC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblFC.Click
        If lblFC.Text = degF Then
            lblFC.Text = degC
            tbTOTemp.Text = Int(((tbTOTemp.Text - 32) / 1.8) * 10) / 10
        Else
            lblFC.Text = degF
            tbTOTemp.Text = Int(tbTOTemp.Text * 1.8 + 32.5)
        End If
    End Sub
    Private Function fnFixValue(ByVal werd As String)
        Dim temp As String = ""
        For x = 1 To Len(werd)
            If Mid(werd, x, 1) >= "0" And Mid(werd, x, 1) <= "9" Then
                temp += Mid(werd, x, 1)
            End If
        Next
        fnFixValue = temp
    End Function
    Private Sub tbTOTemp_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbTOTemp.TextChanged
        tbTOTemp.Text = fnFixValue(tbTOTemp.Text)
        If lblFC.Text = degC Then
            If tbTOTemp.Text <> "" Then
                varTemp = tbTOTemp.Text * 1.8 + 32
            End If
        Else
            varTemp = tbTOTemp.Text
        End If
        If varTemp <> "" Then
            If Val(varTemp) > 19 And Val(varTemp) < 101 Then
                subTOLDFor()
                subTOLDTab()
                lblTabErr.Visible = False
                rbTMTab.Enabled = True
                lblTabTOD.Visible = True
                lblTabAbD.Visible = True
                lblTabAbW.Visible = True
                lblTabLdD.Visible = True
                lblTabLdW.Visible = True
                Label59.Visible = True
                Label60.Visible = True
                Label61.Visible = True
                Label62.Visible = True
                Label63.Visible = True
            Else
                subTOLDFor()
                lblTabErr.Visible = True
                rbTMMan.Checked = False
                rbTMTab.Enabled = False
                lblTabTOD.Visible = False
                lblTabAbD.Visible = False
                lblTabAbW.Visible = False
                lblTabLdD.Visible = False
                lblTabLdW.Visible = False
                Label59.Visible = False
                Label60.Visible = False
                Label61.Visible = False
                Label62.Visible = False
                Label63.Visible = False
            End If
        End If
        subUpdateTOLD()
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub tbTOPA_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbTOPA.TextChanged
        tbTOPA.Text = fnFixValue(tbTOPA.Text)
        varPA = tbTOPA.Text
        If varPA <> "" Then
            If Val(varTemp) > 19 And Val(varTemp) < 101 Then
                subTOLDFor()
                subTOLDTab()
                lblTabErr.Visible = False
            Else
                subTOLDFor()
                lblTabErr.Visible = True
                rbTMMan.Checked = True
            End If
        End If
        subUpdateTOLD()
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub lblDryWet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblDryWet.Click
        If lblDryWet.Text = "Dry" Then
            lblDryWet.Text = "Wet"
            varWetDry = "Wet"
        ElseIf lblDryWet.Text = "Wet" Then
            lblDryWet.Text = "Standing Water"
            varWetDry = "Standing Water"
        ElseIf lblDryWet.Text = "Standing Water" Then
            lblDryWet.Text = "Patchy Standing Water"
            varWetDry = "Patchy Std Water"
        Else
            lblDryWet.Text = "Dry"
            varWetDry = "Dry"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub subFixButton(ByVal x As Integer)
        If x = 1 Then
            btnSave.Text = "Save Status"
            btnSave.ForeColor = Color.Red
            Me.btnSave.Cursor = Cursors.Arrow
            Me.btnSave.Refresh()
        ElseIf x = 0 Then
            btnSave.Text = "Status Saved"
            btnSave.ForeColor = Color.Green
            Me.btnSave.Cursor = Cursors.Arrow
            Me.btnSave.Refresh()
        ElseIf x = 3 Then
            btnSave.Text = "Saving..."
            btnSave.ForeColor = Color.SaddleBrown
            Me.btnSave.Cursor = Cursors.WaitCursor
            Me.btnSave.Refresh()
        End If
    End Sub
    Private Sub cbSOF_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSOF.SelectedIndexChanged
        varSOF = cbSOF.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub cbSUP_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSUP.SelectedIndexChanged
        varSUP = cbSUP.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbBeerOn_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbBeerOn.CheckedChanged
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbBeerOff_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbBeerOff.CheckedChanged
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbHang14R_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbHang14R.CheckedChanged
        If rbHang14R.Checked Then
            varHangRwy = "14R"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbHang32L_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbHang32L.CheckedChanged
        If rbHang32L.Checked Then
            varHangRwy = "32L"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub cbHngStat_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbHngStat.SelectedIndexChanged
        Dim x As Integer = cbHngStat.SelectedIndex + 1
        varStatus = stats(x, 1)
        varStat1 = stats(x, 2)
        varStat2 = stats(x, 3)
        varStatCol = stats(x, 4)
        If varStat2 = "" Then
            varStat2 = varStat1
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub cbHangAlt_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbHangAlt.SelectedIndexChanged
        varAlternate = cbHangAlt.Text
        gblStatus = 1
        subFixButton(gblStatus)
        Me.tbNSFuel.Enabled = True
        Me.Label49.Enabled = True
        Me.Label50.Enabled = True
        If Me.cbHangAlt.Text = "None" Then
            tbNSFuel.Text = ""
            Me.tbNSFuel.Enabled = False
            Me.Label49.Enabled = False
            Me.Label50.Enabled = False
        End If
    End Sub
    Private Sub tbNSFuel_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbNSFuel.TextChanged
        varAltFuel = tbNSFuel.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbAHCYes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbAHCYes.CheckedChanged
        If rbAHCYes.Checked Then
            varAHC = "Yes"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbAHCNo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbAHCNo.CheckedChanged
        If rbAHCNo.Checked Then
            varAHC = "No"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub tbAHC1Be_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbAHC1Be.TextChanged
        varA1B = tbAHC1Be.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub tbAHC1En_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbAHC1En.TextChanged
        varA1E = tbAHC1En.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub tbAHC2Be_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbAHC2Be.TextChanged
        varA2B = tbAHC2Be.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub tbAHC2En_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbAHC2En.TextChanged
        varA2E = tbAHC2En.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub tbAHC3Be_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbAHC3Be.TextChanged
        varA3B = tbAHC3Be.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub tbAHC3En_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbAHC3En.TextChanged
        varA3E = tbAHC3En.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub cbBirds_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbBirds.SelectedIndexChanged
        varBirds = birds(cbBirds.SelectedIndex + 1, 1)
        varBird1 = birds(cbBirds.SelectedIndex + 1, 2)
        varBird2 = birds(cbBirds.SelectedIndex + 1, 3)
        varBirdColor = birds(cbBirds.SelectedIndex + 1, 4)
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub cbITS_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbITS.SelectedIndexChanged
        varITS = its(cbITS.SelectedIndex)
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub cbMOASouth_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbMOASouth.SelectedIndexChanged
        varMoaSouth = MOA(cbMOASouth.SelectedIndex)
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub cbMOATweet_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbMOATweet.SelectedIndexChanged
        varMoaTweet = MOA(cbMOATweet.SelectedIndex)
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub tbBravo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbBravo.TextChanged
        varFCIFB = tbBravo.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub cbBNew_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbBNew.CheckedChanged
        If cbBNew.Checked Then
            varFBNew = "Yes"
        Else
            varFBNew = "No"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub tbCharlie_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbCharlie.TextChanged
        varFCIFC = tbCharlie.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub cbCNew_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbCNew.CheckedChanged
        If cbCNew.Checked Then
            varFCNew = "Yes"
        Else
            varFCNew = "No"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub tbPif_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbPif.TextChanged
        varPIF = tbPif.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub cbPNew_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbPNew.CheckedChanged
        If cbPNew.Checked Then
            varPNew = "Yes"
        Else
            varPNew = "No"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbBfYes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbBfYes.CheckedChanged
        If rbBfYes.Checked Then
            varBF = "Yes"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbBfNo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbBfNo.CheckedChanged
        If rbBfNo.Checked Then
            varBF = "No"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbSRFYes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSRFYes.CheckedChanged
        If rbSRFYes.Checked Then
            varSRF = "Yes"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbSRFNo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSRFNo.CheckedChanged
        If rbSRFNo.Checked Then
            varSRF = "No"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbNavRndUp_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbNavRndUp.CheckedChanged
        If rbNavRndUp.Checked Then
            varNavRND = "Up"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbNavRndDown_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbNavRndDown.CheckedChanged
        If rbNavRndDown.Checked Then
            varNavRND = "Down"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbNavIlsUp_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbNavIlsUp.CheckedChanged
        If rbNavIlsUp.Checked Then
            varNavILS = "Up"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbNavIlsDown_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbNavIlsDown.CheckedChanged
        If rbNavIlsDown.Checked Then
            varNavILS = "Down"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbNavDmeUp_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbNavDmeUp.CheckedChanged
        If rbNavDmeUp.Checked Then
            varNavDME = "Up"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbNavDmeDown_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbNavDmeDown.CheckedChanged
        If rbNavDmeDown.Checked Then
            varNavDME = "Down"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbNavSatUp_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbNavSatUp.CheckedChanged
        If rbNavSatUp.Checked Then
            varNavSAT = "Up"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbNavSatDown_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbNavSatDown.CheckedChanged
        If rbNavSatDown.Checked Then
            varNavSAT = "Down"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbNavSkfUp_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbNavSkfUp.CheckedChanged
        If rbNavSkfUp.Checked Then
            varNavSKF = "Up"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbNavSkfDown_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbNavSkfDown.CheckedChanged
        If rbNavSkfDown.Checked Then
            varNavSKF = "Down"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbNavSsfUp_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbNavSsfUp.CheckedChanged
        If rbNavSsfUp.Checked Then
            varNavSSF = "Up"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbNavSsfDown_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbNavSsfDown.CheckedChanged
        If rbNavSsfDown.Checked Then
            varNavSSF = "Up"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbSr1Open_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSr1Open.CheckedChanged
        If rbSr1Open.Checked Then
            varSR1 = "Open"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbSr1Closed_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSr1Closed.CheckedChanged
        If rbSr1Closed.Checked Then
            varSR1 = "Closed"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbSr2Open_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSr2Open.CheckedChanged
        If rbSr2Open.Checked Then
            varSR2 = "Open"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbSr2Closed_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSr2Closed.CheckedChanged
        If rbSr2Closed.Checked Then
            varSR2 = "Closed"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbSr3Open_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSr3Open.CheckedChanged
        If rbSr3Open.Checked Then
            varSR3 = "Open"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbSr3Closed_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSr3Closed.CheckedChanged
        If rbSr3Closed.Checked Then
            varSR3 = "Closed"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbSr4Open_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSr4Open.CheckedChanged
        If rbSr4Open.Checked Then
            varSR4 = "Open"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbSr4Closed_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSr4Closed.CheckedChanged
        If rbSr4Closed.Checked Then
            varSR4 = "Closed"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbSr5Open_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSr5Open.CheckedChanged
        If rbSr5Open.Checked Then
            varSR5 = "Open"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbSr5Closed_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSr5Closed.CheckedChanged
        If rbSr5Closed.Checked Then
            varSR5 = "Closed"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub tbSupOne_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbSupOne.TextChanged
        varOps1 = tbSupOne.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub tbSupTwo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbSupTwo.TextChanged
        varOps2 = tbSupTwo.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub tbSupThree_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbSupThree.TextChanged
        varOps3 = tbSupThree.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub tbSupFour_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbSupFour.TextChanged
        varOps4 = tbSupFour.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub tbSupFive_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbSupFive.TextChanged
        varOps5 = tbSupFive.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub tbSupSix_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbSupSix.TextChanged
        varOps6 = tbSupSix.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub tbSupSeven_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbSupSeven.TextChanged
        varOps7 = tbSupSeven.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbQoDShow_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbQoDShow.CheckedChanged
        If rbQoDShow.Checked Then
            varQoDShow = "Show"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbQoDHide_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbQoDHide.CheckedChanged
        If rbQoDHide.Checked Then
            varQoDShow = "Hide"
        End If
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub tbQoD_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbQoD.TextChanged
        varQoD = tbQoD.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub tbManTOD_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbManTOD.TextChanged
        varToD = tbManTOD.Text
        varManToD = tbManTOD.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub tbManAbD_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbManAbD.TextChanged
        varAbD = tbManAbD.Text
        varManAbD = tbManAbD.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub tbManAbW_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbManAbW.TextChanged
        varAbW = tbManAbW.Text
        varManAbW = tbManAbW.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub tbManLdD_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbManLdD.TextChanged
        varLdD = tbManLdD.Text
        varManLdD = tbManLdD.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub tbManLdW_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbManLdW.TextChanged
        varLdW = tbManLdW.Text
        varManLdW = tbManLdW.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbTMTab_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbTMTab.CheckedChanged
        varTM = "Tab Data"
        tbManAbD.Enabled = False
        tbManAbW.Enabled = False
        tbManTOD.Enabled = False
        tbManLdD.Enabled = False
        tbManLdW.Enabled = False
        varAbD = lblTabAbD.Text
        varAbW = lblTabAbW.Text
        varLdD = lblTabLdD.Text
        varLdW = lblTabLdW.Text
        varToD = lblTabTOD.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbTMFor_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbTMFor.CheckedChanged
        varTM = "Formula"
        tbManAbD.Enabled = False
        tbManAbW.Enabled = False
        tbManTOD.Enabled = False
        tbManLdD.Enabled = False
        tbManLdW.Enabled = False
        varAbD = lblForAbD.Text
        varAbW = lblForAbW.Text
        varLdD = lblForLdD.Text
        varLdW = lblForLdW.Text
        varToD = lblForTOD.Text
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub rbTMMan_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbTMMan.CheckedChanged
        varTM = "Manual"
        tbManAbD.Enabled = True
        tbManAbW.Enabled = True
        tbManTOD.Enabled = True
        tbManLdD.Enabled = True
        tbManLdW.Enabled = True
        varAbD = tbManAbD.Text
        varAbW = tbManAbW.Text
        varLdD = tbManLdD.Text
        varLdW = tbManLdW.Text
        varToD = tbManTOD.Text
        varManAbD = varAbD
        varManAbW = varAbW
        varManLdD = varLdD
        varManLdW = varLdW
        varManToD = varToD
        gblStatus = 1
        subFixButton(gblStatus)
    End Sub
    Private Sub subWrite(ByVal iFile As Integer, ByVal iSpc As Integer, ByVal iTop As Integer, ByVal iLeft As Integer, _
                            ByVal iSize As Integer, ByVal sColor As String, ByVal iWidth As Integer, _
                            ByVal iBold As Integer, ByVal sBorder As String, ByVal sValue As String)
        Dim str As String = ""
        str = "<div style=""position: absolute; top: " & CStr(iTop) & "px; left: " & CStr(iLeft) & "px"
        If iSize > 0 Then
            str &= "; font-size: " & CStr(iSize) & "px"
        End If
        If sColor <> "0" Then
            str &= "; color: " & sColor
        Else
            str &= "; color: #000"
        End If
        If iWidth > 0 Then
            str &= "; width: " & CStr(iWidth) & "px; text-align: center"
        End If
        If iBold > 0 Then
            str &= "; font-weight: bold"
        End If
        If sBorder <> "0" Then
            str &= "; border: solid #763900 1px; background-color: " & sBorder & ""
        End If
        str &= """>" & sValue & "</div>"
        PrintLine(iFile, SPC(iSpc), str)
    End Sub
    Sub subWriteXML()
        Dim settings As New XmlWriterSettings
        settings.Indent = True
        settings.IndentChars = "  "
        Dim xmlOut As XmlWriter = XmlWriter.Create(lPath & "status.xml", settings)
        xmlOut.WriteStartDocument()
        xmlOut.WriteStartElement("Status")
        xmlOut.WriteElementString("SOF", varSOF)
        xmlOut.WriteElementString("SUP", varSUP)
        xmlOut.WriteElementString("Hang", varStatus)
        xmlOut.WriteElementString("Rwy", varHangRwy)
        xmlOut.WriteElementString("WetDry", varWetDry)
        xmlOut.WriteElementString("Alt", varAlternate)
        xmlOut.WriteElementString("AltF", varAltFuel)
        xmlOut.WriteElementString("AHC", varAHC)
        xmlOut.WriteElementString("A1B", varA1B)
        xmlOut.WriteElementString("A1E", varA1E)
        xmlOut.WriteElementString("A2B", varA2B)
        xmlOut.WriteElementString("A2E", varA2E)
        xmlOut.WriteElementString("A3B", varA3B)
        xmlOut.WriteElementString("A3E", varA3E)
        xmlOut.WriteElementString("Birds", varBirds)
        xmlOut.WriteElementString("Its", varITS)
        xmlOut.WriteElementString("MoaSouth", varMoaSouth)
        xmlOut.WriteElementString("MoaTweet", varMoaTweet)
        xmlOut.WriteElementString("MoaTalon", varMoaTalon)
        xmlOut.WriteElementString("SR1", varSR1)
        xmlOut.WriteElementString("SR2", varSR2)
        xmlOut.WriteElementString("SR3", varSR3)
        xmlOut.WriteElementString("SR4", varSR4)
        xmlOut.WriteElementString("SR5", varSR5)
        xmlOut.WriteElementString("NavRnd", varNavRND)
        xmlOut.WriteElementString("NavIls", varNavILS)
        xmlOut.WriteElementString("NavDme", varNavDME)
        xmlOut.WriteElementString("NavSat", varNavSAT)
        xmlOut.WriteElementString("NavSkf", varNavSKF)
        xmlOut.WriteElementString("NavSsf", varNavSSF)
        xmlOut.WriteElementString("SupOne", varOps1)
        xmlOut.WriteElementString("SupTwo", varOps2)
        xmlOut.WriteElementString("SupThree", varOps3)
        xmlOut.WriteElementString("SupFour", varOps4)
        xmlOut.WriteElementString("SupFive", varOps5)
        xmlOut.WriteElementString("SupSix", varOps6)
        xmlOut.WriteElementString("SupSeven", varOps7)
        xmlOut.WriteElementString("QoD", varQoD)
        xmlOut.WriteElementString("QoDShow", varQoDShow)
        xmlOut.WriteElementString("Bravo", varFCIFB)
        xmlOut.WriteElementString("BrNew", varFBNew)
        xmlOut.WriteElementString("Charlie", varFCIFC)
        xmlOut.WriteElementString("ChNew", varFCNew)
        xmlOut.WriteElementString("PIF", varPIF)
        xmlOut.WriteElementString("PIFNew", varPNew)
        xmlOut.WriteElementString("BF", varBF)
        xmlOut.WriteElementString("SRF", varSRF)
        xmlOut.WriteElementString("ToldTemp", varTemp)
        xmlOut.WriteElementString("ToldPA", varPA)
        xmlOut.WriteElementString("TM", varTM)
        xmlOut.WriteElementString("TMTor", varManTorq)
        xmlOut.WriteElementString("TMToD", varManToD)
        xmlOut.WriteElementString("TMAbD", varManAbD)
        xmlOut.WriteElementString("TMAbW", varManAbW)
        xmlOut.WriteElementString("TMLdD", varManLdD)
        xmlOut.WriteElementString("TMLdW", varManLdW)
        xmlOut.WriteElementString("Now", Now())
        xmlOut.WriteEndElement()
        xmlOut.Close()
    End Sub
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        gblStatus = 3
        subFixButton(gblStatus)
        Dim j As String = ""
        Dim k As String = ""
        Dim l As String = ""
        Dim m As String = ""
        Dim x As Integer = 0
        Dim grn As String = "images\greenball.jpg"
        Dim yel As String = "images\yellowball.jpg"
        Dim red As String = "images\redball.jpg"
        Dim grnb As String = "images\green100.gif"
        Dim yelb As String = "images\yellow100.gif"
        Dim redb As String = "images\red100.gif"
        Dim n As String = " " & Chr(149) & " "
        Dim cr As String = "<br />"
        FileCopy(lPath & "status.txt", lPath & "status.html")
        FileOpen(1, lPath & "status.html", OpenMode.Append)
        FileCopy(lPath & "status2.txt", lPath & "status2.html")
        FileOpen(2, lPath & "status2.html", OpenMode.Append)
        FileCopy(lPath & "status4a.txt", lPath & "status4a.html")
        FileOpen(4, lPath & "status4a.html", OpenMode.Append)
        FileCopy(lPath & "status4b.txt", lPath & "status4b.html")
        If (varStatus = "Standby") Or (varStatus = "Stop Launch") Then
            PrintLine(1, SPC(6), "<div style=""position: absolute; top: 340px; left: 296px""><img src=""images\t6ani7.jpg""></div>")
            PrintLine(2, SPC(6), "<div style=""position: absolute; top: 290px; left: 309px""><img src=""images\t6ani7.jpg""></div>")
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 600px; left:  22px""><img src=""images\t6ani7.jpg"" width=""380px""></div>")
        Else
            PrintLine(1, SPC(6), "<div style=""position: absolute; top: 340px; left: 296px""><img src=""images\t6anil3.gif""></div>")
            PrintLine(2, SPC(6), "<div style=""position: absolute; top: 290px; left: 309px""><img src=""images\t6anil3.gif""></div>")
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 600px; left:  22px""><img src=""images\t6anil3.gif"" width=""380px""></div>")
        End If
        PrintLine(1, "<!-- SOF/SUP Start -->")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 141px; left:  -4px""><img src=""images\sof.jpg""></div>")
        PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 140px; left:  -5px; width: 760px; height:  30px;""></div>")
        PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 141px; left:  -4px; width: 758px; height:  28px; border-color: #c5885f""></div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 143px; left:  80px; font-size: 22px; color: #763900; font-weight: bold"">SOF</div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 143px; left: 480px; font-size: 22px; color: #763900; font-weight: bold"">SUP</div>")
        PrintLine(2, "<!-- SOF/SUP Start -->")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top:  91px; left:  -4px""><img src=""images\sof2.jpg"" width=""1106px"" height=""37px""></div>")
        PrintLine(2, SPC(6), "<div id=""boxes"" style=""top:  90px; left:  -5px; width: 1100px; height:  30px;""></div>")
        PrintLine(2, SPC(6), "<div id=""boxes"" style=""top:  91px; left:  -4px; width: 1098px; height:  28px; border-color: #c5885f""></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top:  93px; left:  80px; font-size: 22px; color: #763900; font-weight: bold"">SOF</div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top:  93px; left: 820px; font-size: 22px; color: #763900; font-weight: bold"">SUP</div>")
        PrintLine(4, "<!-- SOF/SUP Start -->")
        PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 100px; left:  -5px; width: 1080px; height:  50px;""></div>")
        PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 101px; left:  -4px; width: 1078px; height:  48px; border-color: #c5885f""></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 102px; left:  20px; font-size: 40px; color: #763900; font-weight: bold"">SOF</div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 102px; left: 565px; font-size: 40px; color: #763900; font-weight: bold"">SUP</div>")
        subWrite(1, 6, 141, 135, 24, "#000", 0, 1, 0, varSOF)
        subWrite(1, 6, 141, 535, 24, "#000", 0, 1, 0, varSUP)
        subWrite(2, 6, 91, 135, 24, "#000", 0, 1, 0, varSOF)
        subWrite(2, 6, 91, 875, 24, "#000", 0, 1, 0, varSUP)
        subWrite(4, 6, 100, 115, 44, "#000", 0, 1, 0, varSOF)
        subWrite(4, 6, 100, 660, 44, "#000", 0, 1, 0, varSUP)
        PrintLine(1, "<!-- SOF/SUP End -->")
        PrintLine(2, "<!-- SOF/SUP End -->")
        PrintLine(4, "<!-- SOF/SUP End -->")
        PrintLine(1, SPC(6), "<script type=""text/javascript"">")
        PrintLine(1, SPC(6), "<!--")
        PrintLine(1, SPC(6), "function changeColour(elementId) {")
        PrintLine(2, SPC(6), "<script type=""text/javascript"">")
        PrintLine(2, SPC(6), "<!--")
        PrintLine(2, SPC(6), "function changeColour(elementId) {")
        PrintLine(4, SPC(6), "<script type=""text/javascript"">")
        PrintLine(4, SPC(6), "<!--")
        PrintLine(4, SPC(6), "function changeColour(elementId) {")
        PrintLine(1, SPC(8), "var interval = 1500;")
        PrintLine(1, SPC(8), "var tex1 = """ & varStat1 & """, tex2 = """ & varStat2 & """;")
        PrintLine(1, SPC(8), "var texb1 = ""Boldface Due"", texb2 = """";")
        PrintLine(1, SPC(8), "var texs1 = ""Safety Rd File"", texs2 = """";")
        PrintLine(1, SPC(8), "var texm1 = """ & varBird1 & """, texm2 = """ & varBird2 & """;")
        PrintLine(1, SPC(8), "if (document.getElementById) {")
        PrintLine(1, SPC(10), "var element = document.getElementById(elementId);")
        PrintLine(2, SPC(8), "var interval = 1500;")
        PrintLine(2, SPC(8), "var tex1 = """ & varStat1 & """, tex2 = """ & varStat2 & """;")
        PrintLine(2, SPC(8), "var texb1 = ""Boldface Due"", texb2 = """";")
        PrintLine(2, SPC(8), "var texs1 = ""Safety Rd File"", texs2 = """";")
        PrintLine(2, SPC(8), "var texm1 = """ & varBird1 & """, texm2 = """ & varBird2 & """;")
        PrintLine(2, SPC(8), "if (document.getElementById) {")
        PrintLine(2, SPC(10), "var element = document.getElementById(elementId);")
        PrintLine(4, SPC(8), "var interval = 1500;")
        PrintLine(4, SPC(8), "var tex1 = """ & varStat1 & """, tex2 = """ & varStat2 & """;")
        PrintLine(4, SPC(8), "var texb1 = ""Boldface Due"", texb2 = """";")
        PrintLine(4, SPC(8), "var texs1 = ""Safety Rd File"", texs2 = """";")
        PrintLine(4, SPC(8), "var texm1 = """ & varBird1 & """, texm2 = """ & varBird2 & """;")
        PrintLine(4, SPC(8), "if (document.getElementById) {")
        PrintLine(4, SPC(10), "var element = document.getElementById(elementId);")
        If rbBfYes.Checked Then
            PrintLine(1, SPC(10), "document.all.bf.innerHTML = (document.all.bf.innerHTML == texb1) ? texb2 : texb1;")
            PrintLine(2, SPC(10), "document.all.bf.innerHTML = (document.all.bf.innerHTML == texb1) ? texb2 : texb1;")
            PrintLine(4, SPC(10), "document.all.bf.innerHTML = (document.all.bf.innerHTML == texb1) ? texb2 : texb1;")
        End If
        If rbSRFYes.Checked Then
            PrintLine(1, SPC(10), "document.all.srf.innerHTML = (document.all.srf.innerHTML == texs1) ? texs2 : texs1;")
            PrintLine(2, SPC(10), "document.all.srf.innerHTML = (document.all.srf.innerHTML == texs1) ? texs2 : texs1;")
            PrintLine(4, SPC(10), "document.all.srf.innerHTML = (document.all.srf.innerHTML == texs1) ? texs2 : texs1;")
        End If
        PrintLine(1, SPC(10), "document.all.hk.innerHTML = (document.all.hk.innerHTML == tex1) ? tex2 : tex1;")
        PrintLine(2, SPC(10), "document.all.hk.innerHTML = (document.all.hk.innerHTML == tex1) ? tex2 : tex1;")
        PrintLine(4, SPC(10), "document.all.hk.innerHTML = (document.all.hk.innerHTML == tex1) ? tex2 : tex1;")
        PrintLine(1, SPC(10), "document.all.bird.innerHTML = (document.all.bird.innerHTML == texm1) ? texm2 : texm1;")
        PrintLine(2, SPC(10), "document.all.bird.innerHTML = (document.all.bird.innerHTML == texm1) ? texm2 : texm1;")
        PrintLine(4, SPC(10), "document.all.bird.innerHTML = (document.all.bird.innerHTML == texm1) ? texm2 : texm1;")
        PrintLine(1, SPC(10), "setTimeout(""changeColour('"" + elementId + ""')"", interval);")
        PrintLine(1, SPC(8), "}")
        PrintLine(2, SPC(10), "setTimeout(""changeColour('"" + elementId + ""')"", interval);")
        PrintLine(2, SPC(8), "}")
        PrintLine(4, SPC(10), "setTimeout(""changeColour('"" + elementId + ""')"", interval);")
        PrintLine(4, SPC(8), "}")
        PrintLine(1, SPC(6), "}")
        PrintLine(2, SPC(6), "}")
        PrintLine(4, SPC(6), "}")
        PrintLine(1, SPC(6), "  //--></script>")
        PrintLine(2, SPC(6), "  //--></script>")
        PrintLine(4, SPC(6), "  //--></script>")
        PrintLine(1, "")
        PrintLine(1, "<!-- Hangover Start -->")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 181px; left:  -4px;""><img src=""images\hang2.jpg"" width=""307px"" height=""217px""></div>")
        PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 180px; left:  -5px; width: 300px; height: 210px;""></div>")
        PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 181px; left:  -4px; width: 298px; height: 208px; border-color: #c5885f""></div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 186px; left:   0px; font-size: 22px; color: #763900; font-weight: bold""><u>Hangover</u></div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 223px; left:   0px; font-size: 20px; color: #060; font-weight: bold"">Runway</div>")
        PrintLine(2, "")
        PrintLine(2, "<!-- Hangover Start -->")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 131px; left:  -4px""><img src=""images\hang2.jpg""></div>")
        PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 130px; left:  -5px; width: 300px; height: 235px;""></div>")
        PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 131px; left:  -4px; width: 298px; height: 233px; border-color: #c5885f""></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 136px; left:   0px; font-size: 22px; color: #763900; font-weight: bold""><u>Hangover</u></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 173px; left:   0px; font-size: 20px; color: #060; font-weight: bold"">Runway</div>")
        PrintLine(4, "")
        PrintLine(4, "<!-- Hangover Start -->")
        PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 160px; left:  -5px; width: 435px; height: 435px;""></div>")
        PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 161px; left:  -4px; width: 433px; height: 433px; border-color: #c5885f""></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 166px; left:   0px; font-size: 32px; width: 430px; text-align: center; color: #763900; font-weight: bold""><u>Hangover</u></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 290px; left:   0px; font-size: 30px; color: #060; font-weight: bold"">Runway</div>")
        If varStatCol = "Green" Then
            l = "images\green180.gif"
        ElseIf varStatCol = "Yellow" Then
            l = "images\yellow180.gif"
        Else
            l = "images\red180.gif"
        End If
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 185px; left: 105px""><img src=""" & l & """ alt=""Hangover status is " & varStat2 & """></div>")
        PrintLine(1, SPC(6), "<div id=""hk"" style=""position: absolute; top: 188px; left: 105px; font-size: 22px; font-weight: bold; width: 180px; text-align: center"">" & m & "</div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 135px; left: 105px""><img src=""" & l & """ alt=""Hangover status is " & varStat2 & """></div>")
        PrintLine(2, SPC(6), "<div id=""hk"" style=""position: absolute; top: 137px; left: 105px; font-size: 22px; font-weight: bold; width: 180px; text-align: center"">" & m & "</div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 205px; left:  10px""><img src=""" & l & """ alt=""Hangover status is " & varStat2 & """ width=""420px""></div>")
        PrintLine(4, SPC(6), "<div id=""hk"" style=""position: absolute; top: 210px; left:  10px; font-size: 48px; font-weight: bold; width: 405px; text-align: center"">" & m & "</div>")
        If rbHang14R.Checked Then
            m = varHangRwy & " • " & varWetDry
        ElseIf rbHang32L.Checked Then
            m = varHangRwy & " • " & varWetDry
        End If
        Dim irc1 As Integer
        Dim irc2 As Integer
        Dim t1 As Integer
        Dim t2 As Integer
        Dim t4 As Integer
        If varWetDry = "Dry" Or varWetDry = "Wet" Then
            irc1 = 22
            irc2 = 40
            t1 = 223
            t2 = 173
            t4 = 289
        ElseIf varWetDry = "Standing Water" Then
            irc1 = 17
            irc2 = 32
            t1 = 228
            t2 = 178
            t4 = 289
        Else
            irc1 = 16
            irc2 = 32
            t1 = 228
            t2 = 178
            t4 = 289
        End If
        subWrite(1, 6, t1, 109, irc1, "#000", 180, 1, 0, m)
        subWrite(2, 6, t2, 109, irc1, "#000", 180, 1, 0, m)
        subWrite(4, 6, t4, 140, irc2, "#000", 280, 1, 0, m)
        If varAlternate <> "None" Then
            Do Until alts(x, 1) = varAlternate
                x += 1
            Loop
            Dim altsfuel As String = alts(x, 3)
            If varAltFuel <> "" And varAltFuel <> "0" Then
                altsfuel = varAltFuel
            End If
            subWrite(1, 6, 251, 0, 20, "#060", 0, 1, 0, "Alternate")
            subWrite(1, 6, 279, 0, 20, "#060", 0, 1, 0, "Alt/Fuel")
            subWrite(1, 6, 251, 105, 22, "#000", 180, 1, 0, varAlternate)
            PrintLine(1, SPC(6), "<div style=""position: absolute; top: 277px; left: 105px""><img src=""images\yellow180.gif"" alt=""Alternate Required: " & alts(x, 1) & ", Fuel required: " & alts(x, 3) & """></div>")
            subWrite(1, 6, 281, 105, 22, "#000", 180, 1, 0, alts(x, 2) & " • " & altsfuel)
            subWrite(2, 6, 201, 0, 20, "#060", 0, 1, 0, "Alternate")
            subWrite(2, 6, 229, 0, 20, "#060", 0, 1, 0, "Alt/Fuel")
            subWrite(2, 6, 201, 105, 22, "#000", 180, 1, 0, varAlternate)
            PrintLine(2, SPC(6), "<div style=""position: absolute; top: 227px; left: 105px""><img src=""images\yellow180.gif"" alt=""Alternate Required: " & alts(x, 1) & ", Fuel required: " & alts(x, 3) & """></div>")
            subWrite(2, 6, 229, 105, 22, "#000", 180, 1, 0, alts(x, 2) & " • " & altsfuel)
            subWrite(4, 6, 325, 0, 30, "#060", 0, 1, 0, "Alt/Fuel")
            subWrite(4, 6, 427, 0, 28, "#000", 420, 1, 0, varAlternate)
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 360px; left:  30px""><img src=""images\yellow180.gif"" alt=""Alternate Required: " & alts(x, 1) & ", Fuel required: " & alts(x, 3) & """ width=""360px""></div>")
            subWrite(4, 6, 367, 0, 44, "#000", 420, 1, 0, alts(x, 2) & " • " & altsfuel)
        End If
        PrintLine(1, "<!-- Hangover AHC -->")
        PrintLine(2, "<!-- Hangover AHC -->")
        PrintLine(4, "<!-- Hangover AHC -->")
        If varAHC = "Yes" Then
            PrintLine(1, SPC(6), "<div style=""position: absolute; top: 313px; left: 0px; font-size: 20px; color: #763900; font-weight: bold""><u>AHC</u></div>")
            PrintLine(2, SPC(6), "<div style=""position: absolute; top: 260px; left: 0px; font-size: 20px; color: #763900; font-weight: bold""><u>AHC</u></div>")
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 450px; left: 0px; font-size: 30px; color: #763900; font-weight: bold""><u>AHC</u></div>")
            If varA1B <> "" Then
                PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 313px; left: 50px; width: 233px; height: 23px;""></div>")
                PrintLine(1, SPC(6), "<div style=""position: absolute; top: 316px; left: 53px; font-size: 16px; color: #060;"">From</div>")
                PrintLine(1, SPC(6), "<div style=""position: absolute; top: 314px; left: 135px; font-size: 20px; color: #000;"">" & varA1B & " to " & varA1E & "</div>")
                PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 287px; left: 7px; width: 276px; height: 23px;""></div>")
                PrintLine(2, SPC(6), "<div style=""position: absolute; top: 290px; left: 10px; font-size: 16px; color: #060;"">From</div>")
                PrintLine(2, SPC(6), "<div style=""position: absolute; top: 288px; left: 100px; font-size: 20px; color: #000;"">" & varA1B & " to " & varA1E & "</div>")
                PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 488px; left: 5px; width: 415px; height: 33px;""></div>")
                PrintLine(4, SPC(6), "<div style=""position: absolute; top: 491px; left: 10px; font-size: 26px; color: #060;"">From</div>")
                PrintLine(4, SPC(6), "<div style=""position: absolute; top: 489px; left: 150px; font-size: 30px; color: #000;"">" & varA1B & " to " & varA1E & "</div>")
            End If
            If varA2B <> "" Then
                PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 337px; left: 50px; width: 233px; height: 23px;""></div>")
                PrintLine(1, SPC(6), "<div style=""position: absolute; top: 340px; left: 53px; font-size: 16px; color: #060;"">From</div>")
                PrintLine(1, SPC(6), "<div style=""position: absolute; top: 338px; left: 135px; font-size: 20px; color: #000;"">" & varA2B & " to " & varA2E & "</div>")
                PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 311px; left: 7px; width: 276px; height: 23px;""></div>")
                PrintLine(2, SPC(6), "<div style=""position: absolute; top: 314px; left: 10px; font-size: 16px; color: #060;"">From</div>")
                PrintLine(2, SPC(6), "<div style=""position: absolute; top: 312px; left: 100px; font-size: 20px; color: #000;"">" & varA2B & " to " & varA2E & "</div>")
                PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 522px; left: 5px; width: 415px; height: 33px;""></div>")
                PrintLine(4, SPC(6), "<div style=""position: absolute; top: 525px; left: 10px; font-size: 26px; color: #060;"">From</div>")
                PrintLine(4, SPC(6), "<div style=""position: absolute; top: 523px; left: 150px; font-size: 30px; color: #000;"">" & varA2B & " to " & varA2E & "</div>")
            End If
            If varA3B <> "" Then
                PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 361px; left: 50px; width: 233px; height: 23px;""></div>")
                PrintLine(1, SPC(6), "<div style=""position: absolute; top: 364px; left: 53px; font-size: 16px; color: #060;"">From</div>")
                PrintLine(1, SPC(6), "<div style=""position: absolute; top: 362px; left: 135px; font-size: 20px; color: #000;"">" & varA3B & " to " & varA3E & "</div>")
                PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 335px; left: 7px; width: 276px; height: 23px;""></div>")
                PrintLine(2, SPC(6), "<div style=""position: absolute; top: 338px; left: 10px; font-size: 16px; color: #060;"">From</div>")
                PrintLine(2, SPC(6), "<div style=""position: absolute; top: 336px; left: 100px; font-size: 20px; color: #000;"">" & varA3B & " to " & varA3E & "</div>")
                PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 556px; left: 5px; width: 415px; height: 33px;""></div>")
                PrintLine(4, SPC(6), "<div style=""position: absolute; top: 559px; left: 10px; font-size: 26px; color: #060;"">From</div>")
                PrintLine(4, SPC(6), "<div style=""position: absolute; top: 557px; left: 150px; font-size: 30px; color: #000;"">" & varA3B & " to " & varA3E & "</div>")
            End If
        End If
        PrintLine(1, "<!-- Hangover End -->")
        PrintLine(2, "<!-- Hangover End -->")
        PrintLine(4, "<!-- Hangover End -->")
        PrintLine(1, "")
        PrintLine(1, "<!-- Birds Start -->")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 181px; left: 306px""><img src=""images\birds.jpg""></div>")
        PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 180px; left: 305px; width: 105px; height:  70px;""></div>")
        PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 181px; left: 306px; width: 103px; height:  68px; border-color: #c5885f""></div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 185px; left: 310px; font-size: 20px; color: #763900; width: 95px; text-align: center; font-weight: bold""><u>BIRDS</u></div>")
        PrintLine(2, "")
        PrintLine(2, "<!-- Birds Start -->")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 131px; left: 306px""><img src=""images\birds.jpg"" width=""127px"" height=""77px""></div>")
        PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 130px; left: 305px; width: 120px; height:  70px;""></div>")
        PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 131px; left: 306px; width: 118px; height:  68px; border-color: #c5885f""></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 135px; left: 310px; font-size: 22px; color: #763900; width: 110px; text-align: center; font-weight: bold""><u>BIRDS</u></div>")
        PrintLine(4, "")
        PrintLine(4, "<!-- Birds Start -->")
        PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 160px; left: 655px; width: 205px; height: 120px;""></div>")
        PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 161px; left: 656px; width: 203px; height: 118px; border-color: #c5885f""></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 165px; left: 655px; font-size: 30px; color: #763900; width: 205px; text-align: center; font-weight: bold""><u>BIRDS</u></div>")
        If varBirdColor = "Green" Then
            l = grnb
        ElseIf varBirdColor = "Yellow" Then
            l = yelb
        ElseIf varBirdColor = "Red" Then
            l = redb
        End If
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 213px; left: 309px""><img src=""" & l & """ alt=""Bird Conditions is " & varBirds & """ width=""100px"" height=""36px""></div>")
        PrintLine(1, SPC(6), "<div id=""bird"" style=""position: absolute; top: 217px; left: 309px; font-size: 20px; color: #000; width: 95px; text-align: center; font-weight: bold"">" & varBird2 & "</div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 163px; left: 309px""><img src=""" & l & """ alt=""Bird Conditions is " & varBirds & """ width=""115px"" height=""36px""></div>")
        PrintLine(2, SPC(6), "<div id=""bird"" style=""position: absolute; top: 167px; left: 309px; font-size: 22px; color: #000; width: 110px; text-align: center; font-weight: bold"">" & varBird2 & "</div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 203px; left: 665px""><img src=""" & l & """ alt=""Bird Conditions is " & varBirds & """ width=""195px"" height=""70px""></div>")
        PrintLine(4, SPC(6), "<div id=""bird"" style=""position: absolute; top: 209px; left: 665px; font-size: 40px; color: #000; width: 185px; text-align: center; font-weight: bold"">" & varBird2 & "</div>")
        PrintLine(1, "<!-- Birds End -->")
        PrintLine(2, "<!-- Birds End -->")
        PrintLine(4, "<!-- Birds End -->")
        PrintLine(1, "")
        PrintLine(1, "<!-- ITS Start -->")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 181px; left: 421px""><img src=""images\birds.jpg""></div>")
        PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 180px; left: 420px; width: 105px; height:  70px;""></div>")
        PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 181px; left: 421px; width: 103px; height:  68px; border-color: #c5885f""></div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 185px; left: 425px; font-size: 20px; color: #763900; width: 95px; text-align: center; font-weight:bold""><u>ITS</u></div>")
        PrintLine(2, "")
        PrintLine(2, "<!-- ITS Start -->")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 211px; left: 306px""><img src=""images\birds.jpg"" width=""127px"" height=""77px""></div>")
        PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 210px; left: 305px; width: 120px; height:  70px;""></div>")
        PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 211px; left: 306px; width: 118px; height:  68px; border-color: #c5885f""></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 215px; left: 310px; font-size: 22px; color: #763900; width: 110px; text-align: center; font-weight: bold""><u>ITS</u></div>")
        PrintLine(4, "")
        PrintLine(4, "<!-- ITS Start -->")
        PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 160px; left: 870px; width: 205px; height: 120px;""></div>")
        PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 161px; left: 871px; width: 203px; height: 118px; border-color: #c5885f""></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 165px; left: 870px; font-size: 30px; color: #763900; width: 205px; text-align: center; font-weight:bold""><u>ITS</u></div>")
        If varITS = "Normal" Then
            l = grnb
        ElseIf varITS = "Caution" Then
            l = yelb
        ElseIf varITS = "Danger" Then
            l = redb
        End If
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 213px; left: 424px""><img src=""" & l & """ alt=""ITS Condition is " & varITS & """ width=""100px"" height=""36px""></div>")
        subWrite(1, 6, 217, 425, 20, "#000", 95, 1, 0, varITS)
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 243px; left: 309px""><img src=""" & l & """ alt=""ITS Condition is " & varITS & """ width=""115px"" height=""36px""></div>")
        subWrite(2, 6, 247, 310, 22, "#000", 110, 1, 0, varITS)
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 203px; left: 880px""><img src=""" & l & """ alt=""ITS Condition is " & varITS & """ width=""195px"" height=""70px""></div>")
        subWrite(4, 6, 209, 880, 40, "#000", 185, 1, 0, varITS)
        PrintLine(1, "<!-- ITS End -->")
        PrintLine(2, "<!-- ITS End -->")
        PrintLine(4, "<!-- ITS End -->")
        PrintLine(1, "")
        PrintLine(1, "<!-- MOAs Start-->")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 261px; left: 306px""><img src=""images\fcif.jpg"" width=""226px"" height=""58px""></div>")
        PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 260px; left: 305px; width: 220px; height: 55px;""></div>")
        PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 261px; left: 306px; width: 218px; height: 53px; border-color: #c5885f""></div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 265px; left: 305px; font-size: 20px; color: #763900; width: 210px; text-align: center; font-weight: bold""><u>MOAs</u></div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 290px; left: 320px; font-size: 16px; color: #000; font-weight: bold"">South</div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 290px; left: 430px; font-size: 16px; color: #000; font-weight: bold"">Tweet</div>")
        PrintLine(2, "")
        PrintLine(2, "<!-- MOAs Start-->")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 131px; left: 436px""><img src=""images\moa.jpg"" width=""127px"" height=""158px""></div>")
        PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 130px; left: 435px; width: 120px; height: 150px;""></div>")
        PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 131px; left: 436px; width: 118px; height: 148px; border-color: #c5885f""></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 135px; left: 440px; font-size: 22px; color: #763900; width: 110px; text-align: center; font-weight: bold""><u>MOAs</u></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 170px; left: 445px; font-size: 22px; color: #000; font-weight: bold"">South</div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 202px; left: 445px; font-size: 22px; color: #000; font-weight: bold"">Tweet</div>")
        PrintLine(4, "")
        PrintLine(4, "<!-- MOAs Start-->")
        PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 160px; left: 440px; width: 205px; height: 165px;""></div>")
        PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 161px; left: 441px; width: 203px; height: 163px; border-color: #c5885f""></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 165px; left: 440px; font-size: 30px; color: #763900; width: 205px; text-align: center; font-weight: bold""><u>MOAs</u></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 214px; left: 450px; font-size: 34px; color: #000; font-weight: bold"">South</div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 274px; left: 450px; font-size: 34px; color: #000; font-weight: bold"">Tweet</div>")
        If varMoaSouth = MOA(0) Then
            l = grn
            j = grnb
        ElseIf varMoaSouth = MOA(1) Then
            l = yel
            j = yelb
        ElseIf varMoaSouth = MOA(2) Then
            l = red
            j = redb
        End If
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 290px; left: 375px""><img src=""" & l & """  alt=""South MOA is " & varMoaSouth & """ height=""20px"" width=""20px""></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 172px; left: 525px""><img src=""" & l & """  alt=""South MOA is " & varMoaSouth & """ height=""22px"" width=""22px""></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 210px; left: 580px""><img src=""" & l & """ alt=""South MOA is " & varMoaSouth & """ height=""45px"" width=""45px""></div>")
        If varMoaTweet = MOA(0) Then
            l = grn
            j = grnb
        ElseIf varMoaTweet = MOA(1) Then
            l = yel
            j = yelb
        ElseIf varMoaTweet = MOA(2) Then
            l = red
            j = redb
        End If
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 290px; left: 485px""><img src=""" & l & """  alt=""Tweet MOA is " & varMoaTweet & """ height=""20px"" width=""20px""></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 204px; left: 525px""><img src=""" & l & """  alt=""Tweet MOA is " & varMoaTweet & """ height=""22px"" width=""22px""></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 270px; left: 580px""><img src=""" & l & """ alt=""Tweet MOA is " & varMoaTweet & """ height=""45px"" width=""45px""></div>")
        PrintLine(1, "<!-- MOAs End -->")
        PrintLine(2, "<!-- MOAs End -->")
        PrintLine(4, "<!-- MOAs End -->")
        PrintLine(1, "")
        PrintLine(1, "<!-- SR Routes Start -->")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 181px; left: 536px""><img src=""images\nav.jpg"" width=""112px"" height=""217px""></div>")
        PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 180px; left: 535px; width: 105px; height: 210px;""></div>")
        PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 181px; left: 536px; width: 103px; height: 208px; border-color: #c5885f""></div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 185px; left: 535px; font-size: 20px; color: #763900; width: 105px; text-align: center; font-weight: bold""><u>SR Routes</u></div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 216px; left: 545px; font-size: 18px; color: #000; font-weight: bold"">" & sr(1) & "</div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 244px; left: 545px; font-size: 18px; color: #000; font-weight: bold"">" & sr(2) & "</div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 272px; left: 545px; font-size: 18px; color: #000; font-weight: bold"">" & sr(3) & "</div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 300px; left: 545px; font-size: 18px; color: #000; font-weight: bold"">" & sr(4) & "</div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 328px; left: 545px; font-size: 18px; color: #000; font-weight: bold"">" & sr(5) & "</div>")
        PrintLine(2, "")
        PrintLine(2, "<!-- SR Routes Start -->")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 131px; left: 566px""><img src=""images\nav.jpg"" width=""127px"" height=""242px""></div>")
        PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 130px; left: 565px; width: 120px; height: 235px;""></div>")
        PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 131px; left: 566px; width: 118px; height: 233px; border-color: #c5885f""></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 135px; left: 563px; font-size: 22px; color: #763900; width: 126px; text-align: center; font-weight: bold""><u>SR Routes</u></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 170px; left: 575px; font-size: 22px; color: #000; font-weight: bold"">" & sr(1) & "</div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 202px; left: 575px; font-size: 22px; color: #000; font-weight: bold"">" & sr(2) & "</div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 234px; left: 575px; font-size: 22px; color: #000; font-weight: bold"">" & sr(3) & "</div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 266px; left: 575px; font-size: 22px; color: #000; font-weight: bold"">" & sr(4) & "</div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 298px; left: 575px; font-size: 22px; color: #000; font-weight: bold"">" & sr(5) & "</div>")
        PrintLine(4, "")
        PrintLine(4, "<!-- SR Routes Start -->")
        PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 335px; left: 440px; width: 205px; height: 300px;""></div>")
        PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 336px; left: 441px; width: 203px; height: 298px; border-color: #c5885f""></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 340px; left: 440px; font-size: 30px; color: #763900; width: 205px; text-align: center; font-weight: bold""><u>SR Routes</u></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 387px; left: 450px; font-size: 34px; color: #000; font-weight: bold"">" & sr(1) & "</div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 437px; left: 450px; font-size: 34px; color: #000; font-weight: bold"">" & sr(2) & "</div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 487px; left: 450px; font-size: 34px; color: #000; font-weight: bold"">" & sr(3) & "</div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 537px; left: 450px; font-size: 34px; color: #000; font-weight: bold"">" & sr(4) & "</div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 587px; left: 450px; font-size: 34px; color: #000; font-weight: bold"">" & sr(5) & "</div>")
        If rbSr1Open.Checked Then
            l = grn
            j = grnb
        ElseIf rbSr1Closed.Checked Then
            l = red
            j = redb
        End If
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 216px; left: 610px""><img src=""" & l & """ alt=""" & sr(1) & " is " & varSR1 & """ height=""20px"" width=""20px""></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 170px; left: 655px""><img src=""" & l & """ alt=""" & sr(1) & " is " & varSR1 & """ height=""22px"" width=""22px""></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 383px; left: 580px""><img src=""" & l & """ alt=""" & sr(1) & " is " & varSR1 & """ height=""45px"" width=""45px""></div>")
        If rbSr2Open.Checked Then
            l = grn
            j = grnb
        ElseIf rbSr2Closed.Checked Then
            l = red
            j = redb
        End If
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 244px; left: 610px""><img src=""" & l & """ alt=""" & sr(2) & " is " & varSR2 & """ height=""20px"" width=""20px""></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 204px; left: 655px""><img src=""" & l & """ alt=""" & sr(2) & " is " & varSR2 & """ height=""22px"" width=""22px""></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 433px; left: 580px""><img src=""" & l & """ alt=""" & sr(2) & " is " & varSR2 & """ height=""45px"" width=""45px""></div>")
        If rbSr3Open.Checked Then
            l = grn
            j = grnb
        ElseIf rbSr3Closed.Checked Then
            l = red
            j = redb
        End If
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 272px; left: 610px""><img src=""" & l & """ alt=""" & sr(3) & " is " & varSR3 & """ height=""20px"" width=""20px""></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 234px; left: 655px""><img src=""" & l & """ alt=""" & sr(3) & " is " & varSR3 & """ height=""22px"" width=""22px""></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 483px; left: 580px""><img src=""" & l & """ alt=""" & sr(3) & " is " & varSR3 & """ height=""45px"" width=""45px""></div>")
        If rbSr4Open.Checked Then
            l = grn
            j = grnb
        ElseIf rbSr4Closed.Checked Then
            l = red
            j = redb
        End If
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 300px; left: 610px""><img src=""" & l & """ alt=""" & sr(4) & " is " & varSR4 & """ height=""20px"" width=""20px""></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 266px; left: 655px""><img src=""" & l & """ alt=""" & sr(4) & " is " & varSR4 & """ height=""22px"" width=""22px""></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 533px; left: 580px""><img src=""" & l & """ alt=""" & sr(4) & " is " & varSR4 & """ height=""45px"" width=""45px""></div>")
        If rbSr5Open.Checked Then
            l = grn
            j = grnb
        ElseIf rbSr5Closed.Checked Then
            l = red
            j = redb
        End If
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 328px; left: 610px""><img src=""" & l & """ alt=""" & sr(5) & " is " & varSR5 & """ height=""20px"" width=""20px""></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 298px; left: 655px""><img src=""" & l & """ alt=""" & sr(5) & " is " & varSR5 & """ height=""22px"" width=""22px""></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 583px; left: 580px""><img src=""" & l & """ alt=""" & sr(5) & " is " & varSR5 & """ height=""45px"" width=""45px""></div>")
        PrintLine(1, "<!-- SR Routes End -->")
        PrintLine(2, "<!-- SR Routes End -->")
        PrintLine(4, "<!-- SR Routes End -->")
        PrintLine(1, "")
        PrintLine(1, "<!-- NAVAIDs Start -->")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 181px; left: 651px""><img src=""images\nav.jpg"" width=""112px"" height=""217px""></div>")
        PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 180px; left: 650px; width: 105px; height: 210px;""></div>")
        PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 181px; left: 651px; width: 103px; height: 208px; border-color: #c5885f""></div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 185px; left: 655px; font-size: 20px; color: #763900; width: 95px; text-align: center; font-weight: bold""><u>NAVAIDS</u></div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 216px; left: 660px; font-size: 18px; color: #000; font-weight: bold"">VOR</div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 244px; left: 660px; font-size: 18px; color: #000; font-weight: bold"">ILS</div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 272px; left: 660px; font-size: 18px; color: #000; font-weight: bold"">DME</div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 300px; left: 660px; font-size: 18px; color: #000; font-weight: bold"">SAT</div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 328px; left: 660px; font-size: 18px; color: #000; font-weight: bold"">SKF</div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 356px; left: 660px; font-size: 18px; color: #000; font-weight: bold"">SSF</div>")
        PrintLine(2, "")
        PrintLine(2, "<!-- NAVAIDs Start -->")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 131px; left: 696px""><img src=""images\nav.jpg"" width=""122px"" height=""242px""></div>")
        PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 130px; left: 695px; width: 115px; height: 235px;""></div>")
        PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 131px; left: 696px; width: 113px; height: 233px; border-color: #c5885f""></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 135px; left: 700px; font-size: 22px; color: #763900; width: 110px; text-align: center; font-weight: bold""><u>NAVAIDS</u></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 170px; left: 705px; font-size: 22px; color: #000; font-weight: bold"">VOR</div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 202px; left: 705px; font-size: 22px; color: #000; font-weight: bold"">ILS</div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 234px; left: 705px; font-size: 22px; color: #000; font-weight: bold"">DME</div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 266px; left: 705px; font-size: 22px; color: #000; font-weight: bold"">SAT</div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 298px; left: 705px; font-size: 22px; color: #000; font-weight: bold"">SKF</div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 330px; left: 705px; font-size: 22px; color: #000; font-weight: bold"">SSF</div>")
        PrintLine(4, "")
        PrintLine(4, "<!-- NAVAIDs Start -->")
        PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 645px; left: 440px; width: 205px; height: 350px;""></div>")
        PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 646px; left: 441px; width: 203px; height: 348px; border-color: #c5885f""></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 650px; left: 440px; font-size: 30px; color: #763900; width: 205px; text-align: center; font-weight: bold""><u>NAVAIDS</u></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 697px; left: 450px; font-size: 34px; color: #000; font-weight: bold"">VOR</div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 747px; left: 450px; font-size: 34px; color: #000; font-weight: bold"">ILS</div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 792px; left: 450px; font-size: 34px; color: #000; font-weight: bold"">DME</div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 847px; left: 450px; font-size: 34px; color: #000; font-weight: bold"">SAT</div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 892px; left: 450px; font-size: 34px; color: #000; font-weight: bold"">SKF</div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 947px; left: 450px; font-size: 34px; color: #000; font-weight: bold"">SSF</div>")
        If rbNavRndUp.Checked Then
            l = grn
            j = grnb
        ElseIf rbNavRndDown.Checked Then
            l = red
            j = redb
        End If
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 216px; left: 720px""><img src=""" & l & """ alt=""RND VOR is " & varNavRND & """ height=""20px"" width=""20px""></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 172px; left: 775px""><img src=""" & l & """ alt=""RND VOR is " & varNavRND & """ height=""22px"" width=""22px""></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 693px; left: 580px""><img src=""" & l & """ alt=""RND VOR is " & varNavRND & """ height=""45px"" width=""45px""></div>")
        If rbNavIlsUp.Checked Then
            l = grn
            j = grnb
        ElseIf rbNavIlsDown.Checked Then
            l = red
            j = redb
        End If
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 244px; left: 720px""><img src=""" & l & """ alt=""RND ILS is " & varNavILS & """ height=""20px"" width=""20px""></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 204px; left: 775px""><img src=""" & l & """ alt=""RND ILS is " & varNavILS & """ height=""22px"" width=""22px""></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 743px; left: 580px""><img src=""" & l & """ alt=""RND ILS is " & varNavILS & """ height=""45px"" width=""45px""></div>")
        If rbNavDmeUp.Checked Then
            l = grn
            j = grnb
        ElseIf rbNavDmeDown.Checked Then
            l = red
            j = redb
        End If
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 272px; left: 720px""><img src=""" & l & """ alt=""RND DME is " & varNavDME & """ height=""20px"" width=""20px""></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 236px; left: 775px""><img src=""" & l & """ alt=""RND DME is " & varNavDME & """ height=""22px"" width=""22px""></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 793px; left: 580px""><img src=""" & l & """ alt=""RND DME is " & varNavDME & """ height=""45px"" width=""45px""></div>")
        If rbNavSatUp.Checked Then
            l = grn
            j = grnb
        ElseIf rbNavSatDown.Checked Then
            l = red
            j = redb
        End If
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 300px; left: 720px""><img src=""" & l & """ alt=""SAT VOR is " & varNavSAT & """ height=""20px"" width=""20px""></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 268px; left: 775px""><img src=""" & l & """ alt=""SAT VOR is " & varNavSAT & """ height=""22px"" width=""22px""></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 843px; left: 580px""><img src=""" & l & """ alt=""SAT VOR is " & varNavSAT & """ height=""45px"" width=""45px""></div>")
        If rbNavSkfUp.Checked Then
            l = grn
            j = grnb
        ElseIf rbNavSkfDown.Checked Then
            l = red
            j = redb
        End If
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 328px; left: 720px""><img src=""" & l & """ alt=""SKF VORTAC is " & varNavSKF & """ height=""20px"" width=""20px""></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 300px; left: 775px""><img src=""" & l & """ alt=""SKF VORTAC is " & varNavSKF & """ height=""22px"" width=""22px""></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 893px; left: 580px""><img src=""" & l & """ alt=""SKF VORTAC is " & varNavSKF & """ height=""45px"" width=""45px""></div>")
        If rbNavSsfUp.Checked Then
            l = grn
            j = grnb
        ElseIf rbNavSsfDown.Checked Then
            l = red
            j = redb
        End If
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 356px; left: 720px""><img src=""" & l & """ alt=""SSF VOR is " & varNavSSF & """ height=""20px"" width=""20px""></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 332px; left: 775px""><img src=""" & l & """ alt=""SSF VOR is " & varNavSSF & """ height=""22px"" width=""22px""></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 943px; left: 580px""><img src=""" & l & """ alt=""SSF VOR is " & varNavSSF & """ height=""45px"" width=""45px""></div>")
        PrintLine(1, "<!-- NAVAIDs End -->")
        PrintLine(2, "<!-- NAVAIDs End -->")
        PrintLine(4, "<!-- NAVAIDs End -->")
        PrintLine(1, "")
        PrintLine(1, "<!-- FCIF and Boldface Start -->")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 401px; left: 536px""><img src=""images\fcif.jpg"" height=""177px"" width=""227px""></div>")
        PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 400px; left: 535px; width: 220px; height: 170px;""></div>")
        PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 401px; left: 536px; width: 218px; height: 168px; border-color: #c5885f""></div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 405px; left: 540px; font-size: 22px; color: #763900; text-align: left; font-weight: bold""><u>FCIF</u></div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 470px; left: 540px; font-size: 20px; color: #763900; width: 66px; text-align: center; font-weight: bold"">Bravo</div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 470px; left: 612px; font-size: 20px; color: #763900; width: 66px; text-align: center; font-weight: bold"">Charlie</div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 470px; left: 684px; font-size: 20px; color: #763900; width: 66px; text-align: center; font-weight: bold"">PIF</div>")
        PrintLine(2, "")
        PrintLine(2, "<!-- FCIF and Boldface Start -->")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 131px; left:  821px""><img src=""images\fcif.jpg""></div>")
        PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 130px; left: 820px; width: 275px; height: 145px;""></div>")
        PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 131px; left: 821px; width: 273px; height: 143px; border-color: #c5885f""></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 135px; left:  820px; font-size: 22px; color: #763900; width: 275px; text-align: center; font-weight: bold""><u>FCIF</u></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 175px; left:  825px; font-size: 22px; color: #763900; width: 85px; text-align: center; font-weight: bold"">Bravo</div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 175px; left:  915px; font-size: 20px; color: #763900; width: 85px; text-align: center; font-weight: bold"">Charlie</div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 175px; left: 1005px; font-size: 20px; color: #763900; width: 85px; text-align: center; font-weight: bold"">PIF</div>")
        PrintLine(4, "")
        PrintLine(4, "<!-- FCIF and Boldface Start -->")
        PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 735px; left:  -5px; width: 435px; height: 260px;""></div>")
        PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 736px; left:  -4px; width: 433px; height: 258px; border-color: #c5885f""></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 740px; left:  -5px; font-size: 32px; color: #763900; width: 435px; text-align: center; font-weight: bold""><u>FCIF</u></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 830px; left:   5px; font-size: 30px; color: #763900; width: 140px; text-align: center; font-weight: bold"">Bravo</div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 830px; left: 145px; font-size: 30px; color: #763900; width: 140px; text-align: center; font-weight: bold"">Charlie</div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 830px; left: 285px; font-size: 30px; color: #763900; width: 140px; text-align: center; font-weight: bold"">PIF</div>")
        If cbBNew.Checked = True Then
            PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 498px; left: 540px; width:  66px; height:  65px; background-color: #d7d7d7""></div>")
            PrintLine(1, SPC(6), "<div style=""top: 535px; left: 543px""><img src=""images/new2.gif"" width=""60"" height=""25"" alt=""NEW FCIF""></div>")
            subWrite(1, 6, 500, 540, 22, "#000", 66, 1, 0, varFCIFB)
            PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 205px; left: 825px; width:  85px; height:  65px; background-color: #d7d7d7""></div>")
            PrintLine(2, SPC(6), "<div style=""top: 235px; left: 838px""><img src=""images/new2.gif"" width=""60"" height=""25"" alt=""NEW FCIF""></div>")
            subWrite(2, 6, 205, 825, 22, "#000", 85, 1, 0, varFCIFB)
            PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 865px; left:   5px; width: 135px; height: 120px; background-color: #d7d7d7""></div>")
            PrintLine(4, SPC(6), "<div style=""top: 935px; left:  22px""><img src=""images/new2.gif"" width=""100"" alt=""NEW FCIF""></div>")
            subWrite(4, 6, 865, 5, 48, "#000", 135, 1, 0, varFCIFB)
        Else
            subWrite(1, 6, 500, 540, 22, "#000", 66, 1, "#d7d7d7", varFCIFB)
            subWrite(2, 6, 205, 825, 22, "#000", 85, 1, "#d7d7d7", varFCIFB)
            subWrite(4, 6, 865, 5, 48, "#000", 135, 1, "#d7d7d7", varFCIFB)
        End If
        If cbCNew.Checked = True Then
            PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 498px; left: 612px; width:  66px; height:  65px; background-color: #d7d7d7""></div>")
            PrintLine(1, SPC(6), "<div style=""top: 535px; left: 615px""><img src=""images/new2.gif"" width=""60"" height=""25"" alt=""NEW FCIF""></div>")
            subWrite(1, 6, 500, 612, 22, "#000", 66, 1, 0, varFCIFC)
            PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 205px; left: 915px; width:  85px; height:  65px; background-color: #d7d7d7""></div>")
            PrintLine(2, SPC(6), "<div style=""top: 235px; left: 928px""><img src=""images/new2.gif"" width=""60"" height=""25"" alt=""NEW FCIF""></div>")
            subWrite(2, 6, 205, 915, 22, "#000", 85, 1, 0, varFCIFC)
            PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 865px; left: 145px; width:  135px; height: 120px; background-color: #d7d7d7""></div>")
            PrintLine(4, SPC(6), "<div style=""top: 935px; left: 162px""><img src=""images/new2.gif"" width=""100"" alt=""NEW FCIF""></div>")
            subWrite(4, 6, 865, 145, 48, "#000", 135, 1, 0, varFCIFC)
        Else
            subWrite(1, 6, 500, 612, 22, "#000", 66, 1, "#d7d7d7", varFCIFC)
            subWrite(2, 6, 205, 915, 22, "#000", 85, 1, "#d7d7d7", varFCIFC)
            subWrite(4, 6, 865, 145, 48, "#000", 135, 1, "#d7d7d7", varFCIFC)
        End If
        If cbPNew.Checked = True Then
            PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 498px; left: 684px; width:  66px; height:  65px; background-color: #d7d7d7""></div>")
            PrintLine(1, SPC(6), "<div style=""top: 535px; left: 687px""><img src=""images/new2.gif"" width=""60"" height=""25"" alt=""NEW PIF""></div>")
            subWrite(1, 6, 500, 684, 22, "#000", 66, 1, 0, varPIF)
            PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 205px; left: 1005px; width:  85px; height:  65px; background-color: #d7d7d7""></div>")
            PrintLine(2, SPC(6), "<div style=""top: 235px; left: 1018px""><img src=""images/new2.gif"" width=""60"" height=""25"" alt=""NEW PIF""></div>")
            subWrite(2, 6, 205, 1005, 22, "#000", 85, 1, 0, varPIF)
            PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 865px; left: 285px; width: 135px; height: 120px; background-color: #d7d7d7""></div>")
            PrintLine(4, SPC(6), "<div style=""top: 935px; left: 302px""><img src=""images/new2.gif"" width=""100"" alt=""NEW PIF""></div>")
            subWrite(4, 6, 865, 285, 48, "#000", 135, 1, 0, varPIF)
        Else
            subWrite(1, 6, 500, 684, 22, "#000", 66, 1, "#d7d7d7", varPIF)
            subWrite(2, 6, 205, 1005, 22, "#000", 85, 1, "#d7d7d7", varPIF)
            subWrite(4, 6, 865, 285, 48, "#000", 135, 1, "#d7d7d7", varPIF)
        End If
        If rbBfYes.Checked Then
            PrintLine(1, SPC(6), "<div style=""position: absolute; top: 435px; left: 649px""><img src=""images\red100.gif"" alt=""Boldface Due"" width=""106px""></div>")
            PrintLine(1, SPC(6), "<div id=""bf"" style=""position: absolute; top: 442px; left: 649px; font-size: 14px; color: black; width: 100px; text-align: center; font-weight: bold"">Boldface Due</div>")
            PrintLine(2, SPC(6), "<div style=""position: absolute; top: 135px; left: 989px""><img src=""images\red100.gif"" alt=""Boldface Due"" width=""106px""></div>")
            PrintLine(2, SPC(6), "<div id=""bf"" style=""position: absolute; top: 142px; left: 989px; font-size: 14px; color: black; width: 100px; text-align: center; font-weight: bold"">Boldface Due</div>")
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 750px; left: 265px""><img src=""images\red100.gif"" alt=""Boldface Due"" width=""160px"" height=""80px""></div>")
            PrintLine(4, SPC(6), "<div id=""bf"" style=""position: absolute; top: 750px; left: 265px; font-size: 28px; color: black; width: 150px; text-align: center; font-weight: bold"">Boldface Due</div>")
        ElseIf rbBfNo.Checked Then
            PrintLine(1, SPC(6), "<div id=""bf""></div>")
            PrintLine(2, SPC(6), "<div id=""bf""></div>")
            PrintLine(4, SPC(6), "<div id=""bf""></div>")
        End If
        If rbSRFYes.Checked Then
            PrintLine(1, SPC(6), "<div style=""position: absolute; top: 435px; left: 540px""><img src=""images\red100.gif"" alt=""Safety Read File"" width=""106px""></div>")
            PrintLine(1, SPC(6), "<div id=""srf"" style=""position: absolute; top: 442px; left: 540px; font-size: 14px; color: black; width: 100px; text-align: center; font-weight: bold"">Safety Rd File</div>")
            PrintLine(2, SPC(6), "<div style=""position: absolute; top: 135px; left: 825px""><img src=""images\red100.gif"" alt=""Safety Read Filee"" width=""106px""></div>")
            PrintLine(2, SPC(6), "<div id=""srf"" style=""position: absolute; top: 142px; left: 825px; font-size: 14px; color: black; width: 100px; text-align: center; font-weight: bold"">Safety Rd File</div>")
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 750px; left:   5px""><img src=""images\red100.gif"" alt=""Safety Read Filee"" width=""160px"" height=""80px""></div>")
            PrintLine(4, SPC(6), "<div id=""srf"" style=""position: absolute; top: 750px; left:   5px; font-size: 28px; color: black; width: 150px; text-align: center; font-weight: bold"">Safety Rd File</div>")
        ElseIf rbSRFNo.Checked Then
            PrintLine(1, SPC(6), "<div id=""srf""></div>")
            PrintLine(2, SPC(6), "<div id=""srf""></div>")
            PrintLine(4, SPC(6), "<div id=""srf""></div>")
        End If
        PrintLine(1, "<!-- FCIF and Boldface End -->")
        PrintLine(2, "<!-- FCIF and Boldface End -->")
        PrintLine(4, "<!-- FCIF and Boldface End -->")
        PrintLine(1, "")
        PrintLine(1, "<!-- TOLD Start -->")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 581px; left: 536px""><img src=""images\told.jpg"" width=""227px"" height=""207px""></div>")
        PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 580px; left: 535px; width: 220px; height: 200px;""></div>")
        PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 581px; left: 536px; width: 218px; height: 198px; border-color: #c5885f""></div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 590px; left: 540px; font-size: 24px; color: #763900; width: 100px; text-align: left; font-weight: bold""><u>TOLD</u></div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 670px; left: 540px; font-size: 14px; color: #763900; font-weight: bold"">Temp</div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 720px; left: 540px; font-size: 14px; color: #763900; font-weight: bold"">PA</div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 588px; left: 625px; font-size: 14px; color: #763900; width: 127px; text-align: center; font-weight: bold"">Takeoff Roll</div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 653px; left: 625px; font-size: 14px; color: #763900; width: 127px; text-align: center; font-weight: bold"">Abort Speed</div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 718px; left: 625px; font-size: 14px; color: #763900; width: 127px; text-align: center; font-weight: bold"">No Flap Landing</div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 695px; left: 645px; font-size: 14px; color: #763900; font-weight: bold"">Dry</div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 695px; left: 705px; font-size: 14px; color: #763900; font-weight: bold"">Wet</div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 760px; left: 645px; font-size: 14px; color: #763900; font-weight: bold"">Dry</div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 760px; left: 705px; font-size: 14px; color: #763900; font-weight: bold"">Wet</div>")
        PrintLine(2, "")
        PrintLine(2, "<!-- TOLD Start -->")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 286px; left: 821px""><img src=""images\told.jpg"" width=""282px"" height=""207px""></div>")
        PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 285px; left: 820px; width: 275px; height: 200px;""></div>")
        PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 286px; left: 821px; width: 273px; height: 198px; border-color: #c5885f""></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 295px; left: 810px; font-size: 24px; color: #763900; width: 160px; text-align: center; font-weight: bold""><u>TOLD</u></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 415px; left: 825px; font-size: 14px; color: #763900; font-weight: bold"">Temp</div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 440px; left: 835px; font-size: 14px; color: #763900; font-weight: bold"">PA</div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 293px; left: 965px; font-size: 14px; color: #763900; width: 127px; text-align: center; font-weight: bold"">Takeoff Roll</div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 358px; left: 965px; font-size: 14px; color: #763900; width: 127px; text-align: center; font-weight: bold"">Abort Speed</div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 423px; left: 965px; font-size: 14px; color: #763900; width: 127px; text-align: center; font-weight: bold"">No Flap Landing</div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 400px; left: 985px; font-size: 14px; color: #763900; font-weight: bold"">Dry</div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 400px; left: 1045px; font-size: 14px; color: #763900; font-weight: bold"">Wet</div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 465px; left: 985px; font-size: 14px; color: #763900; font-weight: bold"">Dry</div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 465px; left: 1045px; font-size: 14px; color: #763900; font-weight: bold"">Wet</div>")
        PrintLine(4, "")
        PrintLine(4, "<!-- TOLD Start -->")
        PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 290px; left: 655px; width: 420px; height: 445px;""></div>")
        PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 291px; left: 656px; width: 418px; height: 443px; border-color: #c5885f""></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 295px; left: 655px; font-size: 32px; color: #763900; width: 420px; text-align: center; font-weight: bold""><u>TOLD</u></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 345px; left: 920px; font-size: 24px; color: #763900; font-weight: bold"">Temp</div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 385px; left: 950px; font-size: 24px; color: #763900; font-weight: bold"">PA</div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 340px; left: 670px; font-size: 30px; color: #763900; width: 180px; text-align: center; font-weight: bold"">Takeoff Roll</div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 460px; left: 655px; font-size: 30px; color: #763900; width: 420px; text-align: center; font-weight: bold"">Abort Speed</div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 600px; left: 655px; font-size: 30px; color: #763900; width: 420px; text-align: center; font-weight: bold"">No Flap Landing</div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 565px; left: 745px; font-size: 20px; color: #763900; font-weight: bold"">Dry</div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 565px; left: 955px; font-size: 20px; color: #763900; font-weight: bold"">Wet</div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 705px; left: 745px; font-size: 20px; color: #763900; font-weight: bold"">Dry</div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 705px; left: 955px; font-size: 20px; color: #763900; font-weight: bold"">Wet</div>")
        subWrite(1, 6, 688, 540, 16, "gray", 40, 1, "#d7d7d7", varTemp)
        subWrite(1, 6, 738, 540, 16, "gray", 40, 1, "#d7d7d7", varPA)
        subWrite(1, 6, 620, 540, 12, "gray", 0, 0, 0, varTM & " Method")
        subWrite(1, 6, 605, 657, 20, 0, 60, 1, "#d7d7d7", varToD)
        If varAbD >= 85 Then
            subWrite(1, 6, 670, 625, 20, 0, 60, 1, "#d7d7d7", "ROT")
        Else
            subWrite(1, 6, 670, 625, 20, 0, 60, 1, "#d7d7d7", varAbD)
        End If
        If varAbW >= 85 Then
            subWrite(1, 6, 670, 690, 20, 0, 60, 1, "#d7d7d7", "ROT")
        Else
            subWrite(1, 6, 670, 690, 20, 0, 60, 1, "#d7d7d7", varAbW)
        End If
        subWrite(1, 6, 735, 625, 20, 0, 60, 1, "#d7d7d7", varLdD)
        subWrite(1, 6, 735, 690, 20, 0, 60, 1, "#d7d7d7", varLdW)
        subWrite(2, 6, 413, 867, 16, "gray", 40, 1, "#d7d7d7", varTemp)
        subWrite(2, 6, 438, 867, 16, "gray", 40, 1, "#d7d7d7", varPA)
        subWrite(2, 6, 325, 810, 12, "gray", 160, 0, 0, varTM & " Method")
        subWrite(2, 6, 310, 997, 20, 0, 60, 1, "#d7d7d7", varToD)
        If varAbD >= 85 Then
            subWrite(2, 6, 375, 965, 20, 0, 60, 1, "#d7d7d7", "ROT")
        Else
            subWrite(2, 6, 375, 965, 20, 0, 60, 1, "#d7d7d7", varAbD)
        End If
        If varAbW >= 85 Then
            subWrite(2, 6, 375, 1030, 20, 0, 60, 1, "#d7d7d7", "ROT")
        Else
            subWrite(2, 6, 375, 1030, 20, 0, 60, 1, "#d7d7d7", varAbW)
        End If
        subWrite(2, 6, 440, 965, 20, 0, 60, 1, "#d7d7d7", varLdD)
        subWrite(2, 6, 440, 1030, 20, 0, 60, 1, "#d7d7d7", varLdW)
        subWrite(4, 6, 340, 1000, 26, "gray", 60, 1, "#d7d7d7", varTemp)
        subWrite(4, 6, 380, 1000, 26, "gray", 60, 1, "#d7d7d7", varPA)
        subWrite(4, 6, 305, 920, 16, "gray", 160, 0, 0, varTM & " Method")
        subWrite(4, 6, 380, 670, 55, 0, 180, 1, "#d7d7d7", varToD)
        If varAbD >= 85 Then
            subWrite(4, 6, 500, 670, 55, 0, 180, 1, "#d7d7d7", "ROT")
        Else
            subWrite(4, 6, 500, 670, 55, 0, 180, 1, "#d7d7d7", varAbD)
        End If
        If varAbW >= 85 Then
            subWrite(4, 6, 500, 880, 55, 0, 180, 1, "#d7d7d7", "ROT")
        Else
            subWrite(4, 6, 500, 880, 55, 0, 180, 1, "#d7d7d7", varAbW)
        End If
        subWrite(4, 6, 640, 670, 55, 0, 180, 1, "#d7d7d7", varLdD)
        subWrite(4, 6, 640, 880, 55, 0, 180, 1, "#d7d7d7", varLdW)
        PrintLine(1, "<!-- TOLD End -->")
        PrintLine(2, "<!-- TOLD End -->")
        PrintLine(4, "<!-- TOLD End -->")
        If varQoDShow = "Show" Then
            PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 745px; left: 655px; width: 420px; height: 225px;""></div>")
            PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 746px; left: 656px; width: 418px; height: 223px; border-color: #c5885f""></div>")
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 745px; left: 660px; width: 220px; font-size: 22px; color: #763900; font-weight: bold""><u>Question of the Day</u></div>")
            PrintLine(4, SPC(6), "<div id=""qod4"" style=""top: 775px; left: 665px;"">" & varQoD & "</div>")
        Else
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 747px; left:  657px""><img src=""images\chkrbd.gif"" width=""42px"" height=""42px""></div>")
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 747px; left:  699px""><img src=""images\chkrbd.gif"" width=""42px"" height=""42px""></div>")
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 747px; left:  741px""><img src=""images\chkrbd.gif"" width=""42px"" height=""42px""></div>")
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 747px; left:  783px""><img src=""images\chkrbd.gif"" width=""42px"" height=""42px""></div>")
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 747px; left:  825px""><img src=""images\chkrbd.gif"" width=""42px"" height=""42px""></div>")
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 747px; left:  867px""><img src=""images\chkrbd.gif"" width=""42px"" height=""42px""></div>")
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 747px; left:  909px""><img src=""images\chkrbd.gif"" width=""42px"" height=""42px""></div>")
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 747px; left:  951px""><img src=""images\chkrbd.gif"" width=""42px"" height=""42px""></div>")
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 747px; left:  993px""><img src=""images\chkrbd.gif"" width=""42px"" height=""42px""></div>")
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 747px; left: 1035px""><img src=""images\chkrbd.gif"" width=""42px"" height=""42px""></div>")
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 928px; left:  657px""><img src=""images\chkrbd.gif"" width=""42px"" height=""42px""></div>")
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 928px; left:  699px""><img src=""images\chkrbd.gif"" width=""42px"" height=""42px""></div>")
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 928px; left:  741px""><img src=""images\chkrbd.gif"" width=""42px"" height=""42px""></div>")
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 928px; left:  783px""><img src=""images\chkrbd.gif"" width=""42px"" height=""42px""></div>")
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 928px; left:  825px""><img src=""images\chkrbd.gif"" width=""42px"" height=""42px""></div>")
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 928px; left:  867px""><img src=""images\chkrbd.gif"" width=""42px"" height=""42px""></div>")
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 928px; left:  909px""><img src=""images\chkrbd.gif"" width=""42px"" height=""42px""></div>")
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 928px; left:  951px""><img src=""images\chkrbd.gif"" width=""42px"" height=""42px""></div>")
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 928px; left:  993px""><img src=""images\chkrbd.gif"" width=""42px"" height=""42px""></div>")
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 928px; left: 1035px""><img src=""images\chkrbd.gif"" width=""42px"" height=""42px""></div>")
            PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 745px; left: 655px; width: 420px; height: 225px;""></div>")
            PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 746px; left: 656px; width: 418px; height: 223px; border-color: #c5885f""></div>")
            PrintLine(4, SPC(6), "<div style=""position: absolute; top: 784px; left:  657px""><img src=""images\t-6small.jpg"" width=""418px"" height=""148px""></div>")
        End If
        PrintLine(1, "")
        PrintLine(1, "<!-- Last Sup Update Start -->")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 790px; left: 536px""><img src=""images\up.jpg"" height=""23px"" width=""227px""></div>")
        PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 789px; left: 535px; width: 220px; height: 16px;""></div>")
        PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 790px; left: 536px; width: 218px; height: 14px; border-color: #c5885f""></div>")
        PrintLine(2, "")
        PrintLine(2, "<!-- Last Sup Update Start -->")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 666px; left: 821px""><img src=""images\up.jpg"" width=""282px"" height=""29px""></div>")
        PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 665px; left: 820px; width: 275px; height: 20px;""></div>")
        PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 666px; left: 821px; width: 273px; height: 18px; border-color: #c5885f""></div>")
        PrintLine(4, "")
        PrintLine(4, "<!-- Last Sup Update Start -->")
        PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 980px; left: 655px; width: 420px; height: 15px;""></div>")
        PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 981px; left: 656px; width: 418px; height: 13px; border-color: #c5885f""></div>")
        subWrite(1, 6, 791, 540, 12, "#060", 215, 0, 0, "Last Update: " & Now())
        subWrite(2, 6, 667, 820, 14, "#060", 275, 0, 0, "Last Update: " & Now())
        subWrite(4, 6, 981, 655, 14, "#060", 420, 0, 0, "Last Update: " & Now())
        PrintLine(1, "<!-- Last Sup Update End -->")
        PrintLine(2, "<!-- Last Sup Update End -->")
        PrintLine(4, "<!-- Last Sup Update End -->")
        PrintLine(4, "")
        PrintLine(4, SPC(4), "</div")
        PrintLine(4, SPC(2), "</body>")
        PrintLine(4, SPC(0), "</html>")
        FileClose(4)
        FileOpen(4, lPath & "status4b.html", OpenMode.Append)
        PrintLine(1, "")
        PrintLine(1, "<!-- Sups Comments Start -->")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 496px; left:  -4px""><img src=""images\sups.jpg"" width=""537px"" height=""317px""></div>")
        PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 495px; left:  -5px; width: 530px; height: 310px;""></div>")
        PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 496px; left:  -4px; width: 528px; height: 308px; border-color: #c5885f""></div>")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 495px; left:   0px; font-size: 22px; color: #763900; font-weight: bold""><u>Operations Notes</u></div>")
        PrintLine(2, "")
        PrintLine(2, "<!-- Sups Comments Start -->")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 376px; left:  -4px""><img src=""images\sup2.jpg"" width=""822px"" height=""317px""></div>")
        PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 375px; left:  -5px; width: 815px; height: 310px;""></div>")
        PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 376px; left:  -4px; width: 813px; height: 308px; border-color: #c5885f""></div>")
        PrintLine(2, SPC(6), "<div style=""position: absolute; top: 375px; left:   0px; font-size: 22px; color: #763900; font-weight: bold""><u>Operations Notes</u></div>")
        PrintLine(4, "")
        PrintLine(4, "<!-- Sups Comments Start -->")
        PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 100px; left:  -5px; width: 1080px; height: 895px;""></div>")
        PrintLine(4, SPC(6), "<div id=""boxes"" style=""top: 101px; left:  -4px; width: 1078px; height: 893px; border-color: #c5885f""></div>")
        PrintLine(4, SPC(6), "<div style=""position: absolute; top: 100px; left:   0px; font-size: 30px; width: 1080px; text-align: center; color: #763900; font-weight: bold""><u>Operations Notes</u></div>")
        If varOps1 <> "" Then
            PrintLine(1, SPC(6), "<div id=""comments"" style=""top: 520px; left:   0px;"">" & varOps1 & "</div>")
            PrintLine(2, SPC(6), "<div id=""comments"" style=""top: 404px; left:   0px;"">" & varOps1 & "</div>")
            PrintLine(4, SPC(6), "<div id=""comments"" style=""top: 135px; left:   5px;"">" & varOps1 & "</div>")
        End If
        If varOps2 <> "" Then
            PrintLine(1, SPC(6), "<div id=""comments"" style=""top: 560px; left:   0px;"">" & varOps2 & "</div>")
            PrintLine(2, SPC(6), "<div id=""comments"" style=""top: 404px; left: 405px;"">" & varOps2 & "</div>")
            PrintLine(4, SPC(6), "<div id=""comments"" style=""top: 258px; left:   5px;"">" & varOps2 & "</div>")
        End If
        If varOps3 <> "" Then
            PrintLine(1, SPC(6), "<div id=""comments"" style=""top: 600px; left:   0px;"">" & varOps3 & "</div>")
            PrintLine(2, SPC(6), "<div id=""comments"" style=""top: 473px; left:   0px;"">" & varOps3 & "</div>")
            PrintLine(4, SPC(6), "<div id=""comments"" style=""top: 381px; left:   5px;"">" & varOps3 & "</div>")
        End If
        If varOps4 <> "" Then
            PrintLine(1, SPC(6), "<div id=""comments"" style=""top: 640px; left:   0px;"">" & varOps4 & "</div>")
            PrintLine(2, SPC(6), "<div id=""comments"" style=""top: 473px; left: 405px;"">" & varOps4 & "</div>")
            PrintLine(4, SPC(6), "<div id=""comments"" style=""top: 504px; left:   5px;"">" & varOps4 & "</div>")
        End If
        If varOps5 <> "" Then
            PrintLine(1, SPC(6), "<div id=""comments"" style=""top: 680px; left:   0px;"">" & varOps5 & "</div>")
            PrintLine(2, SPC(6), "<div id=""comments"" style=""top: 542px; left:   0px;"">" & varOps5 & "</div>")
            PrintLine(4, SPC(6), "<div id=""comments"" style=""top: 627px; left:   5px;"">" & varOps5 & "</div>")
        End If
        If varOps6 <> "" Then
            PrintLine(1, SPC(6), "<div id=""comments"" style=""top: 720px; left:   0px;"">" & varOps6 & "</div>")
            PrintLine(2, SPC(6), "<div id=""comments"" style=""top: 542px; left: 405px;"">" & varOps6 & "</div>")
            PrintLine(4, SPC(6), "<div id=""comments"" style=""top: 750px; left:   5px;"">" & varOps6 & "</div>")
        End If
        If varOps7 <> "" Then
            PrintLine(1, SPC(6), "<div id=""comments"" style=""top: 760px; left:   0px;"">" & varOps7 & "</div>")
            PrintLine(2, SPC(6), "<div id=""comments"" style=""top: 611px; left:   0px;"">" & varOps7 & "</div>")
            PrintLine(4, SPC(6), "<div id=""comments"" style=""top: 873px; left:   5px;"">" & varOps7 & "</div>")
        End If
        PrintLine(1, "<!-- Sups Comments End -->")
        PrintLine(2, "<!-- Sups Comments End -->")
        PrintLine(4, "<!-- Sups Comments End -->")
        PrintLine(1, "")
        PrintLine(1, "<!-- Question of the Day Start -->")
        PrintLine(1, SPC(6), "<div style=""position: absolute; top: 400px; left:  -4px""><img src=""images\fcif.jpg"" width=""308px"" height=""90px""></div>")
        PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 400px; left:  -5px; width: 300px; height: 85px;""></div>")
        PrintLine(1, SPC(6), "<div id=""boxes"" style=""top: 401px; left:  -4px; width: 298px; height: 83px; border-color: #c5885f""></div>")
        If varQoDShow = "Show" Then
            PrintLine(1, SPC(6), "<div style=""position: absolute; top: 400px; left:   0px; font-size: 22px; color: #763900; font-weight: bold""><u>Question of the Day</u></div>")
            PrintLine(1, SPC(6), "<div id=""qod"" style=""top: 425px; left:   0px;"">" & varQoD & "</div>")
            PrintLine(2, SPC(6), "<div style=""position: absolute; top: 498px; left: 825px; width: 220px; font-size: 18px; color: #763900; font-weight: bold""><u>Question of the Day</u></div>")
            PrintLine(2, SPC(6), "<div id=""qod"" style=""top: 520px; left: 825px;"">" & varQoD & "</div>")
        Else
            PrintLine(1, SPC(6), "<div style=""position: absolute; top: 402px; left:  -3px""><img src=""images\t-6small.jpg"" width=""298px"" height=""83px""></div>")
            PrintLine(2, SPC(6), "<div style=""position: absolute; top: 495px; left: 820px""><img src=""images\t6three.jpg""></div>")
        End If
        PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 495px; left: 820px; width: 275px; height: 160px;""></div>")
        PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 496px; left: 821px; width: 273px; height: 158px; border-color: #c5885f""></div>")
        PrintLine(1, "<!-- Question of the Day End -->")
        If rbBeerOn.Checked Then
            PrintLine(2, "")
            PrintLine(2, "<!-- Beer Light Start -->")
            PrintLine(2, SPC(6), "<div style=""position: absolute; top: 706px; left:  -4px""><img src=""images\sof2.jpg"" width=""1108px"" height=""49px""></div>")
            PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 705px; left: -5px; width: 1108px; height: 42px;""></div>")
            PrintLine(2, SPC(6), "<div id=""boxes"" style=""top: 706px; left: -4px; width: 1106x; height: 40px; border-color: #c5885f""></div>")
            PrintLine(2, SPC(6), "<div style=""position: absolute; top: 707px; left: 204px""><img src=""images\alert.gif"" width=""35px"" height=""35px""></div>")
            PrintLine(2, SPC(6), "<div style=""position: absolute; top: 707px; left: 855px""><img src=""images\alert.gif"" width=""35px"" height=""35px""></div>")
            subWrite(2, 6, 702, 320, 40, "#060", 460, 1, 0, "BEER LIGHT IS ON")
            PrintLine(2, "<!-- Beer Light End -->")
        End If
        PrintLine(1, "")
        PrintLine(1, SPC(4), "</div>")
        PrintLine(1, SPC(2), "</body>")
        PrintLine(1, SPC(0), "</html>")
        PrintLine(2, "")
        PrintLine(2, SPC(4), "</div>")
        PrintLine(2, SPC(2), "</body>")
        PrintLine(2, SPC(0), "</html>")
        PrintLine(4, "")
        PrintLine(4, SPC(4), "</div>")
        PrintLine(4, SPC(2), "</body>")
        PrintLine(4, SPC(0), "</html>")
        FileClose(1)
        FileClose(2)
        FileClose(4)
        FileCopy(lPath & "status.html", oPath & "status.htm")
        FileCopy(lPath & "status2.html", oPath & "status2.htm")
        FileCopy(lPath & "status4a.html", oPath & "status4a.htm")
        FileCopy(lPath & "status4b.html", oPath & "status4b.htm")
        subWriteXML()
        gblStatus = 0
        subFixButton(gblStatus)
    End Sub
    Private Sub btnInfoAdmin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInfoAdmin.Click
        Admin.Show()
    End Sub
    Private Sub btnNews_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNews.Click
        News.Show()
    End Sub
    Private Sub btnAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAbout.Click
        About.Show()
    End Sub
End Class