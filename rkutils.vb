
Imports System
Imports System.Configuration
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlTypes
Imports System.Data.SqlDbType
Imports System.Data.SqlClient
Imports System.Data.Odbc
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Web
Imports Microsoft.VisualBasic

#Const DB2 = 1
#If DB2 = 1 Then
Imports IBM.Data.DB2
#End If

Public Module rkutils

    Sub Wait(ByVal v_sec)
        Dim start As Double
        Dim finish As Double
        Dim totaltime As Double

        If TimeOfDay >= #11:59:55 PM# Then
            finish = 0
        End If

        start = Microsoft.VisualBasic.DateAndTime.Timer
        finish = start + v_sec   ' Set end time for v_sec seconds.
        Do While Microsoft.VisualBasic.DateAndTime.Timer < finish
            System.Windows.Forms.Application.DoEvents()
            System.Threading.Thread.Sleep(200)
        Loop
        totaltime = Microsoft.VisualBasic.DateAndTime.Timer - start
    End Sub

    Public Function STR_TRIM(ByVal tSTRin As String, ByVal iLEN As Integer) As String
        If tSTRin.Length >= iLEN Then
            Return tSTRin.Substring(0, iLEN).Trim
        End If
        Return tSTRin
    End Function

    Public Function STR_LEFT(ByVal tSTRin As String, ByVal iLEN As Integer) As String
        If tSTRin.Length >= iLEN Then
            Return tSTRin.Substring(0, iLEN).Trim
        End If
        Return tSTRin
    End Function

    Public Sub Msg_Error(ByVal tModule As String, ByVal tError As String)
        'MsgBox(tModule + vbCr + tError)
        aoLetters.MsgStatus(tModule + " " + tError, True)
    End Sub

    Public Function STR_RIGHT(ByVal tSTRin As String, ByVal iLEN As Integer) As String
        If tSTRin.Length >= iLEN Then
            Return tSTRin.Substring(tSTRin.Length - iLEN, iLEN).Trim
        End If
        Return tSTRin
    End Function

    Public Function STR_trim_CRLF(ByVal tLineIn As String) As String
        Dim iC As Integer
        Dim tLineOut As String = ""
        Try
            If tLineIn.Contains(vbCr) Or tLineIn.Contains(vbCrLf) Or tLineIn.Contains(Chr(10)) Or tLineIn.Contains(Chr(13)) Then
                For iC = 0 To tLineIn.Length - 1
                    Select Case tLineIn.Substring(iC, 1)
                        Case vbCr, vbCrLf, Chr(10), Chr(13)
                            tLineOut += " "
                        Case Else
                            tLineOut += tLineIn.Substring(iC, 1)
                    End Select
                Next
                Return tLineOut
            Else
                Return tLineIn
            End If
        Catch ex As Exception
            Msg_Error("STR_trim_CRLF", tLineOut + vbCr + vbCr + ex.ToString)
            Return ""
        End Try
    End Function

    Public Function STR_convert_AMP(ByVal tLineIn As String) As String
        Try
            Dim tLineOut As String = tLineIn
            tLineOut = tLineOut.Replace("&#39;", "'")
            tLineOut = tLineOut.Replace("&quot;", Chr(34))
            tLineOut = tLineOut.Replace("&nbsp;", " ")
            tLineOut = tLineOut.Replace("&lt;", "<")
            tLineOut = tLineOut.Replace("&gt;", ">")
            tLineOut = tLineOut.Replace("&amp;", "&")
            Return tLineOut
        Catch ex As Exception
            Msg_Error("STR_convert_AMP", ex.ToString)
            Return ""
        End Try
    End Function

    Public Function STR_TRAN(ByVal tLineIn As String, ByVal tFind As String, ByVal tReplace As String) As String
        Try
            Dim tLineOut As String = tLineIn
            tLineOut = tLineOut.Replace(tFind, tReplace)
            Return tLineOut
        Catch ex As Exception
            Msg_Error("STR_TRAN", ex.ToString)
            Return ""
        End Try
    End Function

    Public Function STR_LETTTERSONLY(ByVal tLineIn As String) As String
        Dim iC As Integer
        Dim tLineOut As String = ""
        Try
            For iC = 0 To tLineIn.Length - 1
                Select Case tLineIn.Substring(iC, 1)
                    Case "A" To "Z"
                        tLineOut += tLineIn.Substring(iC, 1)
                    Case "a" To "z"
                        tLineOut += tLineIn.Substring(iC, 1)
                    Case " "
                        tLineOut += tLineIn.Substring(iC, 1)
                    Case Else
                        tLineOut += ""
                End Select
            Next
            Return tLineOut
        Catch ex As Exception
            Msg_Error("STR_LETTERSONLY", tLineOut + vbCr + vbCr + ex.ToString)
            Return ""
        End Try
    End Function

    Public Function STR_BREAK(ByVal tSTRin As String, ByVal FirstOr2nd As Integer) As String
        If tSTRin.Contains(" ") Then
            If FirstOr2nd = 1 Then
                Return tSTRin.Substring(0, tSTRin.IndexOf(" ")).Trim
            Else
                Return tSTRin.Substring(tSTRin.IndexOf(" ")).Trim
            End If
        End If
        Return tSTRin
    End Function

    Public Function STR_BREAK_AT(ByVal tSTRin As String, ByVal FirstOr2nd As Integer, ByVal tBreakCharacter As String) As String
        Try
            If tSTRin.Contains(tBreakCharacter) Then
                If FirstOr2nd = 1 Then
                    Return tSTRin.Substring(0, tSTRin.IndexOf(tBreakCharacter)).Trim
                Else
                    Return tSTRin.Substring(tSTRin.IndexOf(tBreakCharacter) + 1).Trim
                End If
            End If
            Return tSTRin
        Catch ex As Exception
            Msg_Error("STR_BREAK_AT", ex.ToString)
        End Try
        Return ""
    End Function

    Public Function STR_BREAK_STR(ByVal tSTRin As String, ByVal tStartString As String, ByVal tStopString As String, ByVal iAfterStart As Integer) As String
        Try
            If tSTRin.Contains(tStartString) And tSTRin.Contains(tStopString) Then
                Dim i1 As Integer = tSTRin.IndexOf(tStartString)
                Dim i2 As Integer = tSTRin.IndexOf(tStopString)
                If iAfterStart > 0 Then i1 = i1 + tStartString.Length
                If i1 >= 0 And i2 > 0 And i2 - i1 <= tSTRin.Length Then Return tSTRin.Substring(i1, i2 - i1).Trim
            End If
            Return tSTRin
        Catch ex As Exception
            Msg_Error("STR_BREAK_STR", ex.ToString)
        End Try
        Return ""
    End Function

    Public Function SecondsToTime(ByVal sttSeconds As Long, ByVal Num2Return As Integer) As String
        Dim tHour As String, tMin As String, tSec As String
        Dim dN As Double, dM As Double, dS As Double
        Dim TmpH As Integer, tmpM As Integer, TmpS As Integer

        dN = sttSeconds
        'if(Average) if(OCtr > 0) TmpN=(TmpN/OCtr);
        TmpH = 0
        dM = (dN / 60)
        dS = dN Mod 60

        Select Case dM
            Case 60 To 117
                dM = dM - 60
                TmpH = 1
            Case 119 To 176
                dM = dM - 120
                TmpH = 2
            Case 178 To 235
                dM = dM - 180
                TmpH = 3
            Case 237 To 294
                dM = dM - 240
                TmpH = 4
            Case 296 To 353
                dM = dM - 300
                TmpH = 5
            Case 355 To 412
                dM = dM - 360
                TmpH = 6
            Case 414 To 471
                dM = dM - 420
                TmpH = 7
            Case 473 To 530
                dM = dM - 480
                TmpH = 8
            Case 532 To 589
                dM = dM - 540
                TmpH = 9
            Case 591 To 648
                dM = dM - 600
                TmpH = 10
            Case 650 To 707
                dM = dM - 660
                TmpH = 11
            Case 709 To 766
                dM = dM - 720
                TmpH = 12
            Case 768 To 825
                dM = dM - 780
                TmpH = 13
            Case 827 To 882
                dM = dM - 840
                TmpH = 14
            Case 884 To 941
                dM = dM - 900
                TmpH = 15
            Case 943 To 1000
                dM = dM - 960
                TmpH = 16
            Case 1002 To 1059
                dM = dM - 1020
                TmpH = 17
            Case 1061 To 1118
                dM = dM - 1080
                TmpH = 18
            Case 1120 To 1177
                dM = dM - 1140
                TmpH = 19
            Case 1179 To 1236
                dM = dM - 1200
                TmpH = 20
            Case 1238 To 1295
                dM = dM - 1260
                TmpH = 21
            Case 1297 To 1354
                dM = dM - 1320
                TmpH = 22
            Case 1356 To 1413
                dM = dM - 1380
                TmpH = 23
            Case 1415 To 1472
                dM = dM - 1440
                TmpH = 24
            Case 1474 To 9999
                TmpH = -1
        End Select

        tmpM = Fix(Str(dM))
        TmpS = Fix(Str(dS))
        If ((TmpS < 1) Or (TmpS > 59)) Then TmpS = 0
        If ((tmpM < 1) Or (tmpM > 59)) Then tmpM = 0
        If TmpH < 0 Then
            tHour = "++"
        Else
            If TmpH < 10 Then
                tHour = "0" + Trim(Str(TmpH))
            Else
                tHour = Trim(Str(TmpH))
            End If
        End If
        If tmpM < 10 Then
            tMin = "0" + Trim(Str(tmpM))
        Else
            tMin = Trim(Str(tmpM))
        End If
        If TmpS < 10 Then
            tSec = "0" + Trim(Str(TmpS))
        Else
            tSec = Trim(Str(TmpS))
        End If
        Select Case Num2Return
            Case 5
                SecondsToTime = tMin + ":" + tSec
            Case 8
                SecondsToTime = tHour + ":" + tMin + ":" + tSec
            Case Else
                SecondsToTime = tHour + ":" + tMin + ":" + tSec
        End Select
    End Function

    Public Function TimeToSeconds(ByVal TIMEin As String) As Long
        Dim dH As Double, dM As Double, dS As Double
        Dim TmpH As Integer, tmpM As Integer, TmpS As Integer

        If Len(Trim(TIMEin)) = 8 And Mid(TIMEin, 3, 1) = ":" Then
            TmpH = Mid(TIMEin, 1, 2)
            tmpM = Mid(TIMEin, 4, 2)
            TmpS = Mid(TIMEin, 7, 2)

            dH = (Val(TmpH) * 60) * 60
            dM = Val(tmpM) * 60
            dS = Val(TmpS)
            'TimeToSeconds = dH + dM + dS
            Return dH + dM + dS
        End If
        Return 0
    End Function

    Public Function TimeSecondsElapsed(ByVal timeInElapsed As String) As Long
        Dim dStart As Double, dNow As Double

        dStart = TimeToSeconds(timeInElapsed)
        'dNow = TimeToSeconds(Time24(8))

        TimeSecondsElapsed = dNow - dStart
    End Function

    Public Function DateDaysInMonth(ByVal dmMonth As Integer) As Integer
        Select Case dmMonth
            Case 1  'Jan
                DateDaysInMonth = 31
            Case 2  'Feb
                DateDaysInMonth = 28
            Case 3  'Mar
                DateDaysInMonth = 31
            Case 4  'Apr
                DateDaysInMonth = 30
            Case 5  'May
                DateDaysInMonth = 31
            Case 6  'June
                DateDaysInMonth = 30
            Case 7  'July
                DateDaysInMonth = 31
            Case 8  'Aug
                DateDaysInMonth = 31
            Case 9  'Sep
                DateDaysInMonth = 30
            Case 10 'Oct
                DateDaysInMonth = 31
            Case 11 'Nov
                DateDaysInMonth = 30
            Case 12 'Dec
                DateDaysInMonth = 31
            Case Else
                DateDaysInMonth = 0
        End Select
    End Function

    Public Function DateToday(ByVal Num2Return) As String
        '2008-05-20 RFK: This is a copy from rklib.vb
        Dim tSTR As String
        tSTR = ""
        Select Case Num2Return
            Case 8
                tSTR = Today.Year.ToString
                If Today.Month >= 10 Then
                    tSTR = tSTR + Today.Month.ToString
                Else
                    tSTR = tSTR + "0" + Today.Month.ToString
                End If
                If Today.Day >= 10 Then
                    tSTR = tSTR + Today.Day.ToString
                Else
                    tSTR = tSTR + "0" + Today.Day.ToString
                End If
            Case 16     'ccyymmddHHMMSSss
                tSTR = Today.Year.ToString
                If Today.Month >= 10 Then
                    tSTR = tSTR + Today.Month.ToString
                Else
                    tSTR = tSTR + "0" + Today.Month.ToString
                End If
                If Today.Day >= 10 Then
                    tSTR = tSTR + Today.Day.ToString
                Else
                    tSTR = tSTR + "0" + Today.Day.ToString
                End If
                If Now.Hour >= 10 Then
                    tSTR = tSTR + Now.Hour.ToString
                Else
                    tSTR = tSTR + "0" + Now.Hour.ToString
                End If
                If Now.Minute >= 10 Then
                    tSTR = tSTR + Now.Minute.ToString
                Else
                    tSTR = tSTR + "0" + Now.Minute.ToString
                End If
                If Now.Second >= 10 Then
                    tSTR = tSTR + Now.Second.ToString
                Else
                    tSTR = tSTR + "0" + Now.Second.ToString
                End If
                tSTR = tSTR + Today.Millisecond.ToString
            Case 18 'ccyy-MM-dd HH:mm:ss
                tSTR = Today.Year.ToString
                tSTR = tSTR + "-"
                If Today.Month >= 10 Then
                    tSTR = tSTR + Today.Month.ToString
                Else
                    tSTR = tSTR + "0" + Today.Month.ToString
                End If
                tSTR = tSTR + "-"
                If Today.Day >= 10 Then
                    tSTR = tSTR + Today.Day.ToString
                Else
                    tSTR = tSTR + "0" + Today.Day.ToString
                End If
                tSTR = tSTR + " "
                If Now.Hour >= 10 Then
                    tSTR = tSTR + Now.Hour.ToString
                Else
                    tSTR = tSTR + "0" + Now.Hour.ToString
                End If
                tSTR = tSTR + ":"
                If Now.Minute >= 10 Then
                    tSTR = tSTR + Now.Minute.ToString
                Else
                    tSTR = tSTR + "0" + Now.Minute.ToString
                End If
                tSTR = tSTR + ":"
                If Now.Second >= 10 Then
                    tSTR = tSTR + Now.Second.ToString
                Else
                    tSTR = tSTR + "0" + Now.Second.ToString
                End If
        End Select
        Return tSTR
    End Function

    Public Function TimeToday(ByVal Num2Return) As String
        '2009-01-26 RFK: This is a copy from rklib.vb
        Dim tSTR As String
        tSTR = ""
        Select Case Num2Return
            Case 5     'HH:MM   'SSss
                If Now.Hour >= 10 Then
                    tSTR += Now.Hour.ToString
                Else
                    tSTR += "0" + Now.Hour.ToString
                End If
                tSTR += ":"
                If Now.Minute >= 10 Then
                    tSTR += Now.Minute.ToString
                Else
                    tSTR += "0" + Now.Minute.ToString
                End If
            Case 6      'HHmmss
                If Now.Second >= 10 Then
                    tSTR = tSTR + Now.Second.ToString
                Else
                    tSTR = tSTR + "0" + Now.Second.ToString
                End If
                tSTR = tSTR + Today.Millisecond.ToString
        End Select
        Return tSTR
    End Function

    Public Function STR_DATE_PLUS(ByVal tSTRin As String, ByVal tFormat As String, ByVal tVALUE As String) As String
        Try
            '******************************************
            '* 2012-04-20 RFK: Check for invalid tVALUE
            If Len(tVALUE.Trim) = 0 Then Return tSTRin
            '******************************************
            Dim tSTRout As String = tSTRin

            If tSTRin = "TODAY" Then
                tSTRin = Now.Date.ToString
            End If
            If tSTRin.Length >= 8 Then
                Dim dDATE As Date = tSTRin
                Select Case tFormat
                    Case "+"
                        If tSTRin.Substring(0, 8) = "1/1/1800" Then
                            tSTRin = ""
                            tSTRout = ""
                        Else
                            dDATE = dDATE.AddDays(tVALUE)
                        End If
                    Case "+m", "+M"
                        If tSTRin.Substring(0, 8) = "1/1/1800" Then
                            tSTRin = ""
                            tSTRout = ""
                        Else
                            dDATE = dDATE.AddMonths(tVALUE)
                        End If
                    Case "-"
                        If tSTRin.Substring(0, 8) = "1/1/1800" Then
                            tSTRin = ""
                            tSTRout = ""
                        Else
                            dDATE = dDATE.AddDays(Val(tVALUE) * -1)
                        End If
                    Case "-m", "-M"
                        If tSTRin.Substring(0, 8) = "1/1/1800" Then
                            tSTRin = ""
                            tSTRout = ""
                        Else
                            dDATE = dDATE.AddMonths(Val(tVALUE) * -1)
                        End If
                End Select
                If IsDate(dDATE) Then
                    tSTRout = dDATE.Month.ToString.PadLeft(2, "0")
                    tSTRout += "/"
                    tSTRout += dDATE.Day.ToString.PadLeft(2, "0")
                    tSTRout += "/"
                    tSTRout += dDATE.Year.ToString.PadLeft(4, "0")
                End If
            End If
            Return tSTRout
        Catch ex As Exception
            Msg_Error("STR_DATE_PLUS", ex.ToString)
        End Try
        Return ""
    End Function

    Public Function FILE_delete(ByVal FileFullPath As String) As Boolean
        If My.Computer.FileSystem.FileExists(FileFullPath) Then
            My.Computer.FileSystem.DeleteFile(FileFullPath)
            If My.Computer.FileSystem.FileExists(FileFullPath) Then
                Return False
            Else
                Return True
            End If
        End If
        Return False
    End Function

    Public Function FILE_read(ByVal FileFullPath As String) As String
        If My.Computer.FileSystem.FileExists(FileFullPath) Then
            Return My.Computer.FileSystem.ReadAllText(FileFullPath)
        End If
        Return ""
    End Function

    Public Function FILE_contains(ByVal FileFullPath As String, ByVal tContains As String) As String
        Dim tSTR As String, nSTR As String
        If My.Computer.FileSystem.FileExists(FileFullPath) Then
            tSTR = My.Computer.FileSystem.ReadAllText(FileFullPath)
            If tSTR.Contains(tContains) Then
                nSTR = tSTR.Substring(tSTR.IndexOf(tContains) + Len(tContains))
                If nSTR.Contains(vbCr) Then
                    nSTR = nSTR.Substring(0, nSTR.IndexOf(vbCr))
                End If
                Return (nSTR)
            Else
                Return ""
            End If
        End If
        Return ""
    End Function

    Function FILE_create(ByVal FileFullPath As String, ByVal OverWrite As Boolean, ByVal AppendToFile As Boolean, ByVal toWrite As String) As String
        Try
            If My.Computer.FileSystem.FileExists(FileFullPath) Then
                If OverWrite Then
                    FILE_delete(FileFullPath)
                    If My.Computer.FileSystem.FileExists(FileFullPath) Then
                        Return False
                    End If
                Else
                    If AppendToFile Then
                        My.Computer.FileSystem.WriteAllText(FileFullPath, toWrite, True)
                        Return True
                    Else
                        Return False
                    End If
                End If
            End If
            My.Computer.FileSystem.WriteAllText(FileFullPath, toWrite, False)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function LineFind(ByVal tSTR As String, ByVal tFind As String) As Integer
        Dim tLine As String
        Dim i1 As Integer, iLine As Integer

        tLine = ""
        iLine = 1
        i1 = 1
        Do While i1 < Len(tSTR)
            If Mid(tSTR, i1, 1) = Chr(10) Or i1 = Len(tSTR) + 1 Then
                If tLine.Contains(tFind) Then
                    Return iLine
                End If
                If Mid(tSTR, i1 + 1, 1) = vbCr Then
                    i1 = i1 + 1
                End If

                tLine = ""
                iLine = iLine + 1
            Else
                If Mid(tSTR, i1, 1) = Chr(9) Then
                    tLine = tLine + " "
                Else
                    tLine = tLine + Mid(tSTR, i1, 1)
                End If
            End If
            i1 = i1 + 1
        Loop
        Return 0
    End Function

    Public Function LineRead(ByVal tSTR As String, ByVal iLineRead As Integer) As String
        Dim tLine As String
        Dim i1 As Integer, iLine As Integer

        tLine = ""
        iLine = 0
        i1 = 1
        Do While i1 < Len(tSTR)
            If Mid(tSTR, i1, 1) = Chr(10) Or i1 = Len(tSTR) + 1 Then
                If iLine = iLineRead Then
                    Return tLine
                End If
                If Mid(tSTR, i1 + 1, 1) = vbCr Then
                    i1 = i1 + 1
                End If
                tLine = ""
                iLine = iLine + 1
            Else
                tLine = tLine + Mid(tSTR, i1, 1)
            End If
            i1 = i1 + 1
        Loop
        Return ""
    End Function

    Public Function STR_FromLine(ByVal tSTR As String, ByVal iLineRead As Integer) As String
        Dim tLine As String
        Dim i1 As Integer, i2 As Integer, iLine As Integer

        tLine = ""
        iLine = 0
        i1 = 1
        Do While i1 < Len(tSTR)
            If Mid(tSTR, i1, 1) = Chr(10) Or i1 = Len(tSTR) + 1 Then
                If iLine = iLineRead Then
                    Return Mid(tSTR, i2)
                End If
                If Mid(tSTR, i1 + 1, 1) = vbCr Then
                    i1 = i1 + 1
                End If
                tLine = ""
                i2 = i1
                iLine = iLine + 1
            Else
                tLine = tLine + Mid(tSTR, i1, 1)
            End If
            i1 = i1 + 1
        Loop
        Return ""
    End Function

    Public Function BreakSPAN(ByVal tTDline As String) As String
        Dim i1 As Integer = 0, i2 As Integer = 0, i3 As Integer = 0

        BreakSPAN = ""
        If tTDline.Contains("</SPAN") Then
            For i1 = tTDline.Length - 8 To 0 Step -1
                If tTDline.Substring(i1, 7) = "</SPAN>" Then
                    i3 = i1 - 1
                End If
                If i3 > 0 Then
                    If tTDline.Substring(i1, 1) = ">" Then
                        i2 = i1
                        Exit For
                    End If
                End If
            Next
        End If
        If i2 > 0 And i3 > 0 And i3 < tTDline.Length - 1 Then
            BreakSPAN = tTDline.Substring(i2 + 1, i3 - i2)
        End If
    End Function

    Public Function BreakTD(ByVal tTDline As String) As String
        Dim i1 As Integer, i2 As Integer
        Dim tNewLine As String

        i1 = InStr(1, tTDline, ">")
        If i1 > 0 Then
            i2 = InStr(i1, tTDline, "<")
            If i2 > 0 Then
                tTDline = Mid(tTDline, i1 + 1, i2 - i1 - 1)
            Else
                tTDline = Mid(tTDline, i1 + 1)
            End If
        Else
            tTDline = Mid(tTDline, i1 + 1)
        End If
        If tTDline = "&nbsp;" Then
            tTDline = ""
        End If
        If tTDline.Length > 0 Then
            If tTDline.Substring(0, 1) = "<" Or tTDline.Substring(0, 1) = ">" Then
                tTDline = ""
            End If
            If tTDline.Contains("<") Then
                'MsgBox(tTDline)
            End If
        End If
        tNewLine = ""
        For i1 = 0 To tTDline.Length - 1
            If tTDline.Length >= i1 + 6 Then
                If tTDline.Substring(i1, 6) = "&nbsp;" Then
                    tNewLine = tNewLine + " "
                    i1 = i1 + 5
                Else
                    If tTDline.Substring(i1, 1) >= " " And tTDline.Substring(i1, 1) <= "z" Then
                        tNewLine = tNewLine + tTDline.Substring(i1, 1)
                    End If
                End If
            Else
                If tTDline.Substring(i1, 1) >= " " And tTDline.Substring(i1, 1) <= "z" Then
                    tNewLine = tNewLine + tTDline.Substring(i1, 1)
                Else
                    'MsgBox(tTDline.Substring(i1, 1))
                End If
            End If
        Next
        Return tNewLine
    End Function

    Public Function ComboBox_SetValue(ByVal cCombo As ComboBox, ByVal tVal As String) As Integer
        Try
            Dim i2 As Integer
            For i2 = 0 To cCombo.Items.Count - 1
                If UCase(cCombo.Items(i2).ToString) = UCase(tVal) Then
                    cCombo.SelectedIndex = i2
                    Return i2
                End If
            Next
        Catch ex As Exception
            'MSG_warning(ex.ToString)
        End Try
        Return -1
    End Function

    Public Function DataGridView_ColumnByName(ByVal gGrid As DataGridView, ByVal tColName As String) As Integer
        Try
            Dim i2 As Integer
            For i2 = 0 To gGrid.ColumnCount - 1
                If UCase(gGrid.Columns(i2).Name) = UCase(tColName) Then
                    Return i2
                End If
            Next
        Catch ex As Exception
            'MSG_warning(ex.ToString)
        End Try
        Return -1
    End Function

    Public Function DataGridView_ValueByColumnName(ByVal gGrid As DataGridView, ByVal tColName As String, ByVal iRow As Integer) As String
        Try
            Dim rc1 As Integer = DataGridView_ColumnByName(gGrid, tColName)
            If rc1 >= 0 Then
                If Len(gGrid.Item(rc1, iRow).Value) > 0 Then
                    Return gGrid.Item(rc1, iRow).Value.ToString()
                Else
                    Return ""
                End If
            End If
        Catch ex As Exception
            'MSG_warning(ex.ToString)
        End Try
        Return ""
    End Function

    Public Function DataGridView_SetValueByColumnName(ByVal gGrid As DataGridView, ByVal tColName As String, ByVal iRow As Integer, ByVal tValue As String)
        Try
            Dim rc1 As Integer = DataGridView_ColumnByName(gGrid, tColName)
            If rc1 >= 0 Then
                gGrid.Item(rc1, iRow).Value = tValue
            End If
        Catch ex As Exception
            'MSG_warning(ex.ToString)
        End Try
        Return ""
    End Function

    Public Function DataTable_SetValueByColumnName(ByVal dT As DataTable, ByVal tColName As String, ByVal iRow As Integer, ByVal tValue As String)
        Try
            Dim rc1 As Integer = DataTable_ColumnByName(dT, tColName)
            If rc1 >= 0 Then dT.Rows(rc1)(iRow) = tValue
        Catch ex As Exception
            'MSG_warning(ex.ToString)
        End Try
        Return ""
    End Function

    Public Function DataGridView_Contains(ByVal gGrid As DataGridView, ByVal tColName As String, ByVal tValue As String) As Integer
        Try
            Dim rc1 As Integer = DataGridView_ColumnByName(gGrid, tColName)
            If rc1 >= 0 Then
                For i1 = 0 To gGrid.RowCount - 1
                    If gGrid.Item(rc1, i1).Value.ToString.Trim = tValue Then Return i1
                Next
            End If
            Return -1
        Catch ex As Exception
            'MSG_warning(ex.ToString)
        End Try
        Return -1
    End Function

    Public Function DataGridView_AddColumn(ByVal gGrid As DataGridView, ByVal tColumnName As String)
        Try
            If DataGridView_ColumnByName(gGrid, tColumnName) >= 0 Then
                'Nothing
            Else
                gGrid.Columns.Add(tColumnName, tColumnName)
            End If
            Return True
        Catch ex As Exception
            'MSG_warning(ex.ToString)
        End Try
        Return False
    End Function

    Public Function WhoAmI() As String
        'Dim tUSR As String = HttpContext.Current.User.Identity.Name.ToString
        Dim tUSR As String = My.User.Name
        Dim lSlash As Integer = tUSR.LastIndexOf("\")
        If lSlash > 0 Then
            tUSR = tUSR.Substring(lSlash + 1)
        End If
        Return tUSR
    End Function

    Public Function DateOrBlank(ByVal tDateIn As String, ByVal tFormat As String) As String
        Try
            If IsDate(tDateIn) Then
                Return STR_format(tDateIn, tFormat)
            End If
            Return ""
        Catch ex As Exception
            'Nothing
        End Try
        Return ""
    End Function

    Public Function STR_format(ByVal tSTRin As String, ByVal tFormat As String) As String
        Try
            If tSTRin.Trim.Length = 0 Then Return ""
            If tSTRin = "TODAY" Then
                If tFormat.Contains("HH") Then
                    tSTRin = Now.Month.ToString + "/" + Now.Day.ToString + "/" + Now.Year.ToString + " " + Now.Hour.ToString + ":" + Now.Minute.ToString + ":" + Now.Second.ToString
                Else
                    tSTRin = Now.Month.ToString + "/" + Now.Day.ToString + "/" + Now.Year.ToString
                End If
            Else
                If tSTRin = "YESTERDAY" Then
                    tSTRin = STR_DATE_PLUS("TODAY", "-", "1")
                End If
            End If
            Dim tSTRout As String = tSTRin
            Select Case tFormat
                Case "ISVALID", "VALID"
                    tSTRout = ""
                    Dim iVAL As Integer
                    For i1 = 0 To tSTRin.Length - 1
                        iVAL = Asc(tSTRin.Substring(i1, 1))
                        Select Case iVAL
                            Case Asc("a") To Asc("z")
                                tSTRout += tSTRin.Substring(i1, 1)
                            Case Asc("A") To Asc("Z")
                                tSTRout += tSTRin.Substring(i1, 1)
                            Case Asc("0") To Asc("9")
                                tSTRout += tSTRin.Substring(i1, 1)
                            Case Asc(" "), Asc("_"), Asc("#"), Asc("$"), Asc("%"), Asc("&"), Asc("("), Asc(")"), Asc("-"), Asc("+"), Asc("."), Asc("?"), Asc("<"), Asc(">"), Asc("="), Asc("@"), Asc(":")
                                tSTRout += tSTRin.Substring(i1, 1)
                            Case Else
                                If tFormat = "ISVALID" Then
                                    Return ""
                                End If
                                tSTRout += "_"
                        End Select
                    Next
                Case "$"
                    tSTRout = String.Format("{0:C}", Val(tSTRin))
                Case "#"
                    tSTRout = String.Format("{0:#.##}", Val(tSTRin))
                Case "0"
                    tSTRout = String.Format("{0:0.##}", Val(tSTRin))
                Case "0.00"
                    tSTRout = String.Format("{0:0.00}", Val(tSTRin))
                Case "#,##0.00"
                    tSTRout = String.Format("{0:###,###,##0.00}", Val(tSTRin))
                Case "#,###"
                    tSTRout = String.Format("{0:###,###,###}", Val(tSTRin))
                Case "000"
                    tSTRout = String.Format("{0:0.00}", Val(tSTRin)).Replace(".", "").Replace(",", "")
                Case "PHONE"
                    If tSTRin.Length = 10 Then
                        tSTRout = tSTRin.Substring(0, 3) + "-" + tSTRin.Substring(3, 3) + "-" + tSTRin.Substring(6, 4)
                    End If
                Case "SSN"
                    If tSTRin.Length = 9 Then
                        tSTRout = tSTRin.Trim.Substring(0, 3) + "-" + tSTRin.Trim.Substring(3, 2) + "-" + tSTRin.Trim.Substring(5, 4)
                    Else
                        If tSTRin.Length >= 4 Then
                            tSTRout = "000-00-" + rkutils.STR_RIGHT(tSTRin.Trim, 4)
                        Else
                            tSTRout = "000-00-0000"
                        End If
                    End If
                Case "mm"
                    If tSTRin = "01011900" Then
                        tSTRout = ""
                    Else
                        Dim dDATE As Date = tSTRin
                        tSTRout = ""
                        If dDATE.Month >= 10 Then
                            tSTRout += dDATE.Month.ToString
                        Else
                            tSTRout += "0" + dDATE.Month.ToString
                        End If
                    End If
                Case "dd"
                    If tSTRin = "01011900" Then
                        tSTRout = ""
                    Else
                        Dim dDATE As Date = tSTRin
                        tSTRout = ""
                        If dDATE.Day >= 10 Then
                            tSTRout += dDATE.Day.ToString
                        Else
                            tSTRout += "0" + dDATE.Day.ToString
                        End If
                    End If
                Case "yy"
                    If tSTRin = "01011900" Then
                        tSTRout = ""
                    Else
                        Dim dDATE As Date = tSTRin
                        tSTRout = Right(dDATE.Year.ToString, 2)
                    End If
                Case "ccyy"
                    If tSTRin = "01011900" Then
                        tSTRout = ""
                    Else
                        Dim dDATE As Date = tSTRin
                        tSTRout = dDATE.Year.ToString
                    End If
                Case "yymmdd"
                    If tSTRin = "01011900" Then
                        tSTRout = ""
                    Else
                        Dim dDATE As Date = tSTRin
                        tSTRout = Right(dDATE.Year.ToString, 2)
                        tSTRout += dDATE.Month.ToString.PadLeft(2, "0")
                        tSTRout += dDATE.Day.ToString.PadLeft(2, "0")
                    End If
                Case "ccyy-mm"
                    If tSTRin = "01011900" Then
                        tSTRout = ""
                    Else
                        Dim dDATE As Date = tSTRin
                        tSTRout = Right(dDATE.Year.ToString, 4)
                        tSTRout += "-"
                        tSTRout += dDATE.Month.ToString.PadLeft(2, "0")
                    End If
                Case "ccyymmdd"
                    If tSTRin = "01011900" Then
                        tSTRout = ""
                    Else
                        If IsDate(tSTRin) Then
                            Dim dDATE As Date = tSTRin
                            tSTRout = dDATE.Year.ToString
                            If dDATE.Month >= 10 Then
                                tSTRout += dDATE.Month.ToString
                            Else
                                tSTRout += "0" + dDATE.Month.ToString
                            End If
                            If dDATE.Day >= 10 Then
                                tSTRout += dDATE.Day.ToString
                            Else
                                tSTRout += "0" + dDATE.Day.ToString
                            End If
                        End If
                    End If
                Case "ccyy-mm-dd"
                    If tSTRin = "01011900" Then
                        tSTRout = ""
                    Else
                        If IsDate(tSTRin) Then
                            Dim dDATE As Date = tSTRin
                            tSTRout = dDATE.Year.ToString
                            tSTRout += "-"
                            If dDATE.Month >= 10 Then
                                tSTRout += dDATE.Month.ToString
                            Else
                                tSTRout += "0" + dDATE.Month.ToString
                            End If
                            tSTRout += "-"
                            If dDATE.Day >= 10 Then
                                tSTRout += dDATE.Day.ToString
                            Else
                                tSTRout += "0" + dDATE.Day.ToString
                            End If
                        End If
                    End If
                Case "mm-dd"
                    If tSTRin = "01011900" Then
                        tSTRout = ""
                    Else
                        If IsDate(tSTRin) Then
                            Dim dDATE As Date = tSTRin
                            tSTRout = dDATE.Month.ToString.PadLeft(2, "0")
                            tSTRout += "-"
                            tSTRout += dDATE.Day.ToString.PadLeft(2, "0")
                        End If
                    End If
                Case "mm/dd/ccyy"
                    If tSTRin.Length >= 8 Then
                        If tSTRin.Substring(0, 8) = "1/1/1800" Then 'Or tSTRin.Substring(0, 8) = "1/1/1900" Then
                            tSTRout = ""
                        Else
                            If tSTRin.Contains("/") = False And tSTRin.Substring(0, 2) = "19" Or tSTRin.Substring(0, 2) = "20" Then
                                tSTRin = tSTRin.Substring(4, 2) + "/" + tSTRin.Substring(6, 2) + "/" + tSTRin.Substring(0, 4)
                            End If
                            If IsDate(tSTRin) Then
                                Dim dDATE As Date = tSTRin
                                tSTRout = dDATE.Month.ToString.PadLeft(2, "0")
                                tSTRout += "/"
                                tSTRout += dDATE.Day.ToString.PadLeft(2, "0")
                                tSTRout += "/"
                                tSTRout += dDATE.Year.ToString
                            End If
                        End If
                    End If
                Case "mm/dd/yy"
                    If tSTRin.Length >= 8 Then
                        If tSTRin.Substring(0, 8) = "1/1/1800" Then 'Or tSTRin.Substring(0, 8) = "1/1/1900" Then
                            tSTRout = ""
                        Else
                            If tSTRin.Contains("/") = False And tSTRin.Substring(0, 2) = "19" Or tSTRin.Substring(0, 2) = "20" Then
                                tSTRin = tSTRin.Substring(4, 2) + "/" + tSTRin.Substring(6, 2) + "/" + tSTRin.Substring(0, 4)
                            End If
                            If IsDate(tSTRin) Then
                                Dim dDATE As Date = tSTRin
                                tSTRout = dDATE.Month.ToString.PadLeft(2, "0")
                                tSTRout += "/"
                                tSTRout += dDATE.Day.ToString.PadLeft(2, "0")
                                tSTRout += "/"
                                tSTRout += STR_RIGHT(dDATE.Year.ToString, 2)
                            End If
                        End If
                    End If
                Case "mmddyy"
                    If tSTRin.Length >= 8 Then
                        If tSTRin.Substring(0, 8) = "1/1/1800" Then 'Or tSTRin.Substring(0, 8) = "1/1/1900" Then
                            tSTRout = ""
                        Else
                            If tSTRin.Contains("/") = False And tSTRin.Substring(0, 2) = "19" Or tSTRin.Substring(0, 2) = "20" Then
                                tSTRin = tSTRin.Substring(4, 2) + "/" + tSTRin.Substring(6, 2) + "/" + tSTRin.Substring(0, 4)
                            End If
                            If IsDate(tSTRin) Then
                                Dim dDATE As Date = tSTRin
                                tSTRout = dDATE.Month.ToString.PadLeft(2, "0")
                                tSTRout += dDATE.Day.ToString.PadLeft(2, "0")
                                tSTRout += STR_RIGHT(dDATE.Year.ToString, 2)
                            End If
                        End If
                    End If
                Case "ccyy/mm/dd"
                    If tSTRin = "01011900" Then
                        tSTRout = ""
                    Else
                        If IsDate(tSTRin) Then
                            If IsDate(tSTRin) Then
                                Dim dDATE As Date = tSTRin
                                tSTRout = dDATE.Year.ToString
                                tSTRout += "/"
                                If dDATE.Month >= 10 Then
                                    tSTRout += dDATE.Month.ToString
                                Else
                                    tSTRout += "0" + dDATE.Month.ToString
                                End If
                                tSTRout += "/"
                                If dDATE.Day >= 10 Then
                                    tSTRout += dDATE.Day.ToString
                                Else
                                    tSTRout += "0" + dDATE.Day.ToString
                                End If
                            End If
                        End If
                    End If
                Case "mmddccyy"
                    If tSTRin.Length >= 8 Then
                        If IsDate(tSTRin) Then
                            If tSTRin.Substring(0, 8) = "1/1/1800" Then 'Or tSTRin.Substring(0, 8) = "1/1/1900" Then
                                tSTRout = ""
                            Else
                                If IsDate(tSTRin) Then
                                    Dim dDATE As Date = tSTRin
                                    If dDATE.Month >= 10 Then
                                        tSTRout = dDATE.Month.ToString
                                    Else
                                        tSTRout = "0" + dDATE.Month.ToString
                                    End If
                                    If dDATE.Day >= 10 Then
                                        tSTRout += dDATE.Day.ToString
                                    Else
                                        tSTRout += "0" + dDATE.Day.ToString
                                    End If
                                    tSTRout += dDATE.Year.ToString
                                End If
                            End If
                        End If
                    End If
                Case "ccyymmddHHMMSS"
                    Dim dDATE As Date = tSTRin
                    tSTRout = dDATE.Year.ToString
                    tSTRout += dDATE.Month.ToString.PadLeft(2, "0")
                    tSTRout += dDATE.Day.ToString.PadLeft(2, "0")
                    tSTRout += dDATE.Hour.ToString.PadLeft(2, "0")
                    tSTRout += dDATE.Minute.ToString.PadLeft(2, "0")
                    tSTRout += dDATE.Second.ToString.PadLeft(2, "0")
                    tSTRout = tSTRout.Substring(0, 14)
                Case "ccyymmddHHMMSSss"
                    Dim dDATE As Date = tSTRin
                    tSTRout = dDATE.Year.ToString
                    tSTRout += dDATE.Month.ToString.PadLeft(2, "0")
                    tSTRout += dDATE.Day.ToString.PadLeft(2, "0")
                    tSTRout += dDATE.Hour.ToString.PadLeft(2, "0")
                    tSTRout += dDATE.Minute.ToString.PadLeft(2, "0")
                    tSTRout += dDATE.Second.ToString.PadLeft(2, "0")
                    tSTRout += dDATE.Millisecond.ToString.PadLeft(2, "0")
                    tSTRout = tSTRout.Substring(0, 16)
                Case "ccyy-mm-dd HH:MM:SS"
                    Dim dDATE As Date = tSTRin
                    tSTRout = dDATE.Year.ToString
                    tSTRout += "-"
                    If dDATE.Month >= 10 Then
                        tSTRout += dDATE.Month.ToString
                    Else
                        tSTRout += "0" + dDATE.Month.ToString
                    End If
                    tSTRout += "-"
                    If dDATE.Day >= 10 Then
                        tSTRout += dDATE.Day.ToString
                    Else
                        tSTRout += "0" + dDATE.Day.ToString
                    End If
                    tSTRout += " "
                    If dDATE.Hour >= 10 Then
                        tSTRout += dDATE.Hour.ToString
                    Else
                        tSTRout += "0" + dDATE.Hour.ToString
                    End If
                    tSTRout += ":"
                    If dDATE.Minute >= 10 Then
                        tSTRout += dDATE.Minute.ToString
                    Else
                        tSTRout += "0" + dDATE.Minute.ToString
                    End If
                    tSTRout += ":"
                    If dDATE.Second >= 10 Then
                        tSTRout += dDATE.Second.ToString
                    Else
                        tSTRout += "0" + dDATE.Second.ToString
                    End If
                Case "mm/dd/ccyy HH:MM:SS"
                    Dim dDATE As Date = tSTRin
                    tSTRout = ""
                    If dDATE.Month >= 10 Then
                        tSTRout += dDATE.Month.ToString
                    Else
                        tSTRout += "0" + dDATE.Month.ToString
                    End If
                    tSTRout += "/"
                    If dDATE.Day >= 10 Then
                        tSTRout += dDATE.Day.ToString
                    Else
                        tSTRout += "0" + dDATE.Day.ToString
                    End If
                    tSTRout += "/"
                    tSTRout += dDATE.Year.ToString
                    tSTRout += " "
                    If dDATE.Hour >= 10 Then
                        tSTRout += dDATE.Hour.ToString
                    Else
                        tSTRout += "0" + dDATE.Hour.ToString
                    End If
                    tSTRout += ":"
                    If dDATE.Minute >= 10 Then
                        tSTRout += dDATE.Minute.ToString
                    Else
                        tSTRout += "0" + dDATE.Minute.ToString
                    End If
                    tSTRout += ":"
                    If dDATE.Second >= 10 Then
                        tSTRout += dDATE.Second.ToString
                    Else
                        tSTRout += "0" + dDATE.Second.ToString
                    End If
                Case "HH:MM", "HH:MM:SS"
                    Dim dDATE As Date
                    If tSTRin.Contains("/") Then
                        Dim dDATE2 As Date = tSTRin
                        dDATE = dDATE2
                    Else
                        If tSTRin.Contains(":") Then
                            Dim dDATE2 As Date = Today.Month.ToString + "/" + Today.Day.ToString + "/" + Today.Year.ToString + " " + tSTRin
                            dDATE = dDATE2
                        Else
                            If tSTRin.Length > 1 Then
                                Dim dDATE2 As Date = Today.Month.ToString + "/" + Today.Day.ToString + "/" + Today.Year.ToString + " " + Left(tSTRin, tSTRin.Length - 2) + ":" + Right(tSTRin, 2)
                                If IsDate(dDATE2) Then
                                    dDATE = dDATE2
                                End If
                            End If
                        End If
                    End If
                    If dDATE.Hour >= 10 Then
                        tSTRout = dDATE.Hour.ToString
                    Else
                        tSTRout = "0" + dDATE.Hour.ToString
                    End If
                    tSTRout += ":"
                    If dDATE.Minute >= 10 Then
                        tSTRout += dDATE.Minute.ToString
                    Else
                        tSTRout += "0" + dDATE.Minute.ToString
                    End If
                    If tFormat = "HH:MM:SS" Then
                        tSTRout += ":"
                        If dDATE.Second >= 10 Then
                            tSTRout += dDATE.Second.ToString
                        Else
                            tSTRout += "0" + dDATE.Second.ToString
                        End If
                    End If
                Case "HHMMSS"
                    Dim dDATE As Date
                    If IsDate(tSTRin) Then
                        Dim dDATE2 As Date = tSTRin
                        dDATE = dDATE2
                        tSTRout = dDATE.Hour.ToString.PadLeft(2, "0")
                        tSTRout += dDATE.Minute.ToString.PadLeft(2, "0")
                        tSTRout += dDATE.Second.ToString.PadLeft(2, "0")
                    End If
                Case "HH"
                    Dim dDATE As Date
                    If IsDate(tSTRin) Then
                        Dim dDATE2 As Date = tSTRin
                        dDATE = dDATE2
                        tSTRout = dDATE.Hour.ToString.PadLeft(2, "0")
                    End If
                Case "MM"
                    Dim dDATE As Date
                    If IsDate(tSTRin) Then
                        Dim dDATE2 As Date = tSTRin
                        dDATE = dDATE2
                        tSTRout = dDATE.Minute.ToString.PadLeft(2, "0")
                    End If
                Case "SS"
                    Dim dDATE As Date
                    If IsDate(tSTRin) Then
                        Dim dDATE2 As Date = tSTRin
                        dDATE = dDATE2
                        tSTRout = dDATE.Second.ToString.PadLeft(2, "0")
                    End If
                Case "AGE_D"
                    If tSTRin.Length >= 8 Then
                        If tSTRin.Substring(0, 8) = "1/1/1800" Then
                            tSTRout = ""
                        Else
                            Dim dDATE As Date = tSTRin
                            tSTRout = Now.Date.Subtract(dDATE).Days.ToString
                        End If
                    Else
                        tSTRout = ""
                    End If
                Case "AGE_Y"
                    If tSTRin.Length >= 8 Then
                        If tSTRin.Substring(0, 8) = "1/1/1800" Then
                            tSTRout = ""
                        Else
                            Dim dDATE As Date = tSTRin
                            tSTRout = Now.Date.Subtract(dDATE).Days.ToString
                            If Val(tSTRout) > 365 Then
                                tSTRout = (Val(tSTRout) / 365).ToString.Trim
                                If tSTRout.Contains(".") Then
                                    tSTRout = tSTRout.Substring(0, tSTRout.IndexOf("."))
                                End If
                            End If
                        End If
                    Else
                        tSTRout = ""
                    End If
                Case "DOW"
                    Dim dDATE As Date = tSTRin
                    tSTRout = dDATE.DayOfWeek.ToString
            End Select
            Return tSTRout
        Catch ex As Exception
            Msg_Error("STR_format", ex.ToString)
        End Try
        Return ""
    End Function

    Public Function SQL_READ_FIELD(ByVal gGrid As DataGridView, ByVal tDB As String, ByVal tFIELDNAME As String, ByVal tConnectionString As String, ByVal tSQLuser As String, ByVal tSELECTstring As String) As String
        Try
            If SQL_READ_DATAGRID(gGrid, tDB, tFIELDNAME, tConnectionString, tSQLuser, tSELECTstring) Then
                Return DataGridView_ValueByColumnName(gGrid, tFIELDNAME, 0)
            End If
            Return ""
        Catch ex As Exception
            Msg_Error("SQL_READ_FIELD", ex.ToString)
        End Try
        Return ""
    End Function

    Public Function SQL_READ_FIELD_DataTable(ByVal dT As DataTable, ByVal tDB As String, ByVal tFIELDNAME As String, ByVal tConnectionString As String, ByVal tSQLuser As String, ByVal tSELECTstring As String) As String
        Try
            If SQL_READ_DATATABLE(dT, tDB, tFIELDNAME, tConnectionString, tSQLuser, tSELECTstring) Then
                Return DataTable_ValueByColumnName(dT, tFIELDNAME, 0)
            End If
            Return ""
        Catch ex As Exception
            Msg_Error("SQL_READ_FIELD", ex.ToString)
        End Try
        Return ""
    End Function

    Public Function SQL_READ_DATAGRID(ByVal gGRID As DataGridView, ByVal tDB As String, ByVal tMODULE As String, ByVal tSQLConnectionString As String, ByVal tSQLuser As String, ByVal tSELECTstring As String) As Boolean
        Try
            Select Case tDB.ToUpper
                Case "DB2"
#If DB2 = 1 Then
                    Dim dbConnection As IBM.Data.DB2.iSeries.iDB2Connection = New IBM.Data.DB2.iSeries.iDB2Connection(tSQLConnectionString + tSQLuser)
                    Dim dbCommand As IBM.Data.DB2.iSeries.iDB2Command = New IBM.Data.DB2.iSeries.iDB2Command()
                    dbCommand.CommandText = tSELECTstring
                    dbCommand.Connection = dbConnection
                    dbCommand.CommandTimeout = 0

                    Dim dataAdapter As IBM.Data.DB2.iSeries.iDB2DataAdapter = New IBM.Data.DB2.iSeries.iDB2DataAdapter
                    dataAdapter.SelectCommand = dbCommand

                    Dim dataSet As System.Data.DataSet = New System.Data.DataSet
                    dataAdapter.Fill(dataSet, "temp")

                    gGRID.DataSource = dataSet.Tables(0)
                    'gGRID.DataBind()
                    gGRID.Visible = False
                    dbCommand.Dispose()
                    dbCommand = Nothing
                    dbConnection.Close()
                    dbConnection.Dispose()
                    dbConnection = Nothing

                    If gGRID.Rows.Count > 0 Then
                        Return True
                    End If
#Else
                    eLetters.MsgStatus("DB2 not defined", True)
#End If
                Case "MSSQL"
                    Dim dbConnection As New SqlConnection(tSQLConnectionString + tSQLuser)          ' The SqlConnection class allows you to communicate with SQL Server.
                    Dim dbCommand As New SqlCommand(tSELECTstring, dbConnection)                    ' A SqlCommand object is used to execute the SQL commands.
                    dbCommand.CommandTimeout = 0

                    Dim da As New SqlDataAdapter(dbCommand)
                    Dim mDataSet As New DataSet()
                    da.Fill(mDataSet, "temp")


                    gGRID.DataSource = mDataSet.Tables(0)
                    'gGRID.DataBind()
                    gGRID.Visible = False
                    dbCommand.Dispose()
                    dbCommand = Nothing
                    dbConnection.Close()
                    dbConnection.Dispose()
                    dbConnection = Nothing
                    If gGRID.Rows.Count > 0 Then
                        Return True
                    End If
                    'Case "MYSQL"
                    '    Dim dbConnection As New MySqlConnection(tSQLConnectionString + tSQLuser)    'The SqlConnection class allows you to communicate with SQL Server.
                    '    Dim dbCommand As New MySqlCommand(tSELECTstring, dbConnection)            'A SqlCommand object is used to execute the SQL commands.

                    '    Dim da As New MySqlDataAdapter(dbCommand)
                    '    Dim mDataSet As New DataSet()
                    '    da.Fill(mDataSet, "temp")

                    '    gGRID.DataSource = mDataSet.Tables(0)
                    '    'gGRID.DataBind()
                    '    gGRID.Visible = False
                    '    dbCommand.Dispose()
                    '    dbCommand = Nothing
                    '    dbConnection.Close()
                    '    dbConnection.Dispose()
                    '    dbConnection = Nothing
                    '    If gGRID.Rows.Count > 0 Then
                    '        Return True
                    '    End If
                Case "FOXPRO"
                    Dim dbConnection As New OleDbConnection("Provider=vfpoledb.1;Data Source=" + tSQLConnectionString + ";Collating Sequence=machine")
                    Dim dbCommand As New OleDbCommand

                    Dim dbDataAdapter As New OleDbDataAdapter
                    Dim dbDataTable As New DataTable

                    dbCommand.CommandText = tSELECTstring
                    dbCommand.Connection = dbConnection

                    dbDataAdapter.SelectCommand = dbCommand
                    dbDataAdapter.Fill(dbDataTable)

                    gGRID.DataSource = dbDataTable
                    'gGRID.DataBind()
                    gGRID.Visible = False
                    dbCommand.Dispose()
                    dbCommand = Nothing
                    dbConnection.Close()
                    dbConnection.Dispose()
                    dbConnection = Nothing
                    If gGRID.Rows.Count > 0 Then
                        Return True
                    End If
            End Select
        Catch ex As Exception
            'Msg_Error("SQL_READ_DATAGRID" + ex.ToString)
        End Try
        Return False
    End Function


    Public Sub COMMAND_STATUS(ByVal SQLConnectionString As String, ByVal SQLuser As String, ByVal tLOCX As String, ByVal tSTATUS As String, ByVal tCC As String, ByVal tRAC As String, ByVal tMC As String)
        Try
            Dim SQLstring As String = ""
            Select Case aoLetters.sSITE
                Case "iTeleCollect"
                    SQLstring = "INSERT INTO iTeleCollect.dbo.commands "
                Case Else
                    SQLstring = "INSERT INTO RevMD.dbo.commands "
            End Select
            SQLstring += " (COMMAND, LOCX, STATUS, CC, RAC, MC, MODIFIED_BY, MODIFIED_DATE, INSERT_DATE)"
            SQLstring += " VALUES("
            SQLstring += "'STATUS'"
            SQLstring += ", '" + tLOCX + "'"
            SQLstring += ", '" + tSTATUS + "'"
            SQLstring += ", '" + tCC + "'"
            SQLstring += ", '" + tRAC + "'"
            SQLstring += ", '" + tMC + "'"
            SQLstring += ", '" + WhoAmI() + "'"
            SQLstring += ", '" + Date.Now.ToString + "'"
            SQLstring += ", '" + Date.Now.ToString + "'"
            SQLstring += ")"
            DB_COMMAND("MSSQL", SQLConnectionString, SQLuser, SQLstring)
        Catch ex As Exception
            Msg_Error("COMMAND_STATUS", ex.ToString)
        End Try
    End Sub

    Public Function NOTES_ADD(ByVal tDB As String, ByVal SQLConnectionString As String, ByVal SQLuser As String, ByVal User400 As String, ByVal gGRID As DataGridView, ByVal tLOCX As String, ByVal tNUM As String, ByVal tMSGC As String, ByVal tSTAT As String, ByVal tContactCode As String, ByVal tMESSAGE As String) As Boolean
        Try
            Select Case tDB
                Case "DB2"
#If DB2 = 1 Then

                    Dim iMSGnumber As Integer = rkutils.NOTES_MAXNUMBER(tDB, SQLConnectionString, SQLuser, User400, gGRID, tLOCX) + 1
                    Dim tSTATprevious As String = rkutils.SQL_READ_FIELD(gGRID, tDB, "RAPSTA", SQLConnectionString, SQLuser, " SELECT RAPSTA FROM ROIDATA.RACCTP WHERE RALOCX='" + tLOCX + "'")

                    Dim db2SQLCommandString As String = ""
                    db2SQLCommandString += "INSERT INTO ROIDATA.RGMSGP "
                    db2SQLCommandString += " (RGLOCX, RGMSG#, RGMSGC, RGMON, RGDAY, RGYEAR, RGTIME, RGUSER, RGMSG, RGFR30, RGRNA, RGPRST, RGSTAT, RGCOCD, RGLNA)"
                    db2SQLCommandString += String.Format(" VALUES({0}, {1}, '{2}', {3}, {4}, {5}, {6}, '{7}', '{8}', '  REVMD', 0, '{9}', '{10}', '{11}', {12})",
                        tLOCX, iMSGnumber, rkutils.STR_TRIM(tMSGC, 2), Date.Today.Month.ToString, Date.Today.Day.ToString(), Date.Today.Year.ToString(),
                        String.Format("{0:HHmm}", Now), Left(User400, 6),
                        rkutils.STR_TRIM(rkutils.STR_format(tMESSAGE, "VALID"), 87),
                        rkutils.STR_TRIM(tSTATprevious, 3), rkutils.STR_TRIM(tSTAT, 3), "R", "0")

                    Dim dbConnection As IBM.Data.DB2.iSeries.iDB2Connection = New IBM.Data.DB2.iSeries.iDB2Connection(SQLConnectionString + SQLuser)
                    Dim dbCommand As IBM.Data.DB2.iSeries.iDB2Command = New IBM.Data.DB2.iSeries.iDB2Command()
                    dbCommand.CommandText = db2SQLCommandString

                    rkutils.DB_COMMAND(tDB, SQLConnectionString, SQLuser, db2SQLCommandString)
                    rkutils.NOTES_PLUS(tDB, SQLConnectionString, SQLuser, User400, gGRID, tLOCX, "M")
                    Return True
#Else
                    eLetters.MsgStatus("DB2 not defined", True)
#End If
                Case "MSSQL"
                    Dim msSQLCommandString As String = ""
            End Select
        Catch ex As Exception
            Msg_Error("NOTES_ADD", ex.ToString)
        End Try
        Return False
    End Function

    Public Function NOTES_MAXNUMBER(ByVal tDB As String, ByVal SQLConnectionString As String, ByVal SQLuser As String, ByVal User400 As String, ByVal gGRID As DataGridView, ByVal tLOCX As String) As Integer
        Try
            Dim tMaxNo As String = SQL_READ_FIELD(gGRID, tDB, "MAXNO", SQLConnectionString, SQLuser, "SELECT MAX(RGMSG#) AS MAXNO FROM ROIDATA.RGMSGP WHERE RGLOCX='" + tLOCX + "'")
            Return Val(tMaxNo)
        Catch ex As Exception
            Msg_Error("NOTES_MAXNUMBER", ex.ToString)
        End Try
        Return False
    End Function

    Public Sub NOTES_PLUS(ByVal tDB As String, ByVal SQLConnectionString As String, ByVal SQLuser As String, ByVal User400 As String, ByVal gGRID As DataGridView, ByVal tLOCX As String, ByVal InOrOut As String)
        Try
            '**************************************************************************************************************
            '* 2011-08-23 RFK: The RACCTP RAMSGS must contain the EXACT number of NOTES in it, so it is displayed correctly
            '* 2012-01-12 RFK: moved this modified version to RKUTILS for better LIBRARY STANDARDIZATION
            Select Case tDB
                Case "DB2"
                    Dim iMSGnumber As Integer = rkutils.NOTES_MAXNUMBER(tDB, SQLConnectionString, SQLuser, User400, gGRID, tLOCX)
                    Dim db2SQLCommandString As String = ""
                    If iMSGnumber > 0 Then
                        db2SQLCommandString = "UPDATE ROIDATA.RACCTP SET RAMSGS = " + iMSGnumber.ToString
                    Else
                        db2SQLCommandString = "UPDATE ROIDATA.RACCTP SET RAMSGS = 1"
                    End If
                    Select Case InOrOut
                        Case "O"
                            db2SQLCommandString = ", RAOUT# = RAOUT# + 1"
                    End Select
                    ', RAATMP = RAATMP + 1, RAOUT# = RAOUT# + 1, RATOTC = RATOTC + 1, RACHGI = 'Y' ")
                    db2SQLCommandString += " WHERE RALOCX='" + tLOCX + "'"

                    rkutils.DB_COMMAND(tDB, SQLConnectionString, SQLuser, db2SQLCommandString)
                Case "MSSQL"
                    Dim msSQLCommandString As String = ""
            End Select
        Catch ex As Exception
            Msg_Error("NOTES_PLUS", ex.ToString)
        End Try
    End Sub

    Public Function IS_File(ByVal FileFullPath As String) As Boolean
        If FileFullPath.Contains("*") Then
            'Dim tSTR As String = My.Computer.FileSystem.GetFiles(FileFullPath, FileIO.SearchOption.SearchTopLevelOnly, "*.*")
            'Dim tSTRING As string[] files = Directory.GetFiles("D:\Documents and Settings\Lou\My Documents\Visual Studio", "*.txt");
        End If
        Return My.Computer.FileSystem.FileExists(FileFullPath)
    End Function

    Public Function Listbox_Contains(ByVal lList As ListBox, ByVal tValue As String, ByVal bAdd As Boolean) As Integer
        Dim tLIST As String = ""
        For i1ctr = 0 To lList.Items.Count - 1
            tLIST = lList.Items(i1ctr).ToString
            'MsgBox(i1ctr.ToString + vbCr + "List[" + tLIST + "]" + vbCr + "Value[" + tValue + "]")
            If tLIST = tValue Then
                Return i1ctr
            Else
                If tLIST.Contains(" ") Then
                    If tLIST.Substring(0, tLIST.IndexOf(" ")) = tValue Then
                        Return i1ctr
                    End If
                End If
            End If
        Next
        If bAdd And tValue.Trim.Length > 0 Then
            lList.Items.Add(tValue)
        End If
        Return -1
    End Function

    Public Function Listbox_Text_Contains(ByVal lList As ListBox, ByVal tValue As String, ByVal bAdd As Boolean) As Boolean
        '*************************************************************************************
        '* 2012-01-10 RFK: If any of the WORDS within the tValue are in lList then return TRUE
        Try
            Dim i1ctr As Integer, i2ctr As Integer = 0
            If tValue.Contains(" ") Then
                For i1ctr = 0 To tValue.Length
                    If tValue.Substring(i1ctr, 1) = " " Then
                        If Listbox_Contains(lList, tValue.Substring(i2ctr, i1ctr - i2ctr), False) >= 0 Then Return True
                        If Listbox_Contains(lList, UCase(tValue.Substring(i2ctr, i1ctr - i2ctr)), False) >= 0 Then Return True
                        i2ctr = i1ctr + 1
                    End If
                Next
                'Check it 1 more time for end of line
                If Listbox_Contains(lList, tValue.Substring(i2ctr, i1ctr - i2ctr), False) >= 0 Then Return True
                If Listbox_Contains(lList, UCase(tValue.Substring(i2ctr, i1ctr - i2ctr)), False) >= 0 Then Return True
            Else
                Return Listbox_Contains(lList, tValue, False)
            End If
            If bAdd And tValue.Trim.Length > 0 Then
                lList.Items.Add(tValue)
            End If
        Catch ex As Exception
            '
        End Try
        Return False
    End Function

    Public Function Listbox_Value(ByVal lList As ListBox, ByVal tValue As String, ByVal BreakSpace As Boolean) As String
        Dim tLIST As String = ""
        For i1ctr = 0 To lList.Items.Count - 1
            tLIST = lList.Items(i1ctr).ToString
            If tLIST = tValue Then
                Return tLIST
            Else
                If BreakSpace Then
                    If tLIST.Contains(" ") Then
                        If tLIST.Substring(0, tLIST.IndexOf(" ")) = tValue Then
                            Return tLIST
                        End If
                    End If
                End If
            End If
        Next
        Return ""
    End Function

    Public Function Listbox_Select(ByVal lList As ListBox, ByVal tValue As String, ByVal BreakSpace As Boolean) As Integer
        Dim tLIST As String = ""
        For i1ctr = 0 To lList.Items.Count - 1
            tLIST = lList.Items(i1ctr).ToString
            If tLIST = tValue Then
                lList.SelectedIndex = i1ctr
                Return i1ctr
            Else
                If BreakSpace Then
                    If tLIST.Contains(" ") Then
                        If tLIST.Substring(0, tLIST.IndexOf(" ")) = tValue Then
                            lList.SelectedIndex = i1ctr
                            Return i1ctr
                        End If
                    End If
                End If
            End If
        Next
        Return -1
    End Function

    Public Function WhereAnd(ByVal tStrIN As String, ByVal tAdd As String) As String
        If tStrIN.Contains("WHERE") Then Return " AND " + tAdd
        Return " WHERE " + tAdd
    End Function

    Public Function WhereOr(ByVal tStrIN As String, ByVal tAdd As String) As String
        If tStrIN.Contains("WHERE") Then Return " OR " + tAdd
        Return " WHERE (" + tAdd
    End Function

    Public Function WhereOrClosing(ByVal tStrIN As String) As String
        If tStrIN.Contains("WHERE") And tStrIN.Contains("(") Then Return ")"
        Return ""
    End Function

    Public Sub DB_COMMAND(ByVal tDB As String, ByVal SQLConnectionString As String, ByVal SQLuser As String, ByVal tCOMMAND As String)
        Try
            Select Case tDB
                Case "DB2"
#If DB2 = 1 Then

                    Dim dbConnection As IBM.Data.DB2.iSeries.iDB2Connection = New IBM.Data.DB2.iSeries.iDB2Connection(SQLConnectionString + SQLuser)
                    Dim dbCommand As IBM.Data.DB2.iSeries.iDB2Command = New IBM.Data.DB2.iSeries.iDB2Command()
                    dbCommand.CommandText = tCOMMAND
                    dbCommand.CommandType = CommandType.Text
                    'dbCommand.Parameters.AddWithValue("@command", SqlDbType.Text).Value = tCOMMAND
                    dbCommand.Connection = dbConnection
                    dbCommand.CommandTimeout = 0
                    dbCommand.Connection.Open()
                    dbCommand.ExecuteNonQuery()
                    dbCommand.Dispose()
                    dbCommand = Nothing
                    dbConnection.Close()
                    dbConnection.Dispose()
                    dbConnection = Nothing
#Else
                    eLetters.MsgStatus("DB2 not defined", True)
#End If
                Case "MSSQL"
                    Dim dbConnection As New SqlConnection(SQLConnectionString + SQLuser)        ' The SqlConnection class allows you to communicate with SQL Server.
                    Dim dbCommand As New SqlCommand(tCOMMAND, dbConnection)                ' A SqlCommand object is used to execute the SQL commands.
                    dbCommand.CommandTimeout = 0
                    dbCommand.CommandText = tCOMMAND
                    dbCommand.Connection = dbConnection
                    dbCommand.Connection.Open()
                    dbCommand.ExecuteNonQuery()
                    dbCommand.Dispose()
                    dbCommand = Nothing
                    dbConnection.Close()
                    dbConnection.Dispose()
                    dbConnection = Nothing
                    'Case "MYSQL"
                    '    Dim dbConnection As New OdbcConnection(SQLConnectionString + SQLuser)    'The SqlConnection class allows you to communicate with SQL Server.
                    '    Dim dbCommand As New OdbcCommand(tCOMMAND, dbConnection)            'A SqlCommand object is used to execute the SQL commands.
                    '    dbCommand.CommandText = tCOMMAND
                    '    dbCommand.Connection = dbConnection
                    '    dbCommand.Connection.Open()
                    '    dbCommand.ExecuteNonQuery()
                    '    dbCommand.Dispose()
                    '    dbCommand = Nothing
                    '    dbConnection.Close()
                    '    dbConnection.Dispose()
                    '    dbConnection = Nothing
            End Select
        Catch ex As Exception
            Msg_Error("DB_COMMAND", tCOMMAND + vbCr + ex.ToString)
        End Try
    End Sub

    Public Function LISTBOX_init_SQLselect(ByVal lLIST As ListBox, ByVal gTempGrid As DataGridView, ByVal tDB As String, ByVal tSQLConnectionString As String, ByVal tSQLuser As String, ByVal tMODULE As String, ByVal tSELECTstring As String, ByVal bClear As Boolean) As Boolean
        Try
            If bClear Then lLIST.Items.Clear()
            If rkutils.SQL_READ_DATAGRID(gTempGrid, tDB, "*", tSQLConnectionString, tSQLuser, tSELECTstring) Then
                If gTempGrid.Rows.Count > 0 Then
                    Dim iCTR As Integer
                    Dim tSTR As String = "", tSTAT As String = "", tVIEW As String = ""
                    For iCTR = 0 To gTempGrid.Rows.Count - 1
                        Select Case tMODULE
                            Case Else
                                tSTR = rkutils.DataGridView_ValueByColumnName(gTempGrid, tMODULE, iCTR)
                                If tSTR.Trim.Length > 0 Then lLIST.Items.Add(tSTR)
                        End Select
                    Next
                    Return True
                End If
            End If
        Catch ex As Exception
            Msg_Error("LISTBOX_init_SQLselect", ex.ToString)
        End Try
        Return False
    End Function

    Function GridView_LOCXS(ByVal gGRID As DataGridView, ByVal iMaxReturn As Integer, ByVal tLOCXfield As String, ByVal tOrAnd As String, ByVal tFormatString As String) As String
        Try
            Dim tLINE As String = ""
            Dim tLOCX As String = ""
            Dim iMaxRows As Integer = 0
            If iMaxReturn < gGRID.RowCount - 1 Then
                iMaxRows = iMaxReturn
            Else
                iMaxRows = gGRID.RowCount - 1
            End If
            For i1 = 0 To iMaxRows
                tLOCX = rkutils.DataGridView_ValueByColumnName(gGRID, tLOCXfield, i1).Trim
                If tLOCX.Length > 0 Then
                    If i1 > 0 Then
                        tLINE += " " + tOrAnd + " "
                    End If
                    tLINE += tFormatString + "='" + tLOCX + "'"
                End If
            Next
            Return tLINE
        Catch ex As Exception
            Msg_Error("GridView_LOCXS", ex.ToString)
        End Try
        Return ""
    End Function

    Function EMAILIT(ByVal tSQLConnection As String, ByVal tSQLuser As String, ByVal tEmailFrom As String, ByVal tEmailFromName As String, ByVal tEmailTo As String, ByVal tEmailToName As String, ByVal tModule As String, ByVal tSubject As String, ByVal tMessage As String, ByVal tHTML As String, ByVal tATTACH As String) As Boolean
        Try
            Dim SQLstring As String = ""
            Select Case aoLetters.sSITE
                Case "iTeleCollect"
                    SQLstring = "INSERT INTO iTeleCollect.dbo.commands "
                Case Else
                    SQLstring = "INSERT INTO RevMD.dbo.commands "
            End Select
            SQLstring += " (COMMAND, TPARAMETERS, EMAILFROM, EMAILFROMNAME, EMAILTO, EMAILTONAME, EMAILSUBJECT, EMAILMESSAGE"
            SQLstring += ") VALUES("
            SQLstring += "'EMAIL'"
            SQLstring += ", ''"
            SQLstring += ", '" + tEmailFrom + "'"
            SQLstring += ", '" + tEmailFromName + "'"
            SQLstring += ", '" + tEmailTo + "'"
            SQLstring += ", '" + tEmailToName + "'"
            SQLstring += ", '" + tSubject + "'"
            If tHTML.Length > 0 Then
                SQLstring += ", '" + tHTML + "'"
            Else
                SQLstring += ", '" + tMessage + "'"
            End If
            SQLstring += ")"
            rkutils.DB_COMMAND("MSSQL", tSQLConnection, tSQLuser, SQLstring)
            Return True
        Catch ex As Exception
            Msg_Error("EMAILIT", ex.ToString)
        End Try
        Return False
    End Function

    Public Function STR_NORMALIZE(ByVal tLineIn As String) As String
        Try
            Dim tLineOut As String = tLineIn
            tLineOut = tLineOut.Replace("'", "")
            tLineOut = tLineOut.Replace(vbCr, "")
            tLineOut = tLineOut.Replace(vbLf, "")
            tLineOut = tLineOut.Replace(vbCrLf, "")
            Return tLineOut
        Catch ex As Exception
            Msg_Error("STR_NORMALIZE", ex.ToString)
            Return ""
        End Try
    End Function

    Public Function STR_NoSpecialCharacters(ByVal tLineIn As String, ByVal bOnlyNumbersAndLetters As Boolean) As String
        Try
            Dim tLineOut As String = ""
            For i1 = 0 To tLineIn.Length - 1
                If bOnlyNumbersAndLetters Then
                    Select Case Asc(tLineIn.Substring(i1, 1))
                        Case 32 'Space
                            tLineOut += tLineIn.Substring(i1, 1)
                        Case 48 To 57  '0 - 9
                            tLineOut += tLineIn.Substring(i1, 1)
                        Case 65 To 90  'A - Z
                            tLineOut += tLineIn.Substring(i1, 1)
                        Case 97 To 122  'a - z
                            tLineOut += tLineIn.Substring(i1, 1)
                        Case Else
                            tLineOut += " "
                    End Select
                Else
                    Select Case Asc(tLineIn.Substring(i1, 1))
                        Case 32 'Space
                            tLineOut += tLineIn.Substring(i1, 1)
                        Case 33 '!
                            tLineOut += tLineIn.Substring(i1, 1)
                            'Case 34 '"
                            '    tLineOut += tLineIn.Substring(i1, 1)
                        Case 35 '#
                            tLineOut += tLineIn.Substring(i1, 1)
                        Case 36 '$
                            tLineOut += tLineIn.Substring(i1, 1)
                        Case 37 '%
                            tLineOut += tLineIn.Substring(i1, 1)
                            'Case 38 '&
                            'tLineOut += tLineIn.Substring(i1, 1)
                            'Case 39 ''
                            '    tLineOut += tLineIn.Substring(i1, 1)
                        Case 40 '(
                            tLineOut += tLineIn.Substring(i1, 1)
                        Case 41 ')
                            tLineOut += tLineIn.Substring(i1, 1)
                            'Case 42 '*
                            '    tLineOut += tLineIn.Substring(i1, 1)
                        Case 43 '+
                            tLineOut += tLineIn.Substring(i1, 1)
                        Case 44 ',
                            tLineOut += tLineIn.Substring(i1, 1)
                        Case 45 '-
                            tLineOut += tLineIn.Substring(i1, 1)
                        Case 46 '.
                            tLineOut += tLineIn.Substring(i1, 1)
                            'Case 47 '/
                            '    tLineOut += tLineIn.Substring(i1, 1)
                        Case 48 To 57  '0 - 9
                            tLineOut += tLineIn.Substring(i1, 1)
                        Case 58 ':
                            tLineOut += tLineIn.Substring(i1, 1)
                        Case 59 ';
                            tLineOut += tLineIn.Substring(i1, 1)
                        Case 60 '<
                            tLineOut += tLineIn.Substring(i1, 1)
                        Case 61 '=
                            tLineOut += tLineIn.Substring(i1, 1)
                        Case 62 '>
                            tLineOut += tLineIn.Substring(i1, 1)
                        Case 63 '?
                            tLineOut += tLineIn.Substring(i1, 1)
                        Case 64 '@
                            tLineOut += tLineIn.Substring(i1, 1)
                        Case 65 To 90  'A - Z
                            tLineOut += tLineIn.Substring(i1, 1)
                        Case 97 To 122  'a - z
                            tLineOut += tLineIn.Substring(i1, 1)
                        Case Else
                            tLineOut += " "
                    End Select
                End If
            Next
            Return tLineOut
        Catch ex As Exception
            Msg_Error("STR_NoSpecialCharacters", ex.ToString)
            Return ""
        End Try
    End Function

    Function ReadField(ByVal gGrid As DataGridView, ByVal tColumnName As String, ByVal iRow As Integer) As String
        Try
            Return rkutils.STR_convert_AMP(rkutils.STR_NORMALIZE(rkutils.DataGridView_ValueByColumnName(gGrid, tColumnName, iRow).Trim))
        Catch ex As Exception
            Msg_Error("ReadField", ex.ToString)
            Return ""
        End Try
    End Function

    Function ReadFieldNoSpecialCharacters(ByVal gGrid As DataGridView, ByVal tColumnName As String, ByVal iRow As Integer, ByVal bOnlyNumbersAndLetters As Boolean) As String
        Try
            Return rkutils.STR_NoSpecialCharacters(rkutils.STR_convert_AMP(rkutils.STR_NORMALIZE(rkutils.DataGridView_ValueByColumnName(gGrid, tColumnName, iRow).Trim)), bOnlyNumbersAndLetters).Trim
        Catch ex As Exception
            Msg_Error("ReadFieldNoSpecialCharacters", ex.ToString)
            Return ""
        End Try
    End Function

    Public Function iTCS_NOTES_ADD(ByVal SQLConnectionString As String, ByVal SQLuser As String, ByVal sSysAccount As String, ByVal sClient As String, ByVal sNoteType As String, ByVal sNoteDate As String, ByVal sNoteBy As String, ByVal sNoteComment As String) As Boolean
        Try
            Dim tSQL As String = ""
            tSQL = "INSERT INTO iTeleCollect.dbo.Notes"
            tSQL += " (SysAccount,Client,NoteType,NoteDate,NoteBy,NoteComment)"
            tSQL += " values('" + sSysAccount + "'"
            tSQL += ",'" + rkutils.STR_TRIM(sClient, 4) + "'"
            tSQL += ",'" + rkutils.STR_TRIM(sNoteType, 1) + "'"
            If IsDate(rkutils.STR_format(sNoteDate, "mm/dd/ccyy HH:MM:SS")) Then
                tSQL += ",'" + rkutils.STR_format(sNoteDate, "mm/dd/ccyy HH:MM:SS") + "'"
            Else
                tSQL += ",'" + rkutils.STR_format("TODAY", "mm/dd/ccyy HH:MM:SS") + "'"
            End If
            tSQL += ",'" + rkutils.STR_TRIM(sNoteBy, 20) + "'"
            tSQL += ",'" + sNoteComment + "'"
            tSQL += ")"
            DB_COMMAND("MSSQL", SQLConnectionString, SQLuser, tSQL)
            Return True
        Catch ex As Exception
            Msg_Error("NOTES_ADD", ex.ToString)
        End Try
        Return False
    End Function

    Public Function DataGridView_Contains2Cols(ByVal gGrid As DataGridView, ByVal tColName As String, ByVal tValue As String, ByVal tColName2 As String, ByVal tValue2 As String) As Integer
        Try
            Dim rc1 As Integer = DataGridView_ColumnByName(gGrid, tColName)
            Dim rc2 As Integer = DataGridView_ColumnByName(gGrid, tColName2)

            If rc1 >= 0 And rc2 >= 0 Then
                For i1 = 0 To gGrid.RowCount - 1
                    If gGrid.Item(rc1, i1).Value.ToString.Trim = tValue And gGrid.Item(rc2, i1).Value.ToString.Trim = tValue2 Then Return i1
                Next
            End If
            Return -1
        Catch ex As Exception
            'MSG_warning(ex.ToString)
        End Try
        Return -1
    End Function

    Public Function DataTable_Contains2Cols(ByVal dT As DataTable, ByVal tColName As String, ByVal tValue As String, ByVal tColName2 As String, ByVal tValue2 As String) As Integer
        Try
            Dim rc1 As Integer = DataTable_ColumnByName(dT, tColName)
            Dim rc2 As Integer = DataTable_ColumnByName(dT, tColName2)

            If rc1 >= 0 And rc2 >= 0 Then
                For i1 = 0 To dT.Rows.Count - 1
                    If dT.Rows(rc1)(i1).Value.ToString.Trim = tValue And dT.Rows(rc2)(i1).Value.ToString.Trim = tValue2 Then Return i1
                Next
            End If
            Return -1
        Catch ex As Exception
            'MSG_warning(ex.ToString)
        End Try
        Return -1
    End Function

    Public Function Encrypt(clearText As String, EncryptionKey As String) As String
        Try
            Dim clearBytes As Byte() = Encoding.Unicode.GetBytes(clearText)
            Using encryptor As Aes = Aes.Create()
                Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D,
                 &H65, &H64, &H76, &H65, &H64, &H65,
                 &H76})
                encryptor.Key = pdb.GetBytes(32)
                encryptor.IV = pdb.GetBytes(16)
                Using ms As New MemoryStream()
                    Using cs As New CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write)
                        cs.Write(clearBytes, 0, clearBytes.Length)
                        cs.Close()
                    End Using
                    clearText = Convert.ToBase64String(ms.ToArray())
                End Using
            End Using
        Catch ex As Exception
            Msg_Error("Encrypt", ex.ToString)
        End Try
        Return clearText
    End Function

    Public Function Encrypt(sVersion As String, clearText As String, EncryptionKey As String) As String
        Try
            Select Case sVersion
                Case "VB6"
                    '*****************************
                    '* 2015-10-19 RFK: vb6 version
                    Dim iLen As Integer, iX As Integer, iX2 As Integer
                    iLen = Len(EncryptionKey)
                    iX2 = 1
                    Dim sSTR As String
                    sSTR = ""
                    For iX = 1 To Len(clearText)
                        sSTR = sSTR + Trim(Str(Asc(Mid(clearText, iX, 1)) + Asc(Mid(EncryptionKey, iX2, 1)))).PadLeft(3, "0")
                        iX2 = iX2 + 1
                        If iX2 > iLen Then iX2 = 1
                    Next
                    Return sSTR
                Case Else
                    '*****************************
                    '* RFK: VB.NET VERSION (much more secure)
                    Dim clearBytes As Byte() = Encoding.Unicode.GetBytes(clearText)
                    Using encryptor As Aes = Aes.Create()
                        Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D, &H65, &H64, &H76, &H65, &H64, &H65, &H76})
                        encryptor.Key = pdb.GetBytes(32)
                        encryptor.IV = pdb.GetBytes(16)
                        Using ms As New MemoryStream()
                            Using cs As New CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write)
                                cs.Write(clearBytes, 0, clearBytes.Length)
                                cs.Close()
                            End Using
                            clearText = Convert.ToBase64String(ms.ToArray())
                        End Using
                    End Using
            End Select
        Catch ex As Exception
            Msg_Error("Encrypt", ex.ToString)
        End Try
        Return clearText
    End Function

    Public Function Decrypt(sVersion As String, cipherText As String, EncryptionKey As String) As String
        Try
            Select Case sVersion
                Case "VB6"
                    '*****************************
                    '* 2015-10-19 RFK: vb6 version
                    Dim iLen As Integer, iX As Integer, iX2 As Integer
                    iLen = Len(EncryptionKey)
                    iX2 = 1
                    Dim sSTR As String
                    sSTR = ""
                    For iX = 1 To Len(cipherText) Step 3
                        sSTR = sSTR + Trim(Chr(Val(Mid(cipherText, iX, 3)) - Asc(Mid(EncryptionKey, iX2, 1))))
                        iX2 = iX2 + 1
                        If iX2 > iLen Then iX2 = 1
                    Next
                    Return sSTR
                Case Else
                    '*****************************
                    '* RFK: VB.NET VERSION (much more secure)
                    Dim cipherBytes As Byte() = Convert.FromBase64String(cipherText)
                    Using encryptor As Aes = Aes.Create()
                        Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D,
                         &H65, &H64, &H76, &H65, &H64, &H65,
                         &H76})
                        encryptor.Key = pdb.GetBytes(32)
                        encryptor.IV = pdb.GetBytes(16)
                        Using ms As New MemoryStream()
                            Using cs As New CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write)
                                cs.Write(cipherBytes, 0, cipherBytes.Length)
                                cs.Close()
                            End Using
                            cipherText = Encoding.Unicode.GetString(ms.ToArray())
                        End Using
                    End Using
            End Select
        Catch ex As Exception
            Msg_Error("Decrypt", ex.ToString)
        End Try
        Return cipherText
    End Function

    Public Function ReadFieldSelectStringDTable(DT As DataTable, tColumnName As String, iRow As Integer, FieldNameOrValue As Boolean, tFieldName As String, iLen As Integer) As String
        Try
            Dim sTemp As String = DataTable_ValueByColumnName(DT, tColumnName, iRow)
            If sTemp.Trim.Length > 0 Then
                If FieldNameOrValue Then
                    Return "," + tFieldName
                Else
                    Return ",'" + STR_TRIM(sTemp, iLen) + "'"
                End If
            End If
            Return ""
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Function ReadFieldAlternateSelectStringDTable(DT As DataTable, tColumnName As String, tColumnNameAlternate As String, iRow As Integer, FieldNameOrValue As Boolean, tFieldName As String, iLen As Integer) As String
        Try
            Dim sTemp As String = DataTable_ValueByColumnName(DT, tColumnName, iRow)
            If sTemp.Trim.Length > 0 Then
                If FieldNameOrValue Then
                    Return "," + tFieldName
                Else
                    Return ",'" + STR_TRIM(STR_NORMALIZE(sTemp), iLen) + "'"
                End If
            Else
                sTemp = DataTable_ValueByColumnName(DT, tColumnNameAlternate, iRow)
                If sTemp.Trim.Length > 0 Then
                    If FieldNameOrValue Then
                        Return "," + tFieldName
                    Else
                        Return ",'" + STR_TRIM(STR_NORMALIZE(sTemp), iLen) + "'"
                    End If
                End If
            End If
            Return ""
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Function DataTable_ColumnByName(ByVal DT As DataTable, ByVal tColName As String) As Integer
        Try
            '******************************************************************
            ' 2015-08-04 RFK: all the rows
            Dim i2 As Integer = 0
            For Each col As DataColumn In DT.Columns
                If col.ColumnName.ToUpper = tColName.ToUpper Then Return i2
                i2 += 1
            Next
            'For Each row As DataRow In DT.Rows
            '    builder = New System.Text.StringBuilder
            '    '**************************************************************
            '    sw.WriteLine(builder.ToString())
            'Next
        Catch ex As Exception
            'MSG_warning(ex.ToString)
        End Try
        Return -1
    End Function

    Public Function DataTable_ValueByColumnName(ByVal DT As DataTable, ByVal tColName As String, ByVal iRow As Integer) As String
        Try
            Dim iRC As Integer = 0
            iRC = DataTable_ColumnByName(DT, tColName)
            If iRC >= 0 Then
                If DT.Rows(iRow).Item(iRC) IsNot Nothing Then
                    If DT.Rows(iRow).Item(iRC).ToString.Trim.Length > 0 Then
                        Return DT.Rows(iRow).Item(iRC).ToString.Trim
                    End If
                End If
            End If
        Catch ex As Exception
            'MSG_warning(ex.ToString)
        End Try
        Return ""
    End Function

    Public Function ReadFieldDataTable(ByVal DT As DataTable, ByVal tColumnName As String, ByVal iRow As Integer) As String
        Return rkutils.STR_convert_AMP(rkutils.STR_NORMALIZE(DataTable_ValueByColumnName(DT, tColumnName, iRow).Trim))
    End Function

    Public Function STR_BREAK_PIECES(ByVal tLineIn As String, ByVal WhichOne As Integer, ByVal tBreakCharacter As String) As String
        Dim iC As Integer, iLast As Integer = -1, iBreakCTR As Integer = 1
        Dim inQUOTE As Boolean = False, swOK As Boolean = False
        Dim tLineOut As String = ""
        Try
            For iC = 0 To tLineIn.Length - 1
                swOK = False
                If tBreakCharacter = "," Then
                    If tLineIn.Substring(iC, 1) = Chr(34) Then
                        If inQUOTE = False Then
                            inQUOTE = True
                        Else
                            inQUOTE = False
                        End If
                    End If
                    If inQUOTE = False Then swOK = True
                Else
                    swOK = True
                End If
                '******************
                If swOK Then
                    If tLineIn.Substring(iC, 1) = tBreakCharacter Then
                        If iBreakCTR = WhichOne Then
                            Dim t2 As String = tLineIn.Substring(iLast + 1, iC - iLast - 1).Trim
                            If t2.Length = 0 Then t2 = "_" 'Chr(160) 'So Not Blank
                            Return t2
                        Else
                            iLast = iC
                        End If
                        iBreakCTR += 1
                    End If
                End If
            Next
            'iScraper.MessageStatus(WhichOne.ToString + vbCr + iBreakCTR.ToString + vbCr + iLast.ToString + vbCr + iC.ToString)
            If iBreakCTR = WhichOne Then Return tLineIn.Substring(iLast + 1).Trim 'The Last One
            Return ""
        Catch ex As Exception
            Msg_Error("STR_BREAK_PIECES", tLineOut + vbCr + vbCr + ex.ToString)
            Return ""
        End Try
    End Function

    Public Sub TRACKS_update(ByVal tDBtype As String, ByVal SQLConnectionString As String, ByVal SQLuser As String, ByVal tCLIENT As String, ByVal tLOCX As String, ByVal tUNIQUE As String, ByVal tTYPE As String, ByVal tCOMMENT As String)
        Try
            Dim msSQLCommandString As String = ""
            Select Case tDBtype
                Case "MSSQL"
                    If tUNIQUE.Length > 0 Then
                        '
                    Else
                        msSQLCommandString += "INSERT INTO RevMD.dbo.tracks ("
                        msSQLCommandString += "track_date, track_by"
                        msSQLCommandString += ", comment"
                        msSQLCommandString += ", type"
                        msSQLCommandString += ", client"
                        msSQLCommandString += ", locx"
                        msSQLCommandString += ") values("
                        msSQLCommandString += "@now, @by"
                        msSQLCommandString += ", @comment"
                        msSQLCommandString += ", @type"
                        msSQLCommandString += ", @client"
                        msSQLCommandString += ", @locx"
                        msSQLCommandString += ")"
                    End If
                    Dim dbConnection As New SqlConnection(SQLConnectionString + SQLuser)        'The SqlConnection class allows you to communicate with SQL Server.
                    Dim dbCommand As New SqlCommand(msSQLCommandString, dbConnection)           'A SqlCommand object is used to execute the SQL commands.
                    dbCommand.Parameters.AddWithValue("@tunique", SqlDbType.VarChar).Value = System.Guid.NewGuid.ToString()
                    dbCommand.Parameters.AddWithValue("@comment", SqlDbType.Text).Value = tCOMMENT
                    dbCommand.Parameters.AddWithValue("@type", SqlDbType.VarChar).Value = Left(tTYPE, 1)
                    dbCommand.Parameters.AddWithValue("@client", SqlDbType.VarChar).Value = Left(tCLIENT, 20)
                    dbCommand.Parameters.AddWithValue("@locx", SqlDbType.VarChar).Value = tLOCX
                    dbCommand.Parameters.AddWithValue("@now", SqlDbType.DateTime).Value = Date.Now
                    dbCommand.Parameters.AddWithValue("@by", SqlDbType.VarChar).Value = rkutils.WhoAmI()

                    dbCommand.Connection.Open()
                    dbCommand.CommandTimeout = 0
                    dbCommand.ExecuteNonQuery()
                    dbCommand.Connection.Close()
                    dbCommand.Dispose()
                    dbCommand = Nothing
                    dbConnection.Close()
                    dbConnection.Dispose()
                    dbConnection = Nothing
                Case "DB2"
                    If tUNIQUE.Length > 0 Then
                        '
                    Else
                        msSQLCommandString += "INSERT INTO ROIDATA.tracks ("
                        msSQLCommandString += "track_date, track_by"
                        msSQLCommandString += ", comment"
                        msSQLCommandString += ", type"
                        msSQLCommandString += ", client"
                        msSQLCommandString += ", locx"
                        msSQLCommandString += ") values("
                        msSQLCommandString += "@now, @by"
                        msSQLCommandString += ", @comment"
                        msSQLCommandString += ", @type"
                        msSQLCommandString += ", @client"
                        msSQLCommandString += ", @locx"
                        msSQLCommandString += ")"
                    End If
                    DB_COMMAND("DB2", SQLConnectionString, SQLuser, msSQLCommandString)
                Case Else
                    Exit Sub
            End Select
        Catch ex As Exception
            Msg_Error("TRACKS_update", ex.ToString)
        End Try
    End Sub

    Function Mod10_CheckDigit(ByVal tLineIn As String, ByVal sWeight As String, ByVal LeftToRight As Boolean, ByVal WeightSpaces As Boolean) As String
        Try
            '********************************************************************************
            '* 2013-05-29 RFK: Mod10_CheckDigit 
            '* 2013-05-29 RFK: this code is readable instead of compact or speedy
            '********************************************************************************
            Dim sScanLine As String = tLineIn
            Dim i As Integer, iTotalDigits As Integer = 0, iCheckDigit As Integer = 0
            Dim array(1000) As Integer                          'result of weight multiplication
            Dim array_10(1000) As Integer                       '10s digits
            Dim array_1(1000) As Integer                        '1s digits
            Dim array_result(1000) As Integer                   '1s + 10s digits
            Dim swDebug As Boolean = False, swDebug2 As Boolean = False
            Dim iWeight As Integer = Val(STR_LEFT(sWeight, 1))  'Weight starts from the left
            If LeftToRight = False Then
                iWeight = Val(STR_RIGHT(sWeight, 1))            'Weight starts from the right
            End If
            Dim swCalcIt As Boolean = False
            '********************************************************************************
            If swDebug Then aoLetters.MsgStatus(sScanLine, True)
            If LeftToRight Then
                For i = 1 To Len(sScanLine)
                    '********************************************************************************
                    '* 2013-10-15 RFK: Only numeric digits
                    swCalcIt = False
                    If Mid(sScanLine, i, 1) = " " And WeightSpaces Then
                        swCalcIt = True
                    Else
                        If Mid(sScanLine, i, 1) >= "0" And Mid(sScanLine, i, 1) <= "9" Then
                            swCalcIt = True
                        Else
                            If Mid(sScanLine, i, 1) >= "A" And Mid(sScanLine, i, 1) <= "Z" Then
                                swCalcIt = True
                            End If
                        End If
                    End If
                    If swCalcIt Then
                        '********************************************************************************
                        '* Calculate the result of the digit multiplied by weight
                        '* 2013-11-27 RFK: LETTERS A=1 Z=26
                        If Mid(sScanLine, i, 1) >= "A" And Mid(sScanLine, i, 1) <= "Z" Then
                            If swDebug2 Then aoLetters.MsgStatus(Mid(sScanLine, i, 1) + "_" + (Asc(Mid(sScanLine, i, 1)) - 64).ToString + "_" + array(i).ToString, True)
                            '**************************************************************
                            '* 2014-03-13 RFK: Changed to different Letter Weights
                            '* 2016-04-19 RFK: added additional weight 371
                            Select Case sWeight
                                Case "371"
                                    Select Case Mid(sScanLine, i, 1)
                                        Case "A", "J", "S"
                                            array(i) = 8 * iWeight
                                        Case "B", "K", "T"
                                            array(i) = 9 * iWeight
                                        Case "C", "L", "U"
                                            array(i) = 1 * iWeight
                                        Case "D", "M", "V"
                                            array(i) = 2 * iWeight
                                        Case "E", "N", "W"
                                            array(i) = 3 * iWeight
                                        Case "F", "O", "X"
                                            array(i) = 4 * iWeight
                                        Case "G", "P", "Y"
                                            array(i) = 5 * iWeight
                                        Case "H", "Q", "Z"
                                            array(i) = 6 * iWeight
                                        Case "I", "R"
                                            array(i) = 7 * iWeight
                                    End Select
                                Case "3579"
                                    Select Case Mid(sScanLine, i, 1)
                                        Case "A", "J", "S"
                                            array(i) = 1 * iWeight
                                        Case "B", "K", "T"
                                            array(i) = 2 * iWeight
                                        Case "C", "L", "U"
                                            array(i) = 3 * iWeight
                                        Case "D", "M", "V"
                                            array(i) = 4 * iWeight
                                        Case "E", "N", "W"
                                            array(i) = 5 * iWeight
                                        Case "F", "O", "X"
                                            array(i) = 6 * iWeight
                                        Case "G", "P", "Y"
                                            array(i) = 7 * iWeight
                                        Case "H", "Q", "Z"
                                            array(i) = 8 * iWeight
                                        Case "I", "R"
                                            array(i) = 9 * iWeight
                                    End Select
                                Case Else
                                    'A=26
                                    array(i) = CInt((Asc(Mid(sScanLine, i, 1))) - 64) * iWeight
                            End Select
                        Else
                            array(i) = CInt(Mid(sScanLine, i, 1)) * iWeight
                        End If
                        '********************************************************************************
                        '* Calculate the result 
                        If array(i) >= 10 Then
                            array_10(i) = Val(STR_LEFT(array(i).ToString.Trim, 1))  'the 10s digits
                            array_1(i) = Val(STR_RIGHT(array(i).ToString.Trim, 1))  'the 1s digits
                        Else
                            array_10(i) = 0                             'Calculate the 10s digits
                            array_1(i) = array(i)                       'Calculate the 1s digits
                        End If
                        array_result(i) = array_10(i) + array_1(i)      'Add the 10s and 1s digits
                        iTotalDigits += array_result(i)                 'Add up the results
                        '********************************************************************************
                        If swDebug Then aoLetters.MsgStatus("Character:" + Mid(sScanLine, i, 1) + " * (Weight)" + Trim(iWeight.ToString) + " =" + Trim(Str(array(i))) + " 10s:" + Trim(Str(array_10(i))) + " + 1s:" + Trim(Str(array_1(i))) + " =" + Trim(Str(array_result(i))) + " Total:" + iTotalDigits.ToString + "]", True)
                        If swDebug2 Then aoLetters.MsgStatus("[" + Trim(Mid(sScanLine, i, 1)) + "[" + Trim(iWeight.ToString) + "][" + Trim(Str(array(i))) + "][" + Trim(Str(array_10(i))) + "][" + Trim(Str(array_1(i))) + "][" + Trim(Str(array_result(i))) + "][" + iTotalDigits.ToString + "]", True)
                        '********************************************************************************
                        '* 2013-05-29 RFK: Alternate the weight
                        '* 2013-10-15 RFK: added additional weight 7532
                        '* 2016-04-19 RFK: added additional weight 371
                        Select Case sWeight
                            Case "371"
                                Select Case iWeight
                                    Case 3
                                        iWeight = 7
                                    Case 7
                                        iWeight = 1
                                    Case 1
                                        iWeight = 3
                                    Case Else
                                        iWeight = 3
                                End Select
                            Case "12", "21"
                                If iWeight = 1 Then
                                    iWeight = 2
                                Else
                                    iWeight = 1
                                End If
                            Case "3579"
                                Select Case iWeight
                                    Case 3
                                        iWeight = 5
                                    Case 5
                                        iWeight = 7
                                    Case 7
                                        iWeight = 9
                                    Case 9
                                        iWeight = 3
                                    Case Else
                                        iWeight = 3
                                End Select
                            Case "7532"
                                Select Case iWeight
                                    Case 7
                                        iWeight = 5
                                    Case 5
                                        iWeight = 3
                                    Case 3
                                        iWeight = 2
                                    Case 2
                                        iWeight = 7
                                    Case Else
                                        iWeight = 7
                                End Select
                        End Select
                    Else
                        'Skip the Space
                    End If
                    '********************************************************************************
                Next i
            Else
                '********************************************************************************
                For i = Len(sScanLine) To 1 Step -1
                    '********************************************************************************
                    '* 2013-10-15 RFK: Only numeric digits
                    If (Mid(sScanLine, i, 1) >= "0" And Mid(sScanLine, i, 1) <= "9") Or (Mid(sScanLine, i, 1) >= "A" And Mid(sScanLine, i, 1) <= "Z") Then
                        '********************************************************************************
                        '* Calculate the result of the digit multiplied by weight
                        '* 2013-11-27 RFK: LETTERS A=1 Z=26
                        If Mid(sScanLine, i, 1) >= "A" And Mid(sScanLine, i, 1) <= "Z" Then
                            '**************************************************************
                            '* 2014-03-13 RFK: Changed to different Letter Weights
                            '* 2016-04-19 RFK: added additional weight 371
                            Select Case sWeight
                                Case "371"
                                    Select Case Mid(sScanLine, i, 1)
                                        Case "A", "J", "S"
                                            array(i) = 8 * iWeight
                                        Case "B", "K", "T"
                                            array(i) = 9 * iWeight
                                        Case "C", "L", "U"
                                            array(i) = 1 * iWeight
                                        Case "D", "M", "V"
                                            array(i) = 2 * iWeight
                                        Case "E", "N", "W"
                                            array(i) = 3 * iWeight
                                        Case "F", "O", "X"
                                            array(i) = 4 * iWeight
                                        Case "G", "P", "Y"
                                            array(i) = 5 * iWeight
                                        Case "H", "Q", "Z"
                                            array(i) = 6 * iWeight
                                        Case "I", "R"
                                            array(i) = 7 * iWeight
                                    End Select
                                Case "3579"
                                    Select Case Mid(sScanLine, i, 1)
                                        Case "A", "J", "S"
                                            array(i) = 1 * iWeight
                                        Case "B", "K", "T"
                                            array(i) = 2 * iWeight
                                        Case "C", "L", "U"
                                            array(i) = 3 * iWeight
                                        Case "D", "M", "V"
                                            array(i) = 4 * iWeight
                                        Case "E", "N", "W"
                                            array(i) = 5 * iWeight
                                        Case "F", "O", "X"
                                            array(i) = 6 * iWeight
                                        Case "G", "P", "Y"
                                            array(i) = 7 * iWeight
                                        Case "H", "Q", "Z"
                                            array(i) = 8 * iWeight
                                        Case "I", "R"
                                            array(i) = 9 * iWeight
                                    End Select
                                Case Else
                                    'A=26
                                    array(i) = CInt((Asc(Mid(sScanLine, i, 1))) - 39) * iWeight
                            End Select
                        Else
                            array(i) = CInt(Mid(sScanLine, i, 1)) * iWeight
                        End If
                        If array(i) >= 10 Then
                            array_10(i) = 1                             'Calculate the 10s digits
                            array_1(i) = array(i) - 10                  'Calculate the 1s digits
                        Else
                            array_10(i) = 0                             'Calculate the 10s digits
                            array_1(i) = array(i)                       'Calculate the 1s digits
                        End If
                        array_result(i) = array_10(i) + array_1(i)      'Add the 10s and 1s digits
                        iTotalDigits += array_result(i)                 'Add up the results
                        '********************************************************************************
                        If swDebug2 Then aoLetters.MsgStatus("[" + Trim(Mid(sScanLine, i, 1)) + "[" + Trim(iWeight.ToString) + "][" + Trim(Str(array(i))) + "][" + Trim(Str(array_10(i))) + "][" + Trim(Str(array_1(i))) + "][" + Trim(Str(array_result(i))) + "][" + iTotalDigits.ToString + "]", True)
                        '********************************************************************************
                        '* 2013-05-29 RFK: Alternate the weight
                        '* 2013-10-15 RFK: added additional weight 7532
                        '* 2016-04-19 RFK: added additional weight 371
                        Select Case sWeight
                            Case "12", "21"
                                If iWeight = 1 Then
                                    iWeight = 2
                                Else
                                    iWeight = 1
                                End If
                            Case "3579"
                                Select Case iWeight
                                    Case 3
                                        iWeight = 5
                                    Case 5
                                        iWeight = 7
                                    Case 7
                                        iWeight = 9
                                    Case 9
                                        iWeight = 3
                                    Case Else
                                        iWeight = 3
                                End Select
                            Case "371"
                                Select Case iWeight
                                    Case 3
                                        iWeight = 7
                                    Case 7
                                        iWeight = 1
                                    Case 1
                                        iWeight = 3
                                    Case Else
                                        iWeight = 3
                                End Select
                            Case "7532"
                                Select Case iWeight
                                    Case 7
                                        iWeight = 5
                                    Case 5
                                        iWeight = 3
                                    Case 3
                                        iWeight = 2
                                    Case 2
                                        iWeight = 7
                                    Case Else
                                        iWeight = 7
                                End Select
                        End Select
                    End If
                    '**********************************************************
                Next i
            End If
            '******************************************************************
            If swDebug Then aoLetters.MsgStatus("Sum digits = " + iTotalDigits.ToString, True)
            If swDebug Then aoLetters.MsgStatus(iTotalDigits.ToString + " divided by 10 = " + Str(iTotalDigits / 10), True)
            If swDebug Then aoLetters.MsgStatus("Remainder = " + Str(10 * ((iTotalDigits / 10) - CInt(iTotalDigits / 10))), True)
            If swDebug Then aoLetters.MsgStatus("Subtract Remainder = " + Str(10 - (10 * ((iTotalDigits / 10) - CInt(iTotalDigits / 10)))), True)
            '******************************************************************
            '* Calculate the CheckDigit
            iCheckDigit = 10 - (iTotalDigits Mod 10)
            If swDebug Then aoLetters.MsgStatus("CheckDigit= " + iCheckDigit.ToString, True)
            If iCheckDigit >= 10 Then iCheckDigit = 0
            '******************************************************************
            Return iCheckDigit.ToString
        Catch ex As Exception
            'ex.ToString 
            'MsgBox(ex.ToString)
        End Try
        Return ""
    End Function

    Public Function DataGridView_SumColumn(ByVal gGrid As DataGridView, ByVal iC As Integer, ByVal bHasHeader As Boolean) As String
        Try
            '******************************************************************
            '* 2016-08-31 RFK:
            If gGrid.RowCount < 1 Then Return ""
            Dim iR As Integer = 0
            Dim iTotal As Integer = 0
            '******************************************************************
            For iR = 0 To gGrid.RowCount - 1
                If gGrid.Item(iC, iR).Value IsNot Nothing Then
                    If Val(gGrid.Item(iC, iR).Value.ToString) > 0 Then
                        iTotal += Val(gGrid.Item(iC, iR).Value.ToString)
                    End If
                End If
            Next
            Return iTotal.ToString.Trim
        Catch ex As Exception
            'MsgBox(ex.ToString)
            Msg_Error("DataGridView_SumColumn", ex.ToString)
        End Try
        Return -1
    End Function

    Public Function DataGridView_SumRow(ByVal gGrid As DataGridView, ByVal iR As Integer, ByVal iCfrom As Integer, ByVal iCto As Integer) As String
        Try
            '******************************************************************
            '* 2016-08-31 RFK:
            If gGrid.RowCount < 1 Then Return ""
            If iR > gGrid.RowCount Then Return ""
            If iCfrom > gGrid.ColumnCount - 1 Or iCto > gGrid.ColumnCount - 1 Then Return ""
            Dim iC As Integer = 0
            Dim iTotal As Integer = 0
            '******************************************************************
            For iC = iCfrom To iCto
                If gGrid.Item(iC, iR).Value IsNot Nothing Then
                    If Val(gGrid.Item(iC, iR).Value.ToString) > 0 Then
                        iTotal += Val(gGrid.Item(iC, iR).Value.ToString)
                        'eLetters.MsgStatus("iC=" + iC.ToString + " iR=" + iR.ToString + " =" + gGrid.Item(iC, iR).Value.ToString + " " + iTotal.ToString, True)
                    End If
                End If
            Next
            Return iTotal.ToString.Trim
        Catch ex As Exception
            Msg_Error("DataGridView_SumColumn", ex.ToString)
        End Try
        Return -1
    End Function

    Public Function DataGridView_FormatColumn(ByVal gGrid As DataGridView, ByVal iC As Integer, ByVal sFormat As String, ByVal bHeader As Boolean) As Boolean
        Try
            '******************************************************************
            '* 2016-08-31 RFK:
            If gGrid.RowCount < 1 Then Return -1
            Dim iR As Integer = 0
            Dim sTemp As String = ""
            '******************************************************************
            For iR = 1 To gGrid.RowCount - 1
                If gGrid.Item(iC, iR).Value IsNot Nothing Then
                    If Val(gGrid.Item(iC, iR).Value.ToString) > 0 Then
                        sTemp = rkutils.STR_format(gGrid.Item(iC, iR).Value.ToString, sFormat)
                        gGrid.Item(iC, iR).Value = sTemp
                    End If
                End If
            Next
            Return True
        Catch ex As Exception
            'MsgBox(ex.ToString)
            Msg_Error("DataGridView_SumColumn", ex.ToString)
        End Try
        Return -1
    End Function

    Public Function DataGridView_ToHtmlTrTd(ByVal gGrid As DataGridView, ByVal sTd1stAlign As String, ByVal sTdRestAlign As String) As String
        Try
            '******************************************************************
            '* 2016-08-30 RFK:
            Dim sHTML As String = ""
            Dim iC As Integer = 0, iR As Integer = 0
            '******************************************************************
            sHTML += "<tr>"
            For iC = 0 To gGrid.ColumnCount - 1
                If iC = 0 Then
                    If gGrid.Item(iC, iR).Style.BackColor.ToString.Contains("Empty") Then
                        If sTd1stAlign.Length > 0 Then
                            sHTML += "<td align=" + sTd1stAlign + ">" + gGrid.Columns(iC).Name + "</td>"
                        Else
                            sHTML += "<td>" + gGrid.Columns(iC).Name + "</td>"
                        End If
                    Else
                        If sTd1stAlign.Length > 0 Then
                            sHTML += "<td align=" + sTd1stAlign + " style=background-color:" + gGrid.Item(iC, iR).Style.BackColor.ToString.Replace("Color [", "").Replace("]", "") + ">" + gGrid.Columns(iC).Name + "</td>"
                        Else
                            sHTML += "<td style=background-color:" + gGrid.Item(iC, iR).Style.BackColor.ToString.Replace("Color [", "").Replace("]", "") + ">" + gGrid.Columns(iC).Name + "</td>"
                        End If
                    End If
                Else
                    If gGrid.Item(iC, iR).Style.BackColor.ToString.Contains("Empty") Then
                        If sTdRestAlign.Length > 0 Then
                            sHTML += "<td align=" + sTdRestAlign + ">" + gGrid.Columns(iC).Name + "</td>"
                        Else
                            sHTML += "<td>" + gGrid.Columns(iC).Name + "</td>"
                        End If
                    Else
                        If sTdRestAlign.Length > 0 Then
                            sHTML += "<td align=" + sTdRestAlign + " style=background-color:" + gGrid.Item(iC, iR).Style.BackColor.ToString.Replace("Color [", "").Replace("]", "") + ">" + gGrid.Columns(iC).Name + "</td>"
                        Else
                            sHTML += "<td style=background-color:" + gGrid.Item(iC, iR).Style.BackColor.ToString.Replace("Color [", "").Replace("]", "") + ">" + gGrid.Columns(iC).Name + "</td>"
                        End If
                    End If
                End If
            Next
            sHTML += "</tr>"
            '******************************************************************
            For iR = 0 To gGrid.RowCount - 1
                sHTML += "<tr>"
                For iC = 0 To gGrid.ColumnCount - 1
                    If gGrid.Item(iC, iR).Value IsNot Nothing Then
                        'eLetters.MsgStatus(gGrid.Item(iC, iR).Style.BackColor.ToString, True)
                        If iC = 0 Then
                            If gGrid.Item(iC, iR).Style.BackColor.ToString.Contains("Empty") Then
                                If sTd1stAlign.Length > 0 Then
                                    sHTML += "<td align=" + sTd1stAlign + ">" + gGrid.Item(iC, iR).Value.ToString.Trim + "</td>"
                                Else
                                    sHTML += "<td>" + gGrid.Item(iC, iR).Value.ToString.Trim + "</td>"
                                End If
                            Else
                                If sTd1stAlign.Length > 0 Then
                                    sHTML += "<td align=" + sTd1stAlign + " style=background-color:" + gGrid.Item(iC, iR).Style.BackColor.ToString.Replace("Color [", "").Replace("]", "") + ">" + gGrid.Item(iC, iR).Value.ToString.Trim + "</td>"
                                Else
                                    sHTML += "<td style=background-color:" + gGrid.Item(iC, iR).Style.BackColor.ToString.Replace("Color [", "").Replace("]", "") + ">" + gGrid.Item(iC, iR).Value.ToString.Trim + "</td>"
                                End If
                            End If
                        Else
                            If gGrid.Item(iC, iR).Style.BackColor.ToString.Contains("Empty") Then
                                If sTdRestAlign.Length > 0 Then
                                    sHTML += "<td align=" + sTdRestAlign + ">" + gGrid.Item(iC, iR).Value.ToString.Trim + "</td>"
                                Else
                                    sHTML += "<td>" + gGrid.Item(iC, iR).Value.ToString.Trim + "</td>"
                                End If
                            Else
                                If sTdRestAlign.Length > 0 Then
                                    sHTML += "<td align=" + sTdRestAlign + " style=background-color:" + gGrid.Item(iC, iR).Style.BackColor.ToString.Replace("Color [", "").Replace("]", "") + ">" + gGrid.Item(iC, iR).Value.ToString.Trim + "</td>"
                                Else
                                    sHTML += "<td style=background-color:" + gGrid.Item(iC, iR).Style.BackColor.ToString.Replace("Color [", "").Replace("]", "") + ">" + gGrid.Item(iC, iR).Value.ToString.Trim + "</td>"
                                End If
                            End If
                        End If
                    Else
                        sHTML += "<td></td>"
                    End If
                Next
                sHTML += "</tr>"
            Next
            Return sHTML
        Catch ex As Exception
            Msg_Error("DataGridView_ToHtmlTable", ex.ToString)
        End Try
        Return -1
    End Function

    Public Function DataTable_to_CSV(ByVal dT As DataTable, ByVal sFileName As String, ByVal IncludeColumnHeader As Boolean, ByVal lLabelRowUpdate As Label) As Boolean
        Try
            '******************************************************************
            '* 2015-07-30 RFK: Check for existing
            Dim sep As String = ""
            Dim builder As New System.Text.StringBuilder
            If File.Exists(sFileName) Then
                aoLetters.MsgStatus("Deleting:" + sFileName, True)
                File.Delete(sFileName)
            End If
            If File.Exists(sFileName) Then
                aoLetters.MsgStatus("Unable to create, exists:" + sFileName, True)
                Return False
            End If
            '******************************************************************
            '* 2015-07-30 RFK: Open
            aoLetters.MsgStatus("Creating:" + sFileName, True)
            Dim sw As System.IO.StreamWriter
            sw = My.Computer.FileSystem.OpenTextFileWriter(sFileName, True)
            '******************************************************************
            ' 2015-07-30 RFK: columns
            If IncludeColumnHeader Then
                sep = ""
                For Each col As DataColumn In dT.Columns
                    builder.Append(sep).Append(Trim(col.ColumnName))
                    sep = ","   'After 1st one now add the seperator
                Next
                sw.WriteLine(builder.ToString())
            End If
            '******************************************************************
            ' 2015-07-30 RFK: all the rows
            Dim iCTR As Integer = 0
            For Each row As DataRow In dT.Rows
                lLabelRowUpdate.Text = Val(lLabelRowUpdate.Text) - 1.ToString.Trim
                System.Windows.Forms.Application.DoEvents()
                '**************************************************************
                sep = ""
                builder = New System.Text.StringBuilder
                '**************************************************************
                iCTR = 0
                For Each col As DataColumn In dT.Columns
                    'eLetters.MsgStatus("-----", True)
                    'eLetters.MsgStatus(col.ColumnName, True)
                    'eLetters.MsgStatus(col.DataType.Name, True)
                    'eLetters.MsgStatus(row.Item(iCTR).ToString, True)
                    If row.Item(iCTR).ToString.Length > 0 Then
                        If row(col.ColumnName).ToString IsNot Nothing Then
                            builder.Append(sep).Append(Trim(row(col.ColumnName)).Replace(",", " "))
                        End If
                    End If
                    sep = ","   'After 1st one now add the seperator
                    iCTR += 1
                Next
                sw.WriteLine(builder.ToString())
            Next
            '******************************************************************
            '* 2015-07-30 RFK: 
            If Not sw Is Nothing Then sw.Close()
            lLabelRowUpdate.Text = "0"
            aoLetters.MsgStatus("Wrote " + dT.Rows.Count.ToString.Trim + " rows.", True)
            Return True
        Catch ex As Exception
            Msg_Error("DataTable_to_CSV", ex.ToString)
        End Try
        Return False
    End Function

    Public Function DataGridView_to_CSV(ByVal dG As DataGridView, ByVal sFileName As String, ByVal IncludeColumnHeader As Boolean, ByVal lLabelRowUpdate As Label) As Boolean
        Try
            '******************************************************************
            '* 2017-03-24 RFK:
            If File.Exists(sFileName) Then
                aoLetters.MsgStatus("Deleting:" + sFileName, True)
                File.Delete(sFileName)
            End If
            If File.Exists(sFileName) Then
                aoLetters.MsgStatus("Unable to create, exists:" + sFileName, True)
                Return False
            End If
            '******************************************************************
            '* 2017-03-24 RFK:
            aoLetters.MsgStatus("Creating:" + sFileName, True)
            '******************************************************************
            Dim sw As System.IO.StreamWriter
            sw = My.Computer.FileSystem.OpenTextFileWriter(sFileName, True)
            '******************************************************************
            Dim sTmpLine As String = ""
            '******************************************************************
            '* 2017-03-24 RFK:
            For i2 = 0 To dG.ColumnCount - 1
                If sTmpLine.Length > 0 Then sTmpLine += vbTab
                sTmpLine += dG.Columns(i2).Name
            Next
            sw.WriteLine(sTmpLine)
            '******************************************************************
            For iRow = 0 To dG.Rows.Count - 1
                lLabelRowUpdate.Text = Trim(Str(dG.Rows.Count - iRow))
                System.Windows.Forms.Application.DoEvents()
                '**************************************************************
                sTmpLine = ""
                For iCol = 0 To dG.ColumnCount - 1
                    If dG.Item(iCol, iRow).Value IsNot Nothing Then
                        If sTmpLine.Length > 0 Then sTmpLine += vbTab
                        sTmpLine += dG.Item(iCol, iRow).Value.ToString
                    End If
                Next
                sw.WriteLine(sTmpLine)
            Next
            '******************************************************************
            sw.Close()
            '******************************************************************
            lLabelRowUpdate.Text = "0"
            aoLetters.MsgStatus("Wrote " + dG.Rows.Count.ToString.Trim + " rows.", True)
            Return True
        Catch ex As Exception
            Msg_Error("DataGridView_to_CSV", ex.ToString)
        End Try
        Return False
    End Function

    Public Function SQL_READ_DATATABLE(ByVal dT As DataTable, ByVal tDB As String, ByVal tMODULE As String, ByVal tSQLConnectionString As String, ByVal tSQLuser As String, ByVal tSELECTstring As String) As Boolean
        Try
            Select Case tDB.ToUpper
                Case "DB2"
#If DB2 = 1 Then
                    If dT IsNot Nothing Then
                        Dim dbConnection As IBM.Data.DB2.iSeries.iDB2Connection = New IBM.Data.DB2.iSeries.iDB2Connection(tSQLConnectionString + tSQLuser)
                        Dim dbCommand As IBM.Data.DB2.iSeries.iDB2Command = New IBM.Data.DB2.iSeries.iDB2Command()
                        dbCommand.CommandText = tSELECTstring
                        dbCommand.Connection = dbConnection
                        dbCommand.CommandTimeout = 0
                        '******************************************************
                        Dim dAdapter As IBM.Data.DB2.iSeries.iDB2DataAdapter = New IBM.Data.DB2.iSeries.iDB2DataAdapter
                        dAdapter.SelectCommand = dbCommand
                        dT.Clear()
                        Dim dataSet As System.Data.DataSet = New System.Data.DataSet
                        dAdapter.Fill(dT)
                        '******************************************************
                        dbCommand.Dispose()
                        dbCommand = Nothing
                        dbConnection.Close()
                        dbConnection.Dispose()
                        dbConnection = Nothing
                        '******************************************************
                        If dT.Rows.Count > 0 Then Return True
                        '******************************************************
                    Else
                        aoLetters.MsgStatus("TABLE NOT DEFINED", True)
                    End If
#Else
                    eletters.MsgStatus("DB2 NOT DEFINED", True)
#End If
                Case "MSSQL"
                    Dim dbConnection As New SqlConnection(tSQLConnectionString + tSQLuser)        ' The SqlConnection class allows you to communicate with SQL Server.
                    Dim dbCommand As New SqlCommand(tSELECTstring, dbConnection)                ' A SqlCommand object is used to execute the SQL commands.
                    dbCommand.CommandTimeout = 0

                    Dim dAdapter As New SqlDataAdapter(dbCommand)
                    Dim mDataSet As New DataSet()
                    dAdapter.Fill(dT)

                    dbCommand.Dispose()
                    dbCommand = Nothing
                    dbConnection.Close()
                    dbConnection.Dispose()
                    dbConnection = Nothing
                    If dT.Rows.Count > 0 Then Return True
                Case "MYSQL"
                    '
                Case "FOXPRO"
                    Dim dbConnection As New OleDbConnection("Provider=vfpoledb.1;Data Source=" + tSQLConnectionString + ";Collating Sequence=machine")
                    Dim dbCommand As New OleDbCommand
                    Dim dbDataAdapter As New OleDbDataAdapter

                    dbCommand.CommandText = tSELECTstring
                    dbCommand.Connection = dbConnection

                    dbDataAdapter.SelectCommand = dbCommand
                    dbDataAdapter.Fill(dT)

                    dbCommand.Dispose()
                    dbCommand = Nothing
                    dbConnection.Close()
                    dbConnection.Dispose()
                    dbConnection = Nothing
                    If dT.Rows.Count > 0 Then Return True
            End Select
        Catch ex As Exception
            MsgBox(ex.ToString)
            aoLetters.MsgStatus(ex.ToString, True)
            Console.Write("SQL_READ_DATATABLE" + ex.ToString)
        End Try
        Return False
    End Function

    Public Function DataTable_ToHtmlTrTd(ByVal dT As DataTable, ByVal sTd1stAlign As String, ByVal sTdRestAlign As String) As String
        Try
            '******************************************************************
            '* 2016-08-30 RFK:
            MsgBox("NEED TO CONVERT")
            Dim sHTML As String = ""
            Dim iC As Integer = 0, iR As Integer = 0
            '******************************************************************
            sHTML += "<tr>"
            'For iC = 0 To dT.Columns.Count - 1
            '    If iC = 0 Then
            '        If dT.Item(iC, iR).Style.BackColor.ToString.Contains("Empty") Then
            '            If sTd1stAlign.Length > 0 Then
            '                sHTML += "<td align=" + sTd1stAlign + ">" + gGrid.Columns(iC).Name + "</td>"
            '            Else
            '                sHTML += "<td>" + gGrid.Columns(iC).Name + "</td>"
            '            End If
            '        Else
            '            If sTd1stAlign.Length > 0 Then
            '                sHTML += "<td align=" + sTd1stAlign + " style=background-color:" + gGrid.Item(iC, iR).Style.BackColor.ToString.Replace("Color [", "").Replace("]", "") + ">" + gGrid.Columns(iC).Name + "</td>"
            '            Else
            '                sHTML += "<td style=background-color:" + gGrid.Item(iC, iR).Style.BackColor.ToString.Replace("Color [", "").Replace("]", "") + ">" + gGrid.Columns(iC).Name + "</td>"
            '            End If
            '        End If
            '    Else
            '        If gGrid.Item(iC, iR).Style.BackColor.ToString.Contains("Empty") Then
            '            If sTdRestAlign.Length > 0 Then
            '                sHTML += "<td align=" + sTdRestAlign + ">" + gGrid.Columns(iC).Name + "</td>"
            '            Else
            '                sHTML += "<td>" + gGrid.Columns(iC).Name + "</td>"
            '            End If
            '        Else
            '            If sTdRestAlign.Length > 0 Then
            '                sHTML += "<td align=" + sTdRestAlign + " style=background-color:" + gGrid.Item(iC, iR).Style.BackColor.ToString.Replace("Color [", "").Replace("]", "") + ">" + gGrid.Columns(iC).Name + "</td>"
            '            Else
            '                sHTML += "<td style=background-color:" + gGrid.Item(iC, iR).Style.BackColor.ToString.Replace("Color [", "").Replace("]", "") + ">" + gGrid.Columns(iC).Name + "</td>"
            '            End If
            '        End If
            '    End If
            'Next
            sHTML += "</tr>"
            ''******************************************************************
            'For iR = 0 To gGrid.RowCount - 1
            '    sHTML += "<tr>"
            '    For iC = 0 To gGrid.ColumnCount - 1
            '        If gGrid.Item(iC, iR).Value IsNot Nothing Then
            '            'eLetters.MsgStatus(gGrid.Item(iC, iR).Style.BackColor.ToString, True)
            '            If iC = 0 Then
            '                If gGrid.Item(iC, iR).Style.BackColor.ToString.Contains("Empty") Then
            '                    If sTd1stAlign.Length > 0 Then
            '                        sHTML += "<td align=" + sTd1stAlign + ">" + gGrid.Item(iC, iR).Value.ToString.Trim + "</td>"
            '                    Else
            '                        sHTML += "<td>" + gGrid.Item(iC, iR).Value.ToString.Trim + "</td>"
            '                    End If
            '                Else
            '                    If sTd1stAlign.Length > 0 Then
            '                        sHTML += "<td align=" + sTd1stAlign + " style=background-color:" + gGrid.Item(iC, iR).Style.BackColor.ToString.Replace("Color [", "").Replace("]", "") + ">" + gGrid.Item(iC, iR).Value.ToString.Trim + "</td>"
            '                    Else
            '                        sHTML += "<td style=background-color:" + gGrid.Item(iC, iR).Style.BackColor.ToString.Replace("Color [", "").Replace("]", "") + ">" + gGrid.Item(iC, iR).Value.ToString.Trim + "</td>"
            '                    End If
            '                End If
            '            Else
            '                If gGrid.Item(iC, iR).Style.BackColor.ToString.Contains("Empty") Then
            '                    If sTdRestAlign.Length > 0 Then
            '                        sHTML += "<td align=" + sTdRestAlign + ">" + gGrid.Item(iC, iR).Value.ToString.Trim + "</td>"
            '                    Else
            '                        sHTML += "<td>" + gGrid.Item(iC, iR).Value.ToString.Trim + "</td>"
            '                    End If
            '                Else
            '                    If sTdRestAlign.Length > 0 Then
            '                        sHTML += "<td align=" + sTdRestAlign + " style=background-color:" + gGrid.Item(iC, iR).Style.BackColor.ToString.Replace("Color [", "").Replace("]", "") + ">" + gGrid.Item(iC, iR).Value.ToString.Trim + "</td>"
            '                    Else
            '                        sHTML += "<td style=background-color:" + gGrid.Item(iC, iR).Style.BackColor.ToString.Replace("Color [", "").Replace("]", "") + ">" + gGrid.Item(iC, iR).Value.ToString.Trim + "</td>"
            '                    End If
            '                End If
            '            End If
            '        Else
            '            sHTML += "<td></td>"
            '        End If
            '    Next
            '    sHTML += "</tr>"
            'Next
            Return sHTML
        Catch ex As Exception
            Msg_Error("DataGridView_ToHtmlTable", ex.ToString)
        End Try
        Return -1
    End Function

    Public Function ComboBox_InitYesNo(ByVal dDrop As ComboBox, ByVal iAddABlankFirst As Boolean, ByVal iYesFirstThenNo As Boolean)
        Try
            dDrop.Items.Clear()
            If iAddABlankFirst Then dDrop.Items.Add(" ")
            If iYesFirstThenNo Then
                dDrop.Items.Add("Yes")
                dDrop.Items.Add("No")
            Else
                dDrop.Items.Add("No")
                dDrop.Items.Add("Yes")
            End If
            Return True
        Catch ex As Exception
            Msg_Error("ComboBox_InitYesNo", ex.ToString)
            Return False
        End Try
    End Function

End Module
