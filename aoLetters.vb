
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO

Public Class aoLetters
    Public sVersion As String = "Annuity One Letters v202206.21"
    'Public sSITE As String = "iTeleCollect"
    Public sSITE As String = "AnnuityOne"
    Public sDBO As String = "RevMD"
    '**************************************************************************
    '* 2018-08-07 RFK: NEXT NEW LETTER TYPE = 30
    '**************************************************************************
    Public swReadAllMatched As Boolean = False
    Public swDTable As Boolean = False

    Public msSQLConnectionString As String = ""
    Public msSQLuser As String = ""

    Public DB2SQLConnectionString As String = ""
    Public DB2SQLuser As String = ""
    '**************************************************************************
    '* 2018-08-07 RFK: "\\TeleServer\TeleServer$\DATA\" is GONE
    Public dir_CHK As String = "\\production\automation$\CHK\"
    Public dir_DATA As String = "\\production\required_files\INI\"
    Public dir_EXE As String = "\\production\Required_Files\EXE\"
    'Public dir_FTP As String = "\\TeleServer\TeleServer$\SCHEDULE\FTP\"
    Public dir_FTP As String = "\\production\automation$\FTP\"
    Public dir_INI As String = "\\production\Required_Files\INI\"
    Public dir_LOG As String = "\\reporting\report$\LOG\"
    'Public dir_LOG As String = "\\production\reports\LOG\"
    Public dir_SCHEDULE As String = "\\production\automation$\SCHEDULE\"
    Public dir_LETTERS As String = "\\production\Process\LETTERS\SENT\"
    Public dir_TEST As String = "\\production\Process\LETTERS\TEST\"
    Public dir_PRODUCTION As String = "\\production\PROCESS\"
    Public dir_REPORTS As String = "\\production\REPORTS\"
    Public dir_REPORTING As String = "\\reporting\REPORT$\"

    Public iRow As Integer = 0
    Public tSysAccount As String = "", tAccountNumber As String = "", tAccountSuffix As String = "", tLOCX As String = "", tRamLOCX As String = "", tmLOCX As String = "", tsLOCX As String = ""
    Public tGSSN As String = "", tsGSSN As String = ""
    Public tGNAMEL As String = "", tsGNAMEL As String = "", tGNAMEF As String = "", tsGNAMEF As String = "", tGNAMEM As String = "", tsGNAMEM As String = ""
    Public tGADDR As String = "", tsGADDR As String = ""
    Public tGCITY As String = "", tsGCITY As String = ""
    Public tGZIP As String = "", tsGZIP As String = ""
    Public tGDOB As String = "", tsGDOB As String = ""
    Public tPNAMEL As String = "", tsPNAMEL As String = "", tPNAMEF As String = "", tsPNAMEF As String = "", tPNAMEM As String = "", tsPNAMEM As String = ""
    Public tPSSN As String = "", tsPSSN As String = ""
    Public tMedRec As String = "", tsMedRec As String = ""
    Public dBalance As Double = 0, dBalanceRU As Double = 0, dAmount As Double = 0, dTotal As Double = 0
    Public iField1Pad As Integer = 25, iField2Pad As Integer = 12, iField3Pad As Integer = 12
    Public tRALOCX As String = ""
    Public sRAMLOCX As String = "", sPlacementDate As String = "", sPlacementDate1 As String = "", sPlacementDate1S As String = ""
    Public sSysAccountMatched As String = ""
    Public sIpOp As String = "", tRACLOS As String = ""
    Public sRalNAC As String = "", sRAACCT As String = ""
    Public sRABALD As String = ""
    Public bRollUp As Boolean = False
    Public sDOS As String = "", sDOSsave As String = ""
    Public tCurrentLetterVendor As String = "", tLetterPrinted As String = "", tLetterPrintedDate As String = ""
    Public tLetterCurrent As String, tLetterNext As String = "", tLetterNextDays As String = ""
    Public sCurrentBalance As String = "", sMatchedBalance As String = ""
    Public sPPLamount As String = "", sPPLpaid As String = "", sPPLdate As String = ""
    Public sTmpSTR As String = ""
    Public sStatusCodeBadGuarantorName As String = "BGN"
    Public iGhosts_CTR As Integer = 0
    Public DTable As DataTable, dTable_Select As New DataTable
    Public sClient As String = ""
    Dim dSummaryLine As Double = 0
    Dim tSUM As String
    Dim PrintLine_tLINE As String = "", tQuote As String = Chr(34), tValue As String = ""
    Dim tDelimiter As String = "|"
    Dim tSuffix As String = ""
    Dim tAddress As String = "", tAddress2 As String = "", tCity As String = "", tState As String = "", tStateAllow As String = "", tZip As String = "", tZip4 As String = ""
    Dim tSTR As String = "", sScanLine As String = "", sClientGroup As String = "", sSentBalance As String = ""
    Dim sSQL As String = "", sSQLout As String = "", sPromiseAmount As String = ""
    Dim sDate As String = ""
    Dim gtLetterType As String = "1"
    Dim swOverRideDate As Boolean = True
    Dim sLocxs As String = "", sAccountSave As String = "", sAmount As String = "", sAmountNew As String = "", sRAFACL As String = ""
    Dim MatchedFuture_SQLstring As String = ""
    Dim MatchedFuture_iLocxRow As Integer = 0
    Dim MatchedFuture_iMatchedRow As Integer = 0, CalculateCAPB_iLocxRow As Integer = 0
    Dim bOK As Boolean = False
    Dim MatchedPrint_iMatchedRow As Integer = 0
    Dim tMultiMessage As String = ""
    Dim MultiAccounts_sSysAccountMatched As String = ""
    Dim MultiAccounts_imLocxRow As Integer = 0
    Dim MultiAccounts_iMatchedRow As Integer = 0
    Dim AnnuityOne_MatchedCheck_tLetterNumber As String = "", AnnuityOne_MatchedCheck_tLetterDate As String = ""
    Dim AnnuityOne_MatchedCheck_iMatchedRow As Integer = 0
    Dim tLetterNextDate As String = ""
    Dim iRulesRow As Integer = 0
    Dim iNumPrinted As Integer = 0
    Dim iTestLetter As Integer = 0
    Dim sHTML As String = ""

    Dim MatchedPrint_sSysAccount As String = ""
    Dim MatchedPrint_iLocxRow As Integer = 0
    Dim MatchedPrint_iPrintedLines As Integer = 0
    Dim tDESCR As String = ""
    Dim imLocxRow As Integer = 0
    Dim MatchedPrint_imMatchedRow As Integer = 0
    Dim letter_sent_tLetterCounter As String = "", tRABALD As String = "", tRAMBAL As String = ""
    Dim iLocxMultiRow As Integer = 0
    Dim iLetterRow As Integer = 0
    Dim sTemp As String = ""
    Dim dTemp As Double = 0, dTemp2 As Double = 0
    Dim sCharges As String, sInsurance As String, sSelf As String, sCredits As String
    Dim tFacility As String = ""
    Dim iFacilityRow As Integer = 0
    Dim iCrow As Integer = 0
    Public SentBalance As Double, SentBalanceDollars30 As Double, SentBalanceDollars60 As Double, SentBalanceDollars90 As Double, SentBalanceDollars120 As Double, SentBalanceDollars121 As Double
    Public sEncryptionPW As String = "Ryan.Kiechle_2194404440"

    Private Sub eLetters_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Text = sVersion
            '****************************************************
            '* 2012-04-27 RFK:
            '* 2011-11-01 RFK: commands
            '* 2021-08-31 RFK: Letter Type 26 = Lubbock
            Select Case sSITE
                Case "AnnuityOne"
                    msSQLConnectionString = rkutils.Decrypt(".NET", rkutils.FILE_read(dir_DATA + "msSQLConnectionString.INI"), sEncryptionPW)
                    msSQLuser = rkutils.Decrypt(".NET", rkutils.FILE_read(dir_DATA + "sa_Dialer-PW.INI"), sEncryptionPW)
                    DB2SQLConnectionString = rkutils.Decrypt(".NET", rkutils.FILE_read(dir_DATA + "DB2SQLConnectionString.INI"), sEncryptionPW)
                    DB2SQLuser = rkutils.Decrypt(".NET", rkutils.FILE_read(dir_DATA + "aoMatching-PW.INI"), sEncryptionPW)
                Case "iTeleCollect", "TeleCollect"
                    If IS_File("C:\Tele\TCS-DIR.INI") Then
                        Dim tLINE As String = FILE_read("C:\Tele\TCS-DIR.INI")
                        dir_PRODUCTION = STR_NORMALIZE(LineRead(tLINE, 0))
                        dir_INI = STR_NORMALIZE(LineRead(tLINE, 1))
                        dir_REPORTS = STR_NORMALIZE(LineRead(tLINE, 2))
                        dir_LETTERS = STR_NORMALIZE(LineRead(tLINE, 3))
                        dir_TEST = STR_NORMALIZE(LineRead(tLINE, 4))
                        dir_FTP = STR_NORMALIZE(LineRead(tLINE, 5))
                    End If
                    If IS_File("C:\Tele\TCS.INI") Then
                        '******************************************************
                        '* 2015-06-09 RFK:
                        Dim tLINE As String = FILE_read("C:\Tele\TCS.INI")
                        msSQLConnectionString = rkutils.Decrypt(".NET", rkutils.FILE_read(dir_DATA + "msSQLConnectionString.INI"), sEncryptionPW)
                        msSQLuser = rkutils.Decrypt(".NET", rkutils.FILE_read(dir_DATA + "sa_Dialer-PW.INI"), sEncryptionPW)
                        '******************************************************
                    End If
            End Select
            MsgStatus("INI=" + dir_INI, True)
            MsgStatus("LETTERS=" + dir_LETTERS, True)
            MsgStatus("FTP=" + dir_FTP, True)
            MsgStatus(msSQLConnectionString, True)
            '****************************************************
            InitApp()
            '****************************************************
            Timer1.Enabled = False
            Timer1.Interval = 1000
            '****************************************************
            '* 2014-08-26 RFK:
            ComboBox_Processor.Items.Clear()
            ComboBox_Processor.Items.Add("ALL")
            ComboBox_Processor.Items.Add("1")
            ComboBox_Processor.Items.Add("2")
            ComboBox_Processor.Items.Add("3")
            ComboBox_Processor.Items.Add("4")
            ComboBox_Processor.Items.Add("5")
            ComboBox_Processor.Items.Add("6")
            ComboBox_Processor.Items.Add("7")
            ComboBox_Processor.Items.Add("8")
            ComboBox_Processor.Items.Add("9")
            ComboBox_Processor.SelectedIndex = 0
            '******************************************************************
            ComboBox_MatchType.Items.Clear()
            ComboBox_MatchType.Items.Add("Active Clients")
            ComboBox_MatchType.Items.Add("Testing Clients")
            ComboBox_MatchType.Items.Add("Sent Today")
            ComboBox_MatchType.Items.Add("Inactive Clients")
            ComboBox_MatchType.SelectedIndex = 0
            '******************************************************************
            '* 2021-11-26 RFK:
            rkutils.ComboBox_InitYesNo(ComboBox_TestAll, False, False)
            rkutils.ComboBox_InitYesNo(ComboBox_IgnoreDate, False, False)
            '******************************************************************
            ComboBox_ClientActive.Items.Clear()
            ComboBox_ClientActive.Items.Add("Yes letters")
            ComboBox_ClientActive.Items.Add("Testing")
            ComboBox_ClientActive.Items.Add("No letters")
            ComboBox_ClientActive.Items.Add("test System")
            ComboBox_ClientActive.SelectedIndex = 0
            '******************************************************************
            '* 2016-08-19 RFK: 
            ComboBox_Facility.Items.Clear()
            ComboBox_Facility.Items.Add("=")
            ComboBox_Facility.Items.Add("<>")
            ComboBox_Facility.Items.Add("<=")
            ComboBox_Facility.SelectedIndex = 0
            '******************************************************************
            '* 2022-02-10 RFK: 
            ComboBox_DOP.Items.Clear()
            ComboBox_DOP.Items.Add("")
            ComboBox_DOP.Items.Add(">=")
            ComboBox_DOP.Items.Add("<=")
            ComboBox_DOP.Items.Add("=")
            ComboBox_DOP.SelectedIndex = 0
            '******************************************************************
            '* 2017-06-02 RFK:
            ComboBox_Ghosts.Items.Clear()
            ComboBox_Ghosts.Items.Add("No")
            ComboBox_Ghosts.Items.Add("Yes")
            ComboBox_Ghosts.SelectedIndex = 0
            '**************************************************************************************
            '* 2022-03-25 RFK:
            ComboBox_RATZ.Items.Clear()
            ComboBox_RATZ.Items.Add("")
            ComboBox_RATZ.Items.Add("RATZ01")
            ComboBox_RATZ.Items.Add("RATZ02")
            ComboBox_RATZ.Items.Add("RATZ03")
            ComboBox_RATZ.Items.Add("RATZ04")
            ComboBox_RATZ.Items.Add("RATZ05")
            ComboBox_RATZ.Items.Add("RATZ06")
            ComboBox_RATZ.Items.Add("RATZ07")
            ComboBox_RATZ.Items.Add("RATZ08")
            ComboBox_RATZ.Items.Add("RATZ09")
            ComboBox_RATZ.Items.Add("RATZ10")
            ComboBox_RATZ.Items.Add("RATZ11")
            ComboBox_RATZ.Items.Add("RATZ12")
            ComboBox_RATZ.Items.Add("RATZ13")
            ComboBox_RATZ.Items.Add("RATZ14")
            ComboBox_RATZ.Items.Add("RATZ15")
            ComboBox_RATZ.Items.Add("RATZ16")
            ComboBox_RATZ.Items.Add("RATZ17")
            ComboBox_RATZ.Items.Add("RATZ18")
            ComboBox_RATZ.Items.Add("RATZ19")
            ComboBox_RATZ.Items.Add("RATZ20")
            ComboBox_RATZ.Items.Add("RATZ21")
            ComboBox_RATZ.Items.Add("RATZ22")
            ComboBox_RATZ.Items.Add("RATZ23")
            ComboBox_RATZ.Items.Add("RATZ24")
            ComboBox_RATZ.Items.Add("RATZ25")
            ComboBox_RATZ.SelectedIndex = 0
            '**************************************************************************************
            '* 2021-11-17 RFK:
            sTemp = "c:\tele\as400-connection.bat"
            If IS_File(sTemp) Then
                MsgStatus(sTemp, True)
                Shell(sTemp, AppWinStyle.NormalFocus)
            End If
            '******************************************************************
            RunReady()
            ComboBox_SetValue(ComboBox_ClientActive, "Active Clients")
            ClientsInit()
            '****************************************************
            Dim tCommand As String = Command.ToString
            'tCommand = "/ALL"
            'tCommand = "/SENTSUMMARY"
            'tCommand = "/1"
            'tCommand = "/2"
            'tCommand = "/NOTTODAY" 
            'tCommand = "/CLIENT NW2"
            Select Case STR_BREAK(tCommand, 1)
                Case "/SENTSUMMARY"
                    '**********************************************************
                    sHTML = "<html><body>"
                    '**********************************************************
                    sHTML += "<table border=1>"
                    '**********************************************************
                    LettersSentLastDays()
                    '**********************************************************
                    '* 2019-03-14 RFK:
                    If swDTable Then
                        sHTML += "<tr><td colspan=" + Trim(Str(dTable_Select.Columns.Count)) + ">Letters sent</td></tr>"
                        sHTML += rkutils.DataTable_ToHtmlTrTd(dTable_Select, "LEFT", "RIGHT")
                    Else
                        sHTML += "<tr><td colspan=" + Trim(Str(DataGridView_Select.ColumnCount)) + ">Letters sent</td></tr>"
                        sHTML += rkutils.DataGridView_ToHtmlTrTd(DataGridView_Select, "LEFT", "RIGHT")
                    End If
                    sHTML += "</table>"
                    sHTML += "</br>"
                    '**********************************************************
                    '* 2017-04-18 RFK: 
                    sHTML += "<table border=1>"
                    'LettersNotSent()
                    '**********************************************************
                    '* 2019-03-14 RFK:
                    If swDTable Then
                        sHTML += "<tr><td colspan=" + Trim(Str(dTable_Select.Columns.Count)) + ">Letters Not sent</td></tr>"
                        sHTML += rkutils.DataTable_ToHtmlTrTd(dTable_Select, "LEFT", "RIGHT")
                    Else
                        sHTML += "<tr><td colspan=" + Trim(Str(DataGridView_Select.ColumnCount)) + ">Letters Not sent</td></tr>"
                        sHTML += rkutils.DataGridView_ToHtmlTrTd(DataGridView_Select, "LEFT", "RIGHT")
                    End If
                    sHTML += "</table>"
                    sHTML += "</br>"
                    '**********************************************************
                    LettersSentLastMonths(12)
                    '**********************************************************
                    '* 2019-03-14 RFK:
                    If swDTable Then
                        MsgStatus("Need to CONVERT", True)
                    Else
                        If DataGridView_Select IsNot Nothing Then
                            If DataGridView_Select.RowCount > 0 Then
                                '******************************************************
                                '* 2016-08-31 RFK: Sum the Rows
                                For iR = 0 To DataGridView_Select.RowCount - 1
                                    DataGridView_Select.Item(DataGridView_Select.ColumnCount - 1, iR).Value = rkutils.STR_format(DataGridView_SumRow(DataGridView_Select, iR, 1, DataGridView_Select.ColumnCount - 2), "#,###")
                                Next
                                '******************************************************
                                '* 2016-08-31 RFK: Sum the Columns
                                DataGridView_Select.Item(0, DataGridView_Select.RowCount - 1).Value = "Total"
                                Dim i1 As Integer
                                For i1 = 1 To Trim(Str(DataGridView_Select.ColumnCount - 1))
                                    DataGridView_Select.Item(i1, DataGridView_Select.RowCount - 1).Value = rkutils.STR_format(DataGridView_SumColumn(DataGridView_Select, i1, False), "#,###")
                                Next
                                '******************************************************
                                '* 2016-09-07 RFK: Format the column
                                Dim iC As Integer = 2
                                For iR = 1 To DataGridView_Select.RowCount - 1
                                    If DataGridView_Select.Item(iC, iR).Value IsNot Nothing Then
                                        If Val(DataGridView_Select.Item(iC, iR).Value.ToString) > 0 Then
                                            'DataGridView_Select.Item(iC, iR).Value = rkutils.STR_format(DataGridView_Select.Item(iC, iR).Value.ToString, "#,###")
                                        End If
                                    End If
                                Next
                                'DataGridView_FormatColumn(DataGridView_Select, 1, "#,###", True)
                                'DataGridView_FormatColumn(DataGridView_Select, 2, "#,###", True)
                            End If
                            '**********************************************************
                            sHTML += "<table border=1>"
                            sHTML += "<tr><td colspan=14>Letters sent for last 12 Months</td></tr>"
                            '**********************************************************
                            If DataGridView_Select.RowCount > 0 Then
                                sHTML += rkutils.DataGridView_ToHtmlTrTd(DataGridView_Select, "LEFT", "RIGHT")
                            End If
                            sHTML += "</table>"
                        End If
                    End If
                    '**********************************************************
                    sHTML += "</body></html>"
                    '**********************************************************
                    rkutils.EMAILIT(msSQLConnectionString, msSQLuser, "DoNotReply@AnnuityHealth.com", "eLetters", "LETTERS@AnnuityHealth.com", "Letters", Me.Text, "eLetters Sent Summary", "", sHTML, "")
                    '**********************************************************
                    End
                    '**********************************************************
                Case "/TCS"
                    'ComboBox_MatchType.Items.Add("TCS")
                Case "/TODAY", "/YESTERDAY", "/YESTERDAYTODAY"
                    '**********************************************************
                    '* 2015-08-31 RFK: Yesterday
                    Select Case STR_BREAK(tCommand, 1)
                        Case "/TODAY"
                            TextBox_TCodeDate.Text = "TODAY"
                        Case "/YESTERDAY"
                            TextBox_TCodeDate.Text = "YESTERDAY"
                        Case "/YESTERDAYTODAY"
                            TextBox_TCodeDate.Text = "YESTERDAYTODAY"
                    End Select
                    '**********************************************************
                    Dim iRow As Integer = rkutils.DataGridView_Contains(DataGridView_Clients, "CLIENTNAME", rkutils.STR_BREAK_PIECES(tCommand, 2, " "))
                    If iRow >= 0 Then
                        DataGridView_Clients.CurrentCell = DataGridView_Clients.Rows(iRow).Cells(0)
                        Label_ClientRow.Text = DataGridView_Clients.CurrentCellAddress.Y.ToString
                        AccountInitSettings()
                        TextBox_TCodes.Text = rkutils.STR_BREAK_PIECES(tCommand, 3, " ")
                        If Panel_RegF.Visible Then
                            MsgStatus("RegF", True)
                        Else
                            AccountsLoadForClient()
                            Run()
                            If Val(Label_Printed.Text) > 0 Then FTP_put()
                        End If
                    End If
                    End
                Case "/ALL", "/RUN", "/1", "/2", "/3", "/4", "/5", "/6", "/7", "/8", "/9"
                    '**********************************************************
                    '* 2016-01-22 RFK: Do Not Run on Weekends
                    If Now.DayOfWeek = DayOfWeek.Saturday Or Now.DayOfWeek = DayOfWeek.Sunday Then
                        rkutils.EMAILIT(msSQLConnectionString, msSQLuser, "DoNotReply@AnnuityHealth.com", "eLetters", "LETTERS@AnnuityHealth.com", "IT", Me.Text, "eLetters", "eLetters Does NOT run on Weekends", "", "")
                        MsgStatus("Do Not Run on Weekends", True)
                    Else
                        sSQL = "select description from " + sDBO + ".dbo.holidays where tdate ='" + rkutils.STR_format("TODAY", "ccyymmdd") + "'"
                        '******************************************************
                        '* 2019-03-14 RFK:
                        If swDTable Then
                            sTemp = rkutils.SQL_READ_FIELD_DataTable(dTable_Select, "MSSQL", "DESCRIPTION", msSQLConnectionString, msSQLuser, sSQL)
                        Else
                            sTemp = rkutils.SQL_READ_FIELD(DataGridView_Select, "MSSQL", "DESCRIPTION", msSQLConnectionString, msSQLuser, sSQL)
                        End If
                        MsgStatus(sSQL + " [" + sTemp + "]", True)
                        If sTemp.Trim.Length > 0 Then
                            rkutils.EMAILIT(msSQLConnectionString, msSQLuser, "DoNotReply@AnnuityHealth.com", "eLetters", "LETTERS@AnnuityHealth.com", "IT", Me.Text, "eLetters", "eLetters Does NOT run on Holidays " + sTemp, "", "")
                            MsgStatus("Do Not Run on Holidays", True)
                        Else
                            MsgStatus("Running ALL", True)
                            '**************************************************
                            '* 2021-04-05 RFK:
                            Select Case STR_BREAK(tCommand, 1)
                                Case "/1", "/2", "/3", "/4", "/5", "/6", "/7", "/8", "/9"
                                    rkutils.ComboBox_SetValue(ComboBox_Processor, STR_BREAK(tCommand.Replace("/", ""), 1))
                                    If Val(STR_BREAK(tCommand.Replace("/", ""), 1)) > 1 Then
                                        Me.Top = (Val(STR_BREAK(tCommand.Replace("/", ""), 1)) * 20)
                                        Me.Left = (Val(STR_BREAK(tCommand.Replace("/", ""), 1)) * 20)
                                    Else
                                        Me.Top = 1
                                        Me.Left = 1
                                    End If
                            End Select
                            '**************************************************
                            '* 2012-01-22 RFK:
                            CheckBox_Update.Checked = True
                            Me.Show()
                            Application.DoEvents()
                            RunALL()
                        End If
                    End If
                    '**********************************************************
                    End
                Case "/CLIENT"
                    '**********************************************************
                    '* 2016-01-22 RFK: Do Not Run on Weekends
                    If Now.DayOfWeek = DayOfWeek.Saturday Or Now.DayOfWeek = DayOfWeek.Sunday Then
                        rkutils.EMAILIT(msSQLConnectionString, msSQLuser, "DoNotReply@AnnuityHealth.com", "eLetters", "LETTERS@AnnuityHealth.com", "IT", Me.Text, "eLetters", "eLetters Does NOT run on Weekends", "", "")
                        MsgStatus("Do Not Run on Weekends", True)
                    Else
                        sSQL = "select description from " + sDBO + ".dbo.holidays where tdate ='" + rkutils.STR_format("TODAY", "ccyymmdd") + "'"
                        '******************************************************
                        '* 2019-03-14 RFK:
                        If swDTable Then
                            sTemp = rkutils.SQL_READ_FIELD_DataTable(dTable_Select, "MSSQL", "DESCRIPTION", msSQLConnectionString, msSQLuser, sSQL)
                        Else
                            sTemp = rkutils.SQL_READ_FIELD(DataGridView_Select, "MSSQL", "DESCRIPTION", msSQLConnectionString, msSQLuser, sSQL)
                        End If
                        MsgStatus(sSQL + " [" + sTemp + "]", True)
                        If sTemp.Trim.Length > 0 Then
                            rkutils.EMAILIT(msSQLConnectionString, msSQLuser, "DoNotReply@AnnuityHealth.com", "eLetters", "LETTERS@AnnuityHealth.com", "IT", Me.Text, "eLetters", "eLetters Does NOT run on Holidays " + sTemp, "", "")
                            MsgStatus("Do Not Run on Holidays", True)
                        Else
                            '**************************************************
                            '* 2012-01-22 RFK:
                            MsgStatus("Running CLIENT " + rkutils.STR_BREAK_PIECES(tCommand, 3, " "), True)
                            '**************************************************
                            CheckBox_Update.Checked = True
                            Me.Show()
                            Application.DoEvents()
                            '**********************************************************************
                            Dim iRow As Integer = rkutils.DataGridView_Contains(DataGridView_Clients, "CLIENTNAME", rkutils.STR_BREAK_PIECES(tCommand, 2, " "))
                            If iRow >= 0 Then
                                DataGridView_Clients.CurrentCell = DataGridView_Clients.Rows(iRow).Cells(0)
                                Label_ClientRow.Text = DataGridView_Clients.CurrentCellAddress.Y.ToString
                                AccountInitSettings()
                                TextBox_TCodes.Text = rkutils.STR_BREAK_PIECES(tCommand, 3, " ")
                                If Panel_RegF.Visible Then
                                    '**************************************************************
                                    '* 2021-12-21 RFK: 910
                                    AnnuityOne_AccountsLoad(rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "TOB", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "FACILITY", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTMATCHTYPE", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "MATCHBY", Val(Label_ClientRow.Text)), "", "*", 1)
                                    Application.DoEvents()
                                    Run()
                                    If Val(Label_Printed.Text) > 0 Then FTP_put()
                                    '**************************************************************
                                    '* 2022-06-20 RFK: PP
                                    ' Ryan IS
                                    AnnuityOne_AccountsLoad(rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "TOB", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "FACILITY", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTMATCHTYPE", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "MATCHBY", Val(Label_ClientRow.Text)), "", "*", 3)
                                    Application.DoEvents()
                                    Run()
                                    If Val(Label_Printed.Text) > 0 Then FTP_put()
                                    '**************************************************************
                                    '* 2021-12-21 RFK: NOT 910 NOT PP
                                    AnnuityOne_AccountsLoad(rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "TOB", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "FACILITY", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTMATCHTYPE", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "MATCHBY", Val(Label_ClientRow.Text)), "", "*", 2)
                                    Application.DoEvents()
                                    Run()
                                    If Val(Label_Printed.Text) > 0 Then FTP_put()
                                    '**************************************************************
                                Else
                                    AccountsLoadForClient()
                                    Run()
                                    If Val(Label_Printed.Text) > 0 Then FTP_put()
                                End If
                            End If
                        End If
                    End If
                    '**********************************************************
                    'End
                    '**********************************************************
            End Select
            '******************************************************************
        Catch ex As Exception
            MsgError("eLetters_Load", ex.ToString)
        End Try
    End Sub

    Private Sub InitApp()
        Dim tFileINI As String = dir_INI + "eLetters-" + rkutils.WhoAmI() + ".INI"
        If IS_File(tFileINI) Then
            Dim tREAD As String = FILE_read(tFileINI)
            Dim iTop As Integer = Val(STR_BREAK(tREAD, 1))
            Dim iLeft As Integer = Val(STR_BREAK(tREAD, 2))
            Me.Top = iTop
            Me.Left = iLeft
            Me.Left = 1
            Me.Top = 1
        End If
        '**********************************************************************
        '* 2021-04-02 RFK:
        Dim tooltipStrip As New ToolTip
        tooltipStrip.Active = True
        tooltipStrip.SetToolTip(TextBox_Processor, "Which letter processor to run on")
        tooltipStrip.SetToolTip(TextBox_MaxLetters, "Maximum number of letters to send")
        tooltipStrip.SetToolTip(TextBox_MaxAccounts, "Maximum number of accounts to load")
        tooltipStrip.SetToolTip(TextBox_LettersOnly, "")
        tooltipStrip.SetToolTip(TextBox_DateMM, "mm (Month 2 digit)")
        tooltipStrip.SetToolTip(TextBox_DateDD, "dd (Day 2 digit)")
        tooltipStrip.SetToolTip(TextBox_DateCCYY, "ccyy (Year 4 digits")
        '**********************************************************************
    End Sub

    Private Sub eLetters_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        ExitApp()
    End Sub

    Private Sub ExitApp()
        FILE_create(dir_INI + "eLetters-" + rkutils.WhoAmI() + ".INI", True, False, Me.Top.ToString + " " + Me.Left.ToString + vbCrLf)
        End
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Try
            'Label_TIME.Text = STR_TRIM(Now.TimeOfDay.ToString, 8)
            If Label_RUNNING.Text = "Ready" Then
                'ProcessInit()
            End If
            FILE_create(dir_CHK + "eLetters.CHK", True, False, Date.Now.ToString)
        Catch ex As Exception
            MsgError("Timer1", ex.ToString)
        End Try
    End Sub

    Private Sub MsgError(ByVal tMODULE As String, ByVal tMSG As String)
        MsgStatus("ERROR:" + tMODULE + "_" + tMSG, True)
    End Sub

    Public Sub MsgStatus(ByVal tMSG As String, bToScreen As Boolean)
        '***************************************************************************
        FILE_create(dir_LOG + "eLetters-" + DateToday(8) + ".LOG", False, True, DateToday(18) + " " + tMSG + vbCrLf)
        '***************************************************************************
        If bToScreen Then
            Me.ListBox_LOG.Items.Add(tMSG)
            If Me.ListBox_LOG.Items.Count > 500 Then
                Me.ListBox_LOG.Items.RemoveAt(1)
            End If
            Me.ListBox_LOG.SelectedIndex = Me.ListBox_LOG.Items.Count - 1
        End If
    End Sub

    Private Sub ListBox_LOG_MouseDoubleClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles ListBox_LOG.MouseDoubleClick
        Shell("NOTEPAD.EXE " + dir_LOG + "eLetters-" + DateToday(8) + ".LOG", AppWinStyle.NormalFocus)
    End Sub

    Private Function RunStart() As Boolean
        Label_RUNNING.Text = "Running"
        Return True
    End Function

    Private Function RunReady() As Boolean
        Label_RUNNING.Text = "Ready"
        Return True
    End Function

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs)
        If Label_RUNNING.Text = "Running" Then
            RunReady()
            Exit Sub
        End If
    End Sub

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs) Handles Button6.Click
        RunALL()
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        ExitApp()
    End Sub

    Private Sub ComboBox_MatchType_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox_MatchType.SelectedIndexChanged
        ClientsInit()
    End Sub

    Private Sub Label19_Click(sender As Object, e As EventArgs) Handles Label19.Click

    End Sub

    Private Sub Label20_Click(sender As Object, e As EventArgs) Handles Label20.Click
        Try
            AOdisplayLocx()
        Catch ex As Exception
            MsgError("Start", ex.ToString)
        End Try
    End Sub

    Private Sub AOdisplayLocx()
        Try
            Dim myProcess As New Process
            myProcess.StartInfo.UseShellExecute = True
            myProcess.StartInfo.FileName = "https://it.annuityOne.com/display_account.aspx?ID=" + TextBox_TEST.Text
            myProcess.Start()
            myProcess.Dispose()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Label14_Click(sender As System.Object, e As System.EventArgs) Handles Label14.Click
        LOCX_pop()
    End Sub


    Private Sub LOCX_pop()
        Try
            If TextBox_Locx.Text.Length > 0 Then
                Dim myProcess As New Process
                myProcess.StartInfo.UseShellExecute = True
                myProcess.StartInfo.FileName = "https:\\it.annuityone.com/display_account.aspx?IDR=" + TextBox_Locx.Text
                myProcess.Start()
            End If
        Catch ex As Exception
            MsgError("Start", ex.ToString)
        End Try
    End Sub

    Private Sub FormatLineOut(ByRef tLine As String, ByVal tAdd As String, ByVal iAddDelim As Boolean)
        tLine += Chr(34) + tAdd + Chr(34)
        If iAddDelim Then tLine += ","
    End Sub

    Private Sub Label_LettersOutput_Click(sender As System.Object, e As System.EventArgs) Handles Label_LettersOutput.Click
        Try
            Dim myProcess As New Process
            myProcess.StartInfo.UseShellExecute = True
            If Me.CheckBox_Update.Checked = True Then
                myProcess.StartInfo.FileName = dir_LETTERS
            Else
                myProcess.StartInfo.FileName = dir_TEST
            End If
            myProcess.Start()
        Catch ex As Exception
            MsgError("Label_LettersOutput_Click", ex.ToString)
        End Try
    End Sub

    Private Sub Letter_Types_Load(ByVal tClient As String, ByVal tTOB As String)
        Try
            '******************************************************************
            '* RFK:
            sSQL = "SELECT LNUMBER, LTYPE, ClientGroup, Company, Vendor, ChargeOffDate, BypassRulesDate, InsuranceJoin, FutureDays, MAXWIDTH, MAXROWS,BCAP,BadDebtNotification,FirstLetter,FinalLetter,PayPlanLetter"
            '******************************************************************
            '* 2015-07-28 RFK:
            Select Case sSITE
                Case "AnnuityOne"
                    'sSQL += ""
                Case "iTeleCollect"
                    'sSQL += ""
            End Select
            '******************************************************************
            sSQL += " FROM " + sDBO + ".dbo.letter_types"
            sSQL += " WHERE CLIENT='" + tClient + "' OR CLIENT='*'"
            Select Case sSITE
                Case "AnnuityOne"
                    sSQL += " AND (TOB='" + tTOB + "' OR TOB='*')"
                Case "iTeleCollect"
                    '
            End Select
            sSQL += " ORDER BY LNUMBER"
            MsgStatus(sSQL, False)
            If rkutils.SQL_READ_DATAGRID(DataGridView_Letter_Types, "MSSQL", "*", msSQLConnectionString, msSQLuser, sSQL) Then
                DataGridView_Letter_Types.Visible = True
                gtLetterType = rkutils.ReadField(DataGridView_Letter_Types, "LTYPE", 0)
            End If
            MsgStatus("Letter_Types_Load", True)
        Catch ex As Exception
            MsgError("Letter_Types_Load", ex.ToString)
        End Try
    End Sub

    Private Sub Facility_Load(ByVal tClient As String, ByVal tTOB As String)
        Try
            '**********************************************
            '* 2012-10-12 RFK:
            sSQL = "SELECT *"
            Select Case sSITE
                Case "AnnuityOne"
                    sSQL += " FROM ROIDATA.FACILP F LEFT JOIN ROIDATA.HFCLNP C ON F.FARCL#=C.HFPCID"
                    sSQL += " WHERE F.FARCL#='" + tClient + "'"
                    sSQL += " ORDER BY F.FANAME"
                    DataGridView_Facilities.Visible = rkutils.SQL_READ_DATAGRID(DataGridView_Facilities, "DB2", "*", DB2SQLConnectionString, DB2SQLuser, sSQL)
                Case Else
                    sSQL += " FROM iTeleCollect.dbo.ClientsFacility F"
                    sSQL += " WHERE F.CLIENT='" + tClient + "'"
                    sSQL += " ORDER BY F.FACILITYNAME"
                    DataGridView_Facilities.Visible = rkutils.SQL_READ_DATAGRID(DataGridView_Facilities, "MSSQL", "*", msSQLConnectionString, msSQLuser, sSQL)
            End Select
            MsgStatus("Facility_Loaded", True)
        Catch ex As Exception
            MsgError("Facility_Load", ex.ToString)
        End Try
    End Sub

    Private Sub Rules_Load(ByVal tClient As String, ByVal tTOB As String)
        Try
            sSQL = "SELECT *"
            Select Case sSITE
                Case "AnnuityOne"
                    sSQL += " FROM ROIDATA.RRULEP"
                    sSQL += " WHERE RRCL#='" + tClient + "'"
                    If rkutils.SQL_READ_DATAGRID(DataGridView_RULES, "DB2", "*", DB2SQLConnectionString, DB2SQLuser, sSQL) Then DataGridView_RULES.Visible = True
                Case Else
                    sSQL += " FROM iTeleCollect.dbo.Rules"
                    sSQL += " WHERE CLIENT='" + tClient + "'"
                    If rkutils.SQL_READ_DATAGRID(DataGridView_RULES, "MSSQL", "*", msSQLConnectionString, msSQLuser, sSQL) Then DataGridView_RULES.Visible = True
            End Select
            MsgStatus("Rules_Loaded", True)
        Catch ex As Exception
            MsgError("Rules_Load", ex.ToString)
        End Try
    End Sub

    Private Sub Ghosts_Load(ByVal tClient As String)
        Try
            sSQL = "SELECT * FROM " + sDBO + ".dbo.letter_ghosts"
            sSQL += rkutils.WhereAnd(sSQL, "Client='*' or Client='" + tClient + "'")
            sSQL += " ORDER BY LName, FName, MName"
            rkutils.SQL_READ_DATAGRID(DataGridView_Ghosts, "MSSQL", "*", msSQLConnectionString, msSQLuser, sSQL)
            DataGridView_Ghosts.Visible = True
            MsgStatus("Ghosts_Loaded", True)
        Catch ex As Exception
            MsgError("Ghosts_Load", ex.ToString)
        End Try
    End Sub

    'Private Sub Diligence_Load(ByVal tClient As String, ByVal tTOB As String)
    '    Try
    '        sSQL = "SELECT *"
    '        Select Case sSITE
    '            Case "AnnuityOne"
    '                sSQL += " FROM ROIDATA.DILIGENCE"
    '                sSQL += " WHERE DICL#='" + tClient + "'"
    '                DataGridView_Diligence.Visible = rkutils.SQL_READ_DATAGRID(DataGridView_Diligence, "DB2", "*", DB2SQLConnectionString, DB2SQLuser, sSQL)
    '            Case Else
    '                sSQL += " FROM iTeleCollect.dbo.Diligence"
    '                sSQL += " WHERE CLIENT='" + tClient + "'"
    '                DataGridView_Diligence.Visible = rkutils.SQL_READ_DATAGRID(DataGridView_Diligence, "MSSQL", "*", msSQLConnectionString, msSQLuser, sSQL)
    '        End Select
    '        MsgStatus("Diligence_Loaded", True)
    '    Catch ex As Exception
    '        MsgError("Diligence_Load", ex.ToString)
    '    End Try
    'End Sub

    Private Sub Button2_Click_1(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        FTP_put()
    End Sub

    Private Function FileNameRPT() As String
        Try
            If Me.CheckBox_Update.Checked = True Then
                Return dir_LETTERS + rkutils.STR_format("TODAY", "ccyymmdd") + "_" + rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)) + ".RPT"
            End If
            Return dir_TEST + rkutils.STR_format("TODAY", "ccyymmdd") + "_" + rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)) + ".RPT"
        Catch ex As Exception
            MsgError("FileNameRPT", ex.ToString)
        End Try
        Return ""
    End Function

    Private Sub CountersClear()
        Try
            '****************************
            Label_Printed.Text = "0"
            Label_PrintedM.Text = "0"
            Label_BadAddress.Text = "0"
            ListBox_Letters.Items.Clear()
            ListBox_Printed.Items.Clear()
            ListBox_Noted.Items.Clear()
        Catch ex As Exception
            MsgError("CountersClear()", ex.ToString)
        End Try
    End Sub

    Private Sub Label_HostSite_Click(sender As System.Object, e As System.EventArgs) Handles Label_HostSite.Click
        Try
            Dim myProcess As New Process
            myProcess.StartInfo.UseShellExecute = True
            myProcess.StartInfo.FileName = "https://secure.dantomsystems.com/MemberLogin.aspx"
            myProcess.Start()
        Catch ex As Exception
            MsgError("Label_HostSite_Click", ex.ToString)
        End Try
    End Sub

    Private Sub eLetters_Resize(sender As Object, e As System.EventArgs) Handles Me.Resize
        DataGridView_Select.Width = Me.Width - (DataGridView_Select.Left * 3)
        DataGridView_Multi.Width = Me.Width - (DataGridView_Multi.Left * 3)
    End Sub

    Private Sub Label_LetterFile_Click(sender As System.Object, e As System.EventArgs) Handles Label_LetterFile.Click
        Try
            Dim myProcess As New Process
            myProcess.StartInfo.UseShellExecute = True
            myProcess.StartInfo.FileName = Label_LetterFile.Text
            myProcess.Start()
        Catch ex As Exception
            MsgError("Label_LetterFile_Click", ex.ToString)
        End Try
    End Sub

    Private Sub Label_SummaryFile_Click(sender As System.Object, e As System.EventArgs) Handles Label_SummaryFile.Click
        Try
            Dim myProcess As New Process
            myProcess.StartInfo.UseShellExecute = True
            myProcess.StartInfo.FileName = Label_SummaryFile.Text
            myProcess.Start()
        Catch ex As Exception
            MsgError("Label_SummaryFile_Click", ex.ToString)
        End Try
    End Sub

    Private Function MatchedCheck(ByVal iLine As Integer, ByVal imLine As Integer, ByVal sSysAccountMatched As String)
        Try
            Select Case sSITE
                Case "AnnuityOne"
                    Return AnnuityOne_MatchedCheck(iLine, imLine, sSysAccountMatched)
                Case Else
                    'Return iTCS_MatchedCheck(iLine, imLine, sSysAccountMatched)
            End Select
        Catch ex As Exception
            MsgError("MatchedCheck", ex.ToString)
        End Try
        Return False
    End Function

    Private Sub DataGridView_Clients_MouseDoubleClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles DataGridView_Clients.MouseDoubleClick
        AccountsLoadForClient()
    End Sub

    Private Sub iTCS_letter_sent(ByVal bMatched As Boolean, ByVal iRow As Integer, ByVal tLetter As String, ByVal sBalanceSent As String)
        Try
            '****************************************************************
            '* 2012-08-21 RFK: UPDATE
            '* 2014-02-07 RFK: REMOVED FROM THIS CODE
        Catch ex As Exception
            MsgError("iTCS_letter_sent", ex.ToString)
        End Try
    End Sub

    Private Function State_Block(ByVal tLetter As String, ByVal tAllowLetters As String, ByVal tSTATE As String) As Boolean
        Try
            '*********************************************************
            '* 2012-10-02 RFK: Check Rules
            '* 2014-08-04 RFK: STATEBLOCKING JOINED AT RACCTP LEVEL
            Select Case rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text))
                Case "C"    'Collections
                    If tAllowLetters = "N" Then Return True
                Case Else
                    Return False
            End Select
            Return False
        Catch ex As Exception
            MsgError("State_Block", ex.ToString)
        End Try
        Return False
    End Function

    Private Sub AccountInitSettings()
        Try
            '******************************************************************
            Select Case sSITE
                Case "AnnuityOne"
                    Label_TCodes.Text = "RARSTA"
                    Label_FinClass.Text = "RAPAYR"
                Case Else
                    Label_TCodes.Text = "TCodes"
                    Label_FinClass.Text = "FinClass"
            End Select
            '******************************************************************
            CountersClear()
            Letter_Types_Load(rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "TOB", Val(Label_ClientRow.Text)))
            Facility_Load(rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "TOB", Val(Label_ClientRow.Text)))
            'Diligence_Load(rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "TOB", Val(Label_ClientRow.Text)))
            Rules_Load(rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "TOB", Val(Label_ClientRow.Text)))
            '******************************************************************
            '* 2021-11-26 RFK:
            Select Case rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTTYPE", Val(Label_ClientRow.Text))
                Case "C"    'Collections
                    '**********************************************************
                    '* 2021-11-26 RFK: RegF letter type MUST BE 10
                    If rkutils.DataGridView_Contains(DataGridView_Letter_Types, "LTYPE", "10") >= 0 Then
                        Panel_RegF.Visible = True
                    Else
                        Panel_RegF.Visible = False
                    End If
                Case "S"    'Self Pay
                    '**********************************************************
                    Panel_RegF.Visible = False
                Case Else
                    '**********************************************************
                    Panel_RegF.Visible = False
            End Select
            '******************************************************************
            If rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "eLettersGhosts", Val(Label_ClientRow.Text)) = "Y" Then
                rkutils.ComboBox_SetValue(ComboBox_Ghosts, "Yes")
                Ghosts_Load(rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)))
            Else
                rkutils.ComboBox_SetValue(ComboBox_Ghosts, "No")
            End If
            '**************************************************************************************
            TextBox_Processor.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "eLettersProcessor", Val(Label_ClientRow.Text))
            TextBox_MaxLetters.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "eLettersMaxPerDay", Val(Label_ClientRow.Text))
            TextBox_MaxAccounts.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "eLettersMaxAccounts", Val(Label_ClientRow.Text))
            TextBox_MaxBalance.Text = ""
            '**************************************************************************************
            '* 2016-08-19 RFK: look in ROIDATA.Diligence instead 
            TextBox_DOSminDays.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "eLettersDOSMin", Val(Label_ClientRow.Text))
            TextBox_DOSmaxDays.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "eLettersDOSMax", Val(Label_ClientRow.Text))
            '**************************************************************************************
            TextBox_TCodes.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "eLettersTCodes", Val(Label_ClientRow.Text))
            TextBox_FinClass.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "eLettersFinClass", Val(Label_ClientRow.Text))
            TextBox_SummaryEMail.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "Letters_EMAIL", Val(Label_ClientRow.Text))
            rkutils.ComboBox_SetValue(ComboBox_Facility, rkutils.STR_BREAK(rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "eLettersFacilities", Val(Label_ClientRow.Text)), 1).Trim)
            '**************************************************************************************
            TextBox_Facility.Text = rkutils.STR_BREAK(rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "eLettersFacilities", Val(Label_ClientRow.Text)), 2).Trim
            If TextBox_Facility.Text = "=" Then TextBox_Facility.Text = ""
            '**************************************************************************************
            '* 2022-02-10 RFK:
            rkutils.ComboBox_SetValue(ComboBox_DOP, rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "eLettersDOPcheck", Val(Label_ClientRow.Text)))
            TextBox_DOP_DATE.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "eLettersDOPdate", Val(Label_ClientRow.Text))
            '**************************************************************************************
            '* 2014-08-05 RFK: Total Active Accounts in Client
            sSQL = "SELECT COUNT(*) AS tCOUNT"
            Select Case rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "ELETTERS", Val(Label_ClientRow.Text))
                Case "t"
                    sSQL += " FROM ROITEST.RACCTP A"
                Case Else
                    sSQL += " FROM ROIDATA.RACCTP A"
            End Select
            sSQL += " WHERE A.RACL#='" + rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)) + "'"
            sSQL += " AND A.RACLOS<>'C'"
            Label_ClientTotalAccounts.Text = rkutils.SQL_READ_FIELD(DataGridView3, "DB2", "TCOUNT", DB2SQLConnectionString, DB2SQLuser, sSQL)
            '******************************************************************
            If Me.CheckBox_Update.Checked = True Then
                Label_SummaryFile.Text = dir_LETTERS + rkutils.STR_format("TODAY", "ccyymmdd") + "_" + rkutils.STR_format("TODAY", "HH") + "_LETTERS.TXT"
            Else
                Label_SummaryFile.Text = dir_TEST + rkutils.STR_format("TODAY", "ccyymmdd") + "_" + rkutils.STR_format("TODAY", "HH") + "_LETTERS.TXT"
            End If
            '******************************************************************
        Catch ex As Exception
            MsgError("AccountInitSettings", ex.ToString)
        End Try
    End Sub

    Private Sub Button_SaveMaxLetters_Click(sender As System.Object, e As System.EventArgs) Handles Button_SaveMaxLetters.Click
        Try
            '******************************************************************
            Dim tClientName As String = rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text))
            If tClientName.Length > 0 Then
                Dim tSQL As String = "UPDATE " + sDBO + ".dbo.clients"
                tSQL += " SET eLettersMaxPerDay='" + STR_TRIM(TextBox_MaxLetters.Text, 8) + "'"
                tSQL += ",eLettersMaxAccounts='" + STR_TRIM(TextBox_MaxAccounts.Text, 8) + "'"
                tSQL += ",eLettersProcessor='" + TextBox_Processor.Text.Trim + "'"
                tSQL += ",eLettersTCodes='" + TextBox_TCodes.Text.Trim + "'"
                tSQL += ",eLettersFinClass='" + TextBox_FinClass.Text.Trim + "'"
                tSQL += ",eLettersDOPcheck='" + STR_LEFT(ComboBox_DOP.Text, 2).Trim + "'"
                tSQL += ",eLettersDOPdate='" + STR_LEFT(TextBox_DOP_DATE.Text.Trim, 10) + "'"
                tSQL += ",eLettersFacilities='" + STR_LEFT(ComboBox_Facility.Text, 2).PadRight(2) + " " + TextBox_Facility.Text + "'"
                tSQL += ",eLettersGhosts='" + rkutils.STR_TRIM(ComboBox_Ghosts.Text, 1) + "'"
                tSQL += ",Letters_EMail='" + TextBox_SummaryEMail.Text.Trim + "'"
                '2016-08-19 RFK: moved to ROIDATA.DILIGENCE tSQL += ",eLettersBalanceMax='" + TextBox_MaxBalance.Text.Trim + "'"
                '2016-08-19 RFK: moved to ROIDATA.DILIGENCE tSQL += ",eLettersBalanceMin='" + TextBox_MinBalance.Text.Trim + "'"
                tSQL += " WHERE ClientName='" + tClientName + "'"
                rkutils.DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, tSQL)
                MsgStatus(tSQL, CheckBox_DEBUG.Checked)
                '**************************************************************
                'tSQL = "UPDATE ROIDATA.DILIGENCE"
                'tSQL += " SET MINLETTERBAL='" + TextBox_MinBalance.Text.Trim + "'"
                'tSQL += " WHERE DICL#='" + tClientName + "'"
                ''rkutils.DB_COMMAND("DB2", DB2SQLConnectionString, DB2SQLuser, tSQL)
                'MsgStatus(tSQL, CheckBox_DEBUG.Checked)
                '**************************************************************
                MsgStatus("Saved " + tClientName + " settings", CheckBox_DEBUG.Checked)
                ClientsInit()
            End If
        Catch ex As Exception
            MsgError("Button_SaveMaxLetters_Click", ex.ToString)
        End Try
    End Sub

    Private Sub Button_ClientSave_Click(sender As System.Object, e As System.EventArgs) Handles Button_ClientSave.Click
        Try
            Dim tClientName As String = rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text))
            If tClientName.Length > 0 Then
                Dim tSQL As String = "UPDATE " + sDBO + ".dbo.clients"
                tSQL += " SET eLetters='" + STR_TRIM(ComboBox_ClientActive.Text, 1) + "'"
                tSQL += " WHERE ClientName='" + tClientName + "'"
                rkutils.DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, tSQL)
                MsgStatus(tSQL, True)
                MsgStatus("Saved " + tClientName + " settings", True)
                ClientsInit()
            End If
        Catch ex As Exception
            MsgError("Button_SaveMaxLetters_Click", ex.ToString)
        End Try
    End Sub

    Private Sub ComboBox_Processor_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox_Processor.SelectedIndexChanged
        Try
            ClientsInit()
        Catch ex As Exception
            MsgError("ComboBox_Processor_SelectedIndexChanged", ex.ToString)
        End Try
    End Sub

    Private Sub Button_Load_Click(sender As System.Object, e As System.EventArgs) Handles Button_Load.Click
        Try
            AccountsLoadForClient()
        Catch ex As Exception
            MsgError("Button_Load_Click", ex.ToString)
        End Try
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Try
        Catch ex As Exception
            MsgError("TestButton", ex.ToString)
        End Try

    End Sub

    Private Sub EMail_Summary()
        Try
            If Val(Label_NumberAccounts.Text) <= 0 Then Exit Sub

            Dim tClient As String = rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text))

            Dim tMSG As String = tClient + " " + STR_format("TODAY", "mm/dd/ccyy HH:MM:SS") + vbCrLf + vbCrLf
            If CheckBox_Update.Checked = False Then
                tMSG += "[TESTING ONLY NOT UPDATED]" + vbCrLf + vbCrLf
            End If
            tMSG += "Total Selected:" + Label_NumberAccounts.Text + vbCrLf
            tMSG += "Letters Sent:" + Label_Printed.Text + vbCrLf
            tMSG += "Matched:" + Label_PrintedM.Text + vbCrLf
            'tMSG += "State Blocked:" + Label_StateBlocked.Text + vbCrLf
            tMSG += "Bad Address:" + Label_BadAddress.Text + vbCrLf
            tMSG += "$ Blocked:" + "0" + vbCrLf
            'tMSG += "Future:" + Label_Future.Text + vbCrLf
            'tMSG += "" + vbCrLf
            If Me.ListBox_Letters.Items.Count > 0 Then
                'tMSG += "Letter breakdown" + vbCrLf
                'tMSG += "----------------" + vbCrLf
                tMSG += vbCrLf
                ListBox_Letters.Sorted = True
                For i1 = 0 To Me.ListBox_Letters.Items.Count - 1
                    tMSG += tClient + "" + Me.ListBox_Letters.Items(i1).ToString
                    If ListBox_Letters.Items(i1).ToString.Contains("ERR_") Then
                        tMSG += vbTab + dir_REPORTS + "LETTERS\" + rkutils.STR_format("TODAY", "ccyymmdd") + "_" + rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)) + "_" + STR_BREAK(ListBox_Letters.Items(i1).ToString, 1).Trim + ".XLS"
                    End If
                    tMSG += vbCrLf
                    '*************************************************************************************
                    '* 2014-09-25 RFK:
                    If STR_LEFT(ListBox_Letters.Items(i1).ToString, 3) = "ERR" Then
                        'DONT detail the errors
                    Else
                        If CheckBox_Update.Checked Then
                            sSQL = "INSERT INTO " + sDBO + ".dbo.LetterVendorFileDetail"
                            sSQL += " (RevMDFile"
                            sSQL += ",Client"
                            sSQL += ",LetterCode"
                            sSQL += ",LetterCount"
                            sSQL += ")"
                            sSQL += " VALUES("
                            sSQL += "'" + Path.GetFileName(Label_LetterFile.Text) + "'"
                            sSQL += ",'" + tClient + "'"
                            sSQL += ",'" + STR_LEFT(ListBox_Letters.Items(i1).ToString, 3) + "'"
                            sSQL += ",'" + STR_BREAK(ListBox_Letters.Items(i1).ToString, 2) + "'"
                            sSQL += ")"
                            rkutils.DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, sSQL)
                        End If
                    End If
                    '*************************************************************************************
                Next
            End If
            If Label_SummaryFile.Text.Length > 0 Then
                tMSG += "----------------" + vbCrLf
                File.AppendAllText(Label_SummaryFile.Text, tMSG)
            End If
            If CheckBox_Update.Checked = True Then
                Dim sEmailTo As String = rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "LETTERS_EMAIL", Val(Label_ClientRow.Text))
                If sEmailTo.Contains("@") Then
                    Select Case sSITE
                        Case "AnnuityOne"
                            rkutils.EMAILIT(msSQLConnectionString, msSQLuser, "eLetters@AnnuityHealth.com", "eLetters", sEmailTo, "Letter Results" + tClient, Me.Text, "eLetter Summary", tMSG, "", "")
                        Case Else
                            rkutils.EMAILIT(msSQLConnectionString, msSQLuser, "DoNotReply@FosterTech.net", "eLetters", sEmailTo, "Letter Results" + tClient, Me.Text, "eLetter Summary", tMSG, "", "")
                    End Select
                End If
            Else
                rkutils.EMAILIT(msSQLConnectionString, msSQLuser, "eLetters@AnnuityHealth.com", "eLetters", "letters@AnnuityHealth.com", "Letter Results" + tClient, Me.Text, "eLetter Summary", tMSG, "", "")
            End If
            MsgStatus(tMSG, True)
        Catch ex As Exception
            MsgError("EMail_Summary", ex.ToString)
        End Try
    End Sub

    Private Function CalculateCAPB(ByVal sPACB As String, ByVal sFacilityEqual As String, ByVal sFacility As String, ByVal sMatchedAccount As String) As String
        Try
            '**************************************************************************************
            '* 2014-07-11 RFK: moved DIM to top for global
            dBalance = 0
            sLocxs = ""
            sAccountSave = ""
            sSQL = ""
            sAmount = ""
            sAmountNew = ""
            sRAFACL = ""
            CalculateCAPB_iLocxRow = 0

            Select Case sSITE
                Case "AnnuityOne"
                    CalculateCAPB_iLocxRow = DataGridView_Contains(DataGridView_Multi, "RAMLOCX", sMatchedAccount)
                    '******************************************************************************
                    If CalculateCAPB_iLocxRow >= 0 Then
                        sAccountSave = ReadField(DataGridView_Multi, "RAMLOCX", CalculateCAPB_iLocxRow)
                        sLocxs = ""
                        Do While CalculateCAPB_iLocxRow < DataGridView_Multi.Rows.Count - 1 And Label_RUNNING.Text = "Running" And sAccountSave = sMatchedAccount
                            sRAFACL = ReadField(DataGridView_Multi, "RAFACL", CalculateCAPB_iLocxRow)
                            If sFacilityEqual = "=" Then
                                If sFacility = sRAFACL Then
                                    If sLocxs.Length > 0 Then sLocxs += ","
                                    sLocxs += ReadField(DataGridView_Multi, "RALOCX", CalculateCAPB_iLocxRow)
                                    '************************
                                    Select Case sPACB
                                        Case "B"
                                            dBalance += Val(ReadField(DataGridView_Multi, "RABALD", CalculateCAPB_iLocxRow))
                                            sAmount = Trim(Str(dBalance))
                                        Case "C"
                                            dBalance += Val(ReadField(DataGridView_Multi, "RAOBAL", CalculateCAPB_iLocxRow))
                                            sAmount = Trim(Str(dBalance))
                                        Case "A"
                                            dBalance += Val(ReadField(DataGridView_Multi, "RATOTA", CalculateCAPB_iLocxRow))
                                        Case "P"
                                            dBalance += Val(ReadField(DataGridView_Multi, "RATOTP", CalculateCAPB_iLocxRow))
                                    End Select
                                    '************************
                                End If
                            Else
                                If sFacility <> sRAFACL Then
                                    If sLocxs.Length > 0 Then sLocxs += ","
                                    sLocxs += ReadField(DataGridView_Multi, "RALOCX", CalculateCAPB_iLocxRow)
                                    '************************
                                    Select Case sPACB
                                        Case "B"
                                            dBalance += Val(ReadField(DataGridView_Multi, "RABALD", CalculateCAPB_iLocxRow))
                                            sAmount = Trim(Str(dBalance))
                                        Case "C"
                                            dBalance += Val(ReadField(DataGridView_Multi, "RAOBAL", CalculateCAPB_iLocxRow))
                                            sAmount = Trim(Str(dBalance))
                                        Case "A"
                                            dBalance += Val(ReadField(DataGridView_Multi, "RATOTA", CalculateCAPB_iLocxRow))
                                        Case "P"
                                            dBalance += Val(ReadField(DataGridView_Multi, "RATOTP", CalculateCAPB_iLocxRow))
                                    End Select
                                    '************************
                                End If
                            End If
                            '**********************************************************************
                            CalculateCAPB_iLocxRow += 1
                            sAccountSave = ReadField(DataGridView_Multi, "RAMLOCX", CalculateCAPB_iLocxRow)
                        Loop
                        If sLocxs.Length > 0 Then
                            '**********************************************************************
                            Select Case sPACB
                                Case "B", "C"
                                    Return sAmount
                            End Select
                            '**********************************************************************
                            sSQL = "SELECT SUM(A.CXAMT) AS TSUM FROM"
                            sSQL += "("
                            sSQL += "SELECT DISTINCT LEFT(CXTYPE,1),CXDATE,CXAMT,CXDESC"
                            '**********************************************************************
                            '* 2014-06-19 RFK:
                            Select Case rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "ELETTERS", Val(Label_ClientRow.Text))
                                Case "t"
                                    sSQL += " FROM ROITEST.CXTRNP"
                                Case Else
                                    sSQL += " FROM ROIDATA.CXTRNP"
                            End Select
                            sSQL += " WHERE CXLOCX IN (" + sLocxs + ")"
                            sSQL += " AND LEFT(CXTYPE,1)='" + sPACB + "'"
                            sSQL += " GROUP BY LEFT(CXTYPE,1),CXDATE,CXDESC,CXAMT"
                            sSQL += ") A"
                            If CheckBox_DEBUG.Checked Then MsgStatus(sSQL, False)
                            sAmount = SQL_READ_FIELD(DataGridView3, "DB2", "TSUM", DB2SQLConnectionString, DB2SQLuser, sSQL)
                            '**********************************************************************
                            Select Case sPACB
                                Case "A"
                                    dBalance += Val(sAmount)
                                    sAmount = Trim(Str(dBalance))
                                Case "P"
                                    dBalance += Val(sAmount)
                                    sAmount = Trim(Str(dBalance))
                            End Select
                            '**********************************************************************
                        End If
                    End If
                    Return sAmount
                Case "iTeleCollect"
                    CalculateCAPB_iLocxRow = DataGridView_Contains(DataGridView_Multi, "SYSACCOUNTMATCHED", sMatchedAccount)
                    If CalculateCAPB_iLocxRow >= 0 Then
                        sAccountSave = ReadField(DataGridView_Multi, "SYSACCOUNTMATCHED", CalculateCAPB_iLocxRow)
                        Do While CalculateCAPB_iLocxRow < DataGridView_Multi.Rows.Count - 1 And Label_RUNNING.Text = "Running" And sAccountSave = sMatchedAccount
                            dBalance += Val(ReadField(DataGridView_Multi, "PatientBalance", CalculateCAPB_iLocxRow))
                            CalculateCAPB_iLocxRow += 1
                            sAccountSave = ReadField(DataGridView_Multi, "SYSACCOUNTMATCHED", CalculateCAPB_iLocxRow)
                        Loop
                    End If
            End Select
            Return dBalance.ToString
        Catch ex As Exception
            MsgError("CalculateCAPB", ex.ToString)
            Return ""
        End Try
    End Function

    Protected Sub FoxProLoad(sModule As String, ByVal sPath As String, ByVal sFileName As String, sRecordFrom As String, sRecordTo As String, ByVal AsTable As Boolean, ByRef DT As DataTable)
        Try
            '**************************************************************************************
            '* 2015-06-01 RFK:
            '* 2015-06-15 RFK:
            sSQL = "SELECT COUNT(*) AS TCOUNT FROM "
            sSQL += sPath + sFileName
            SQL_READ_DATAGRID(DataGridView_Select, "FOXPRO", "*", sPath, "", sSQL)
            Label_NumberClients.Text = ReadField(DataGridView_Select, "TCOUNT", 0)
            'Label_NumberOfAccounts.Text = "0"
            MsgStatus(Label_NumberClients.Text, True)
            '**************************************************************************************
            '* 2015-06-01 RFK:
            sSQL = "SELECT PAT_ACNT, LRESPONSE, LCONTDATE FROM "
            sSQL += sPath + sFileName
            Select Case sModule
                Case "LETTERS"
                    sSQL += rkutils.WhereAnd(sSQL, "LEN(TRIM(PAT_ACNT)) > 0")
                    sSQL += rkutils.WhereAnd(sSQL, "LRESPONSE='" + sRecordFrom + "'")
                    Select Case sRecordTo
                        Case "TODAY"
                            sSQL += rkutils.WhereAnd(sSQL, "LCONTDATE = CTOD('" + rkutils.STR_format("TODAY", "mm/dd/ccyy") + "')")
                            'Case "YESTERDAY"
                            '    sSQL += rkutils.WhereAnd(sSQL, "LCONTDATE = CTOD('" + rkutils.STR_format("TODAY", "mm/dd/ccyy") + "'")
                        Case Else
                            MsgStatus(sRecordTo, True)
                    End Select
            End Select
            MsgStatus(sSQL, False)
            If AsTable = True Then
                rkutils.SQL_READ_DATATABLE(DT, "FOXPRO", "*", sPath, "", sSQL)
                Label_AccountsRemaining.Text = DT.Rows.Count - 1.ToString.Trim
            Else
                DataGridView_Select.Visible = SQL_READ_DATAGRID(DataGridView_Select, "FOXPRO", "*", sPath, "", sSQL)
                Label_AccountsRemaining.Text = (DataGridView_Select.Rows.Count - 1).ToString
            End If
            MsgStatus("Loaded:" + Label_AccountsRemaining.Text, True)
        Catch ex As Exception
            MsgError("FoxProConvert", ex.ToString)
        End Try
    End Sub

    Protected Function SQLtable_CreateFrom_DataGridView(ByVal sDB As String, ByVal sDatabase As String, ByVal sTable As String, ByVal dGV As DataGridView) As Boolean
        Try
            '******************************************************************
            '* 2015-07-29 RFK:
            '* 2015-08-12 RFK: converted to DataGridView from DataTable
            SQLtableDrop(sDB, sDatabase, sTable)
            '******************************************************************
            '* 2015-07-29 RFK:
            MsgStatus("Creating " + sTable, True)
            '******************************************************************
            Select Case sDB
                Case "DB2"
                    sSQL = "CREATE TABLE " + sDatabase + "." + sTable + "("
                Case "MSSQL"
                    sSQL = "USE " + sDatabase + vbCr
                    sSQL += "SET ANSI_NULLS ON" + vbCr
                    sSQL += "SET QUOTED_IDENTIFIER ON" + vbCr
                    sSQL += "SET ANSI_PADDING ON" + vbCr
                    sSQL += "CREATE TABLE dbo." + sTable + "("
            End Select
            '******************************************************************
            For i2 = 0 To dGV.ColumnCount - 1
                '******************************************************************
                Select Case sDB
                    Case "DB2"
                        sSQL += dGV.Columns(i2).Name.Trim
                        sSQL += " varchar(100)"
                        If i2 < dGV.ColumnCount - 1 Then sSQL += ","
                        sSQL += vbCr
                    Case "MSSQL"
                        sSQL += "[" + dGV.Columns(i2).Name.Trim + "]"
                        sSQL += " [varchar]"
                        sSQL += " (100)"
                        sSQL += " NULL"
                        If i2 < dGV.ColumnCount - 1 Then sSQL += ","
                        sSQL += vbCr
                End Select
            Next
            sSQL += ")"
            MsgStatus(sSQL, False)
            Select Case sDB
                Case "DB2"
                    DB_COMMAND("DB2", DB2SQLConnectionString, DB2SQLuser, sSQL)
                Case "MSSQL"
                    DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, sSQL)
            End Select
            Return True
        Catch ex As Exception
            MsgError("TableCreateTempAccount()", ex.ToString)
        End Try
        Return False
    End Function

    Private Sub TempTable_Insert_DataGridView(ByVal sDB As String, ByVal sDatabase As String, ByVal sTable As String, ByVal dGV As DataGridView)
        Try
            '******************************************************************
            Dim sSQLinsert As String = ""
            '* 2015-08-12 RFK:
            If SQLtable_CreateFrom_DataGridView(sDB, sDatabase, sTable, dGV) Then
                '**************************************************************
                '* 2015-08-12 RFK:
                MsgStatus("Inserting into " + sTable, True)
                Select Case sDB
                    Case "DB2"
                        sSQLinsert = "INSERT INTO " + sDatabase + "." + sTable + " ("
                    Case "MSSQL"
                        sSQLinsert = "INSERT INTO " + sDatabase + ".dbo." + sTable + " ("
                End Select
                '**************************************************************
                For i2 = 0 To dGV.ColumnCount - 1
                    Label_AccountsRemaining.Text = Str(i2)
                    Application.DoEvents()
                    '**********************************************************
                    If i2 > 0 Then sSQLinsert += ","
                    sSQLinsert += dGV.Columns(i2).Name.Trim
                Next
                sSQLinsert += ")"
                '**************************************************************
                For i1 = 0 To dGV.RowCount - 1
                    Label_AccountsRemaining.Text = Str(i1)
                    Application.DoEvents()
                    '**********************************************************
                    If dGV.Item(0, i1).Value IsNot Nothing Then
                        '******************************************************
                        sSQL = sSQLinsert + " VALUES("
                        For i2 = 0 To dGV.ColumnCount - 1
                            If dGV.Item(i2, i1).Value IsNot Nothing Then
                                If i2 > 0 Then sSQL += ","
                                sSQL += "'" + dGV.Item(i2, i1).Value.ToString.Trim.Replace("'", "-") + "'"
                            End If
                        Next
                        sSQL += ")"
                        Select Case sDB
                            Case "DB2"
                                DB_COMMAND("DB2", DB2SQLConnectionString, DB2SQLuser, sSQL)
                            Case "MSSQL"
                                DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, sSQL)
                        End Select
                    End If
                Next
                '**********************************************************
            End If
            MsgStatus("Completed Insert into" + sTable, True)
        Catch ex As Exception
            Msg_Error("TempTable_InsertDataGridView", ex.ToString)
        End Try
    End Sub

    Private Sub TempTable_Insert_DataTable(ByVal sDB As String, ByVal sDatabase As String, ByVal sTable As String, ByVal dT As DataTable)
        Try
            '******************************************************************
            Dim sSQLinsert As String = ""
            MsgStatus("Need To Convert", True)
            '* 2015-08-12 RFK:
            'If SQLtable_CreateFrom_DataGridView(sDB, sDatabase, sTable, dGV) Then
            '    '**************************************************************
            '    '* 2015-08-12 RFK:
            '    MsgStatus("Inserting into " + sTable, True)
            '    Select Case sDB
            '        Case "DB2"
            '            sSQLinsert = "INSERT INTO " + sDatabase + "." + sTable + " ("
            '        Case "MSSQL"
            '            sSQLinsert = "INSERT INTO " + sDatabase + ".dbo." + sTable + " ("
            '    End Select
            '    '**************************************************************
            '    For i2 = 0 To dGV.ColumnCount - 1
            '        Label_AccountsRemaining.Text = Str(i2)
            '        Application.DoEvents()
            '        '**********************************************************
            '        If i2 > 0 Then sSQLinsert += ","
            '        sSQLinsert += dGV.Columns(i2).Name.Trim
            '    Next
            '    sSQLinsert += ")"
            '    '**************************************************************
            '    For i1 = 0 To dGV.RowCount - 1
            '        Label_AccountsRemaining.Text = Str(i1)
            '        Application.DoEvents()
            '        '**********************************************************
            '        If dGV.Item(0, i1).Value IsNot Nothing Then
            '            '******************************************************
            '            sSQL = sSQLinsert + " VALUES("
            '            For i2 = 0 To dGV.ColumnCount - 1
            '                If dGV.Item(i2, i1).Value IsNot Nothing Then
            '                    If i2 > 0 Then sSQL += ","
            '                    sSQL += "'" + dGV.Item(i2, i1).Value.ToString.Trim.Replace("'", "-") + "'"
            '                End If
            '            Next
            '            sSQL += ")"
            '            Select Case sDB
            '                Case "DB2"
            '                    DB_COMMAND("DB2", DB2SQLConnectionString, DB2SQLuser, sSQL)
            '                Case "MSSQL"
            '                    DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, sSQL)
            '            End Select
            '        End If
            '    Next
            '    '**********************************************************
            'End If
            MsgStatus("Completed Insert into" + sTable, True)
        Catch ex As Exception
            Msg_Error("TempTable_InsertDataGridView", ex.ToString)
        End Try
    End Sub

    Protected Function SQLtableDrop(ByVal sDB As String, ByVal sDatabase As String, ByVal sTable As String) As Boolean
        Try
            '******************************************************************
            '* 2015-07-29 RFK:
            '* 2015-08-12 RFK: converted to DataGridView from DataTable
            Select Case sDB
                Case "DB2"
                    sSQL = "DROP TABLE " + sDatabase + "." + sTable + vbCr
                    MsgStatus(sSQL, True)
                    DB_COMMAND("DB2", DB2SQLConnectionString, DB2SQLuser, sSQL)
                    Return True
                Case "MSSQL"
                    sSQL = "USE " + sDatabase + vbCr
                    sSQL += "DROP TABLE dbo." + sTable + vbCr
                    MsgStatus(sSQL, True)
                    DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, sSQL)
                    Return True
            End Select
        Catch ex As Exception
            MsgError("SQLtableDrop", ex.ToString)
        End Try
        Return False
    End Function

    Protected Sub SQLtable_CreateFrom_DataTable(ByVal sDatabase As String, ByVal sTable As String, ByVal dT As DataTable)
        Try
            '******************************************************************
            '* 2015-07-29 RFK:
            SQLtableDrop("MSSQL", sDatabase, sTable)
            '******************************************************************
            '* 2015-07-29 RFK:
            MsgStatus("Creating " + sTable, True)
            '******************************************************************
            sSQL = "USE iTeleCollect" + vbCr
            sSQL += "SET ANSI_NULLS ON" + vbCr
            sSQL += "SET QUOTED_IDENTIFIER ON" + vbCr
            sSQL += "SET ANSI_PADDING ON" + vbCr
            sSQL += "CREATE TABLE dbo." + sTable + "("
            For Each col As DataColumn In dT.Columns
                sSQL += "[" + col.ColumnName.ToString.Trim + "]"
                sSQL += " [varchar]"
                sSQL += " (100)"
                sSQL += " NULL"
                sSQL += ","
                sSQL += vbCr
            Next
            sSQL += ")"
            MsgStatus(sSQL, True)
            DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, sSQL)
        Catch ex As Exception
            MsgError("TableCreateTempAccount()", ex.ToString)
        End Try
    End Sub

    Private Sub ClientsInit()
        Try
            Dim tSEL As String = "SELECT"
            Select Case ComboBox_MatchType.Text
                Case "Active Clients", "Inactive Clients", "Testing Clients"
                    tSEL += " DISTINCT C.ClientName,C.FriendlyName,C.LetterName"
                    tSEL += ",VMB.CallBackNumber"
                    tSEL += ",C.LettersStartDate,C.LettersRunDate,C.LettersRunTotal"
                    tSEL += ",C.eLettersProcessor,C.ClientType"
                    tSEL += ",CF.LetterMatchType,CF.LetterType"
                    tSEL += ",C.eLetters,C.eLettersMaxPerDay,C.eLettersMaxAccounts"
                    tSEL += ",C.eLettersDOPcheck, C.eLettersDOPdate"
                    tSEL += ",C.eLettersFutureLetterDates,C.eLettersFacilities,C.eLettersGhosts"
                    tSEL += ",C.eLettersTCodes, C.eLettersFinClass, C.eLettersNameToUse,C.eLettersLockBox"
                    tSEL += ",C.Letters_EMail"
                    tSEL += ",C.AdverseTU,C.AdverseMinBal"
                    tSEL += ",C.PPL_Broken_Status,C.PPL_Broken_MinimumLetterAmount"
                    tSEL += ",CM.MatchBy,CM.MatchStart,CM.MatchComplete"
                Case "By Facility"
                    tSEL += " CF.Client,CF.Facility,CF.LetterMatchType,CF.LetterType"
                Case Else
                    Exit Sub
            End Select
            '****************************************
            tSEL += " FROM " + sDBO + ".dbo.Clients C"
            tSEL += " LEFT JOIN " + sDBO + ".dbo.ClientsFacility CF ON C.ClientName=CF.Client"
            tSEL += " LEFT JOIN " + sDBO + ".dbo.ClientsMatching CM ON C.ClientName=CM.ClientName AND CM.Facility='*'"  '2013-08-28 RFK: checking to see if Client Matched today
            tSEL += " LEFT JOIN " + sDBO + ".dbo.ClientsVMB VMB ON C.ClientName=VMB.ClientName"                         '2022-05-16 RFK:
            tSEL += " WHERE C.Active='Y'"
            '****************************************
            '* 2014-08-26 RFK:
            Select Case ComboBox_Processor.Text
                Case "ALL"
                    'ALL
                Case Else
                    tSEL += rkutils.WhereAnd(tSEL, "C.eLettersProcessor='" + ComboBox_Processor.Text + "'")
            End Select
            '****************************************
            Select Case ComboBox_MatchType.Text
                Case "Active Clients"
                    tSEL += rkutils.WhereAnd(tSEL, "C.eLetters='Y'")
                    tSEL += rkutils.WhereAnd(tSEL, "CF.LetterMatchType='*'")
                    tSEL += " ORDER BY C.ClientName"
                    ComboBox_SetValue(ComboBox_ClientActive, "Yes letters")
                Case "By Facility"
                    tSEL += rkutils.WhereAnd(tSEL, "CF.LetterMatchType<>'*'")
                    tSEL += " ORDER BY CF.Client,CF.Facility"
                Case "Inactive Clients"
                    tSEL += rkutils.WhereAnd(tSEL, "(C.eLetters Is Null OR (C.eLetters<>'Y' AND C.eLetters<>'T'))")
                    tSEL += rkutils.WhereAnd(tSEL, "CF.LetterMatchType='*'")
                    tSEL += " ORDER BY C.ClientName"
                    ComboBox_SetValue(ComboBox_ClientActive, "No letters")
                Case "Testing Clients"
                    tSEL += rkutils.WhereAnd(tSEL, "(C.eLetters='T' OR C.eLetters='t')")
                    tSEL += rkutils.WhereAnd(tSEL, "CF.LetterMatchType='*'")
                    tSEL += " ORDER BY C.ClientName"
                    ComboBox_SetValue(ComboBox_ClientActive, "Testing")
            End Select
            '**************************************************************************************
            MsgStatus(tSEL, False)
            DataGridView_Clients.Visible = SQL_READ_DATAGRID(DataGridView_Clients, "MSSQL", "*", msSQLConnectionString, msSQLuser, tSEL)
            If DataGridView_Clients.Rows.Count > 0 Then
                If Val(Label_ClientRow.Text) > 0 Then
                    DataGridView_Clients.CurrentCell = DataGridView_Clients.Rows(Val(Label_ClientRow.Text)).Cells(0)
                Else
                    Label_ClientRow.Text = "0"
                End If
                '**********************************************************************************
                '* 2021-07-08 RFK:
                Dim dDate As Date, dDateN As Date = Now
                Dim dDateD As Double
                For iRow = 0 To DataGridView_Clients.Rows.Count - 2
                    '******************************************************************************
                    DataGridView_Clients.Item(rkutils.DataGridView_ColumnByName(DataGridView_Clients, "ClientName"), iRow).Style.BackColor = Color.Green
                    DataGridView_Clients.Item(rkutils.DataGridView_ColumnByName(DataGridView_Clients, "ClientName"), iRow).Style.ForeColor = Color.White
                    '******************************************************************************
                    dDate = rkutils.ReadField(DataGridView_Clients, "MatchComplete", iRow)
                    If IsDate(dDate) Then
                        dDateD = DateDiff(DateInterval.Hour, dDate, dDateN)
                    Else
                        dDateD = 99
                    End If
                    '******************************************************************************
                    '* 2022-05-10 RFK: Matched within 20 HOURS
                    If dDateD > 20 Then
                        DataGridView_Clients.Item(rkutils.DataGridView_ColumnByName(DataGridView_Clients, "ClientName"), iRow).Style.BackColor = Color.Red
                        DataGridView_Clients.Item(rkutils.DataGridView_ColumnByName(DataGridView_Clients, "MatchComplete"), iRow).Style.BackColor = Color.Red
                    End If
                Next
                '**********************************************************************************
            Else
                Label_ClientRow.Text = ""
            End If
            '**************************************************************************************
        Catch ex As Exception
            MsgError("ClientsInit", ex.ToString)
        End Try
    End Sub

    Private Sub FTP_put()
        Try
            Dim tFTPput As String = ""
            Select Case ComboBox_MatchType.Text
                Case "Testing Clients"
                    Select Case tCurrentLetterVendor
                        Case "ACCUDOC"
                            tFTPput = dir_FTP + "ACCUDOC\ACCUDOC_TEST_PUT.BAT"
                        Case "APEX"
                            tFTPput = dir_FTP + "APEX\APEX_TEST_PUT.BAT"
                        Case "DANTOM", "REVSPRING"
                            tFTPput = dir_FTP + "DANTOM\DANTOM_TEST_PUT.BAT"
                        Case Else
                            MsgStatus("ERROR FTP LETTER VENDOR [" + tCurrentLetterVendor + "] NOT DEFINED", True)
                    End Select
                Case Else
                    Select Case tCurrentLetterVendor
                        Case "ACCUDOC"
                            tFTPput = dir_FTP + "ACCUDOC\ACCUDOC_PUT.BAT"
                        Case "APEX"
                            tFTPput = dir_FTP + "APEX\APEX_PUT.BAT"
                        Case "DANTOM", "REVSPRING"
                            tFTPput = dir_FTP + "DANTOM\DANTOM_PUT.BAT"
                        Case Else
                            MsgStatus("ERROR FTP LETTER VENDOR [" + tCurrentLetterVendor + "] NOT DEFINED", True)
                    End Select
            End Select
            If IS_File(tFTPput) Then
                If IS_File(Label_LetterFile.Text) Then
                    MsgStatus("FTP (" + tFTPput + ") put [" + Label_LetterFile.Text + "]", True)
                    Shell(tFTPput + " " + Label_LetterFile.Text, AppWinStyle.NormalFocus, True, 120)
                    MsgStatus("FTP COMPLETE (" + tFTPput + ") put [" + Label_LetterFile.Text + "]", True)
                    Wait(60)
                    '**********************************************************
                    '* 2017-03-07 RFK: 
                    Select Case tCurrentLetterVendor
                        Case "ACCUDOC"
                            Dim tLines As String = File.ReadAllText(dir_FTP + "ACCUDOC\ACCUDOC.LOG")
                            If tLines.Contains(Path.GetFileName(Label_LetterFile.Text)) Then
                                MsgStatus("ACCUDOC.LOG CONTAINS [" + Path.GetFileName(Label_LetterFile.Text) + "]", True)
                            Else
                                MsgStatus("ACCUDOC.LOG DOES NOT CONTAIN [" + Path.GetFileName(Label_LetterFile.Text) + "]", True)
                                Wait(10)
                                Shell(tFTPput + " " + Label_LetterFile.Text, AppWinStyle.NormalFocus, True, 60)
                                MsgStatus("FTP AGAIN (" + tFTPput + ") put [" + Label_LetterFile.Text + "]", True)
                            End If
                        Case "DANTOM", "REVSPRING"
                    End Select
                    '**********************************************************
                Else
                    MsgStatus("ERROR FILE [" + Label_LetterFile.Text + "] NOT FOUND", True)
                End If
            Else
                MsgStatus("ERROR FTP [" + tFTPput + "] NOT FOUND", True)
            End If
        Catch ex As Exception
            Msg_Error("FTP_put", ex.ToString)
        End Try
    End Sub

    Private Sub PrintHeader(ByVal tLetterVendor As String, ByVal tFileName As String)
        Try
            '************************************************
            '* 2012-06-21 RFK:
            Select Case tLetterVendor
                Case "ACCUDOC"
                    PrintLine_tLINE = "*Annuity Health" + vbCrLf
                    PrintLine_tLINE += "IT Production (630) 882-3690" + vbCrLf
                    PrintLine_tLINE += vbCrLf
                    File.AppendAllText(tFileName, PrintLine_tLINE)
                Case "DANTOM", "REVSPRING"
                    PrintLine_tLINE = "*0000CSAH AnnuityHealth" + vbCrLf
                    PrintLine_tLINE += "IO Production (855) 896-9990" + vbCrLf
                    PrintLine_tLINE += vbCrLf
                    File.AppendAllText(tFileName, PrintLine_tLINE)
                    If 1 = 2 Then
                        PrintLine_tLINE = tQuote + "LetterAccount" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "LetterCustomer" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "LetterCode" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "ClientNumber" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "ListDate" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "PatAcnt" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "GSOCSEC" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "GNAME" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "GNAMEF" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "GNAMEM" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "PAT_NAME" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "PAT_NAMEF" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "PAT_NAMEM" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "GADDR1" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "GADDR2" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "GCITY" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "GSTATE" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "GZIP" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "GZIPEXT" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "GAREA1" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "GPHONE1" + tQuote + tDelimiter
                        'PrintLine_tLINE += tQuote + "GAREA2" + tQuote + tDelimiter
                        'PrintLine_tLINE += tQuote + "GPHONE2" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "SERVDATE" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "ACNTAGE" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "ORIGINBAL" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "PNETPAY" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "PNETADJ" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "PNETCHG" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "CURRBAL" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "ATTEMPTS" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "ANSWERS" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "CONTACTS" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "LPROMDATE" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "LPROMAMT" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "LCONTDATE" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "LCONTTIME" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "LRESPONSE" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "LCONTCOM" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "FISC" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "FISC_DESCR" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "ROUTE" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "TYPE" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "TYPE_DESCR" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "ID_REC" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "DISCDATE" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "TOTPAY" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "TOTADJ" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "TOTCHG" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "FACILITY" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "FACILITY_D" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "USERNAME" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "ALL_CHARGES" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "ALL_PAYS" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "ALL_ADJS" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "ALL_PAYSADJS" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "ADMITDATE" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "GDOB" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "PDOB" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "EMPI" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "IPOP" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "MAST_ACNT" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "MATCH_BAL" + tQuote + tDelimiter
                        PrintLine_tLINE += tQuote + "EOL" + tQuote + tDelimiter
                    End If
                Case Else
                    MsgStatus("PrintHeader/Invalid Letter Vendor [" + tLetterVendor + "]", True)
            End Select
        Catch ex As Exception
            MsgError("PrintHeader", ex.ToString)
        End Try
    End Sub

    Private Sub PrintTrailer(ByVal tLetterVendor As String, ByVal tFileName As String)
        Try
            '******************************************************************
            '* 2012-06-21 RFK:
            PrintLine_tLINE = ""
            Select Case tLetterVendor
                Case "ACCUDOC"
                    PrintLine_tLINE += "MMMM" + vbCrLf
                    File.AppendAllText(tFileName, PrintLine_tLINE)
                Case "APEX"
                    '**********************************************************
                    '* 2015-07-28 RFK:
                    Select Case sSITE
                        '******************************************************
                        '* 2013-05-01 RFK:
                        Case "AnnuityOne"
                            '
                        Case "iTeleCollect"
                            PrintLine_tLINE += "EOF:" + iNumPrinted.ToString + "-records."
                            PrintLine_tLINE += vbCrLf
                            File.AppendAllText(tFileName, PrintLine_tLINE)
                    End Select
                Case "DANTOM", "REVSPRING"
                    PrintLine_tLINE += "MMMM" + vbCrLf
                    File.AppendAllText(tFileName, PrintLine_tLINE)
                Case "DIAMOND"
                    PrintLine_tLINE += "### BEGIN METADATA"
                    PrintLine_tLINE += "requireExpectedCount = True"
                    PrintLine_tLINE += "expectedCount = " + iNumPrinted.ToString + "."
                    PrintLine_tLINE += "### END METADATA"
                    File.AppendAllText(tFileName, PrintLine_tLINE)
                Case Else
                    MsgStatus("PrintTrailer/Invalid Letter Vendor [" + tLetterVendor + "]", True)
            End Select
        Catch ex As Exception
            MsgError("PrintTrailer", ex.ToString)
        End Try
    End Sub

    Private Function LetterMultiHeader(ByVal sLetterType As String) As String
        Try
            PrintLine_tLINE = ""
            Select Case sLetterType
                Case "1"
                    '
                Case "2"
                    PrintLine_tLINE += "Account #".PadRight(iField1Pad)
                    PrintLine_tLINE += "Patient".PadRight(iField1Pad)
                    PrintLine_tLINE += STR_TRIM("Date Service", 12).PadRight(12)
                    PrintLine_tLINE += "Balance".PadLeft(iField3Pad)
                    PrintLine_tLINE += "  [" + "LOCX]".PadRight(10)
                    PrintLine_tLINE += vbCrLf
                Case "10"
                    '**********************************************************
                    '* 2021-11-26 RFK: REG F LETTER 1
                    PrintLine_tLINE += STR_TRIM("Account #", iField1Pad).PadRight(iField1Pad)
                    PrintLine_tLINE += STR_TRIM("Date Of Serv", iField2Pad).PadRight(iField2Pad)
                    PrintLine_tLINE += STR_TRIM("Balance", iField3Pad).PadLeft(iField3Pad)
                    PrintLine_tLINE += STR_TRIM("Interest", iField3Pad).PadLeft(iField3Pad)
                    PrintLine_tLINE += STR_TRIM("Fees", iField3Pad).PadLeft(iField3Pad)
                    PrintLine_tLINE += STR_TRIM("Payments/Credits", iField3Pad).PadLeft(iField3Pad)
                    PrintLine_tLINE += STR_TRIM("Amount Due", iField3Pad).PadLeft(iField3Pad)
                    PrintLine_tLINE += vbCrLf
                    PrintLine_tLINE += "".PadRight(iField1Pad - 1, "-") + " "
                    PrintLine_tLINE += "".PadRight(iField2Pad - 1, "-") + " "
                    PrintLine_tLINE += " " + "".PadRight(iField3Pad - 1, "-")
                    PrintLine_tLINE += " " + "".PadRight(iField3Pad - 1, "-")
                    PrintLine_tLINE += " " + "".PadRight(iField3Pad - 1, "-")
                    PrintLine_tLINE += " " + "".PadRight(iField3Pad - 1, "-")
                    PrintLine_tLINE += " " + "".PadRight(iField3Pad - 1, "-")
                    PrintLine_tLINE += vbCrLf
                Case "11", "27", "28"
                    PrintLine_tLINE += "Account #".PadRight(iField1Pad)
                    PrintLine_tLINE += "Patient".PadRight(iField1Pad)
                    PrintLine_tLINE += STR_TRIM("Date Service", 12).PadRight(12)
                    PrintLine_tLINE += "Balance".PadLeft(iField3Pad)
                    PrintLine_tLINE += vbCrLf
                    PrintLine_tLINE += "".PadRight(iField1Pad + iField2Pad + 12 + 20, "-") + vbCrLf
                    PrintLine_tLINE += vbCrLf
                Case "12"
                    PrintLine_tLINE += "Facility/Doctor".PadRight(iField1Pad + iField1Pad)
                    'PrintLine_tLINE  += "Doctor".PadRight(iField1Pad)
                    PrintLine_tLINE += "Account #".PadRight(iField1Pad)
                    PrintLine_tLINE += "Patient".PadRight(iField1Pad)
                    PrintLine_tLINE += STR_TRIM("Date Service", 12).PadRight(12)
                    PrintLine_tLINE += "Balance".PadLeft(iField3Pad)
                    PrintLine_tLINE += vbCrLf
                    PrintLine_tLINE += "".PadRight(iField1Pad + iField1Pad + iField1Pad + iField2Pad + 12 + 20, "-") + vbCrLf
                    PrintLine_tLINE += vbCrLf
                Case "14"
                    '
                Case "17", "18"
                    PrintLine_tLINE += "Date     Provider         Patient       Account         Charges   Adjust    Payments  Balance"
                    PrintLine_tLINE += vbCrLf
                    PrintLine_tLINE += "-------- ---------------- ------------- --------------- --------- --------- --------- ---------"
                    PrintLine_tLINE += vbCrLf
                Case "19", "20", "21", "22", "26", "27", "28"
                    PrintLine_tLINE += "Date     Service          Patient       Account         Charges   Adjust    Payments  Balance"
                    PrintLine_tLINE += vbCrLf
                    PrintLine_tLINE += "-------- ---------------- ------------- --------------- --------- --------- --------- ---------"
                    PrintLine_tLINE += vbCrLf
                Case "23", "24", "25", "29"
                    PrintLine_tLINE += "Date".PadRight(11, " ")
                    PrintLine_tLINE += "Service".PadRight(31, " ")
                    PrintLine_tLINE += "Charges".PadRight(20, " ")
                    PrintLine_tLINE += vbCrLf
                    PrintLine_tLINE += "".PadRight(10, "-")
                    PrintLine_tLINE += " "
                    PrintLine_tLINE += "".PadRight(30, "-")
                    PrintLine_tLINE += " "
                    PrintLine_tLINE += "".PadRight(20, "-")
                    PrintLine_tLINE += vbCrLf
            End Select
            Return PrintLine_tLINE
        Catch ex As Exception
            MsgError("LetterHeader", ex.ToString)
        End Try
        Return ""
    End Function

    Private Function LetterMultiTrailer(ByVal sLetterType As String) As String
        Try
            PrintLine_tLINE = ""
            Select Case sLetterType
                Case "1"
                    '
                Case "2"
                    PrintLine_tLINE += "+".PadRight(iField1Pad + iField2Pad + 12) + " ===========".PadLeft(iField3Pad) + vbCrLf
                    PrintLine_tLINE += "Total".PadRight(iField1Pad + iField2Pad + 12) + STR_format(Str(dBalance), "$").PadLeft(iField3Pad) + vbCrLf
                Case "10"
                    '**********************************************************
                    '* 2021-11-26 RFK: REG F LETTER 1
                    PrintLine_tLINE += "".PadRight(iField1Pad + iField2Pad + (iField3Pad * 5), "=") + vbCrLf
                    PrintLine_tLINE += "Total".PadRight(iField1Pad + iField2Pad + (iField3Pad * 4)) + STR_format(Str(dBalance), "$").PadLeft(iField3Pad) + vbCrLf
                Case "11", "27", "28"
                    PrintLine_tLINE += "".PadRight(iField1Pad + iField2Pad + 12 + 20, "=") + vbCrLf
                    PrintLine_tLINE += "Total".PadRight(iField1Pad + iField2Pad + 12) + STR_format(Str(dBalance), "$").PadLeft(iField3Pad) + vbCrLf
                Case "12"
                    PrintLine_tLINE += "".PadRight(iField1Pad + iField1Pad + iField1Pad + iField2Pad + 12 + 20, "=") + vbCrLf
                    PrintLine_tLINE += "Total".PadRight(iField1Pad + iField1Pad + iField1Pad + iField2Pad + 12) + STR_format(Str(dBalance), "$").PadLeft(iField3Pad) + vbCrLf
                Case "14"
                    PrintLine_tLINE += "-----".PadRight(83) + "".PadRight(10, "=") + vbCrLf
                    PrintLine_tLINE += "Total".PadRight(83) + STR_format(Str(dBalance), "$").PadLeft(10) + vbCrLf
                Case "25"
                    PrintLine_tLINE += "".PadRight(iField1Pad + iField2Pad + 12 + 20, "=") + vbCrLf
                    PrintLine_tLINE += "Total".PadRight(iField1Pad + iField2Pad + 12) + STR_format(Str(dBalance), "$").PadLeft(iField3Pad) + vbCrLf
            End Select
            Return PrintLine_tLINE
        Catch ex As Exception
            MsgError("LetterMultiTrailer", ex.ToString)
        End Try
        Return ""
    End Function

    Private Sub TextBox_StatusCodeAfterSent_TextChanged(sender As Object, e As EventArgs) Handles TextBox_StatusCodeAfterSent.TextChanged
        Try
            Label_StatusCodeSentDescription.Text = ""
            If TextBox_StatusCodeAfterSent.Text.Trim.Length > 0 Then
                Select Case rkutils.ReadField(DataGridView_Clients, "ClientType", Val(Label_ClientRow.Text))
                    Case "C"
                        Label_StatusCodeSentDescription.Text = rkutils.SQL_READ_FIELD(DataGridView3, "DB2", "STDESC", DB2SQLConnectionString, DB2SQLuser, "SELECT STDESC FROM ROIDATA.STATP WHERE STMTTP='C' AND STSTAT='" + TextBox_StatusCodeAfterSent.Text.Trim + "'")
                    Case "S"
                        Label_StatusCodeSentDescription.Text = rkutils.SQL_READ_FIELD(DataGridView3, "DB2", "STDESC", DB2SQLConnectionString, DB2SQLuser, "SELECT STDESC FROM ROIDATA.STATP WHERE STMTTP='A' AND STSTAT='" + TextBox_StatusCodeAfterSent.Text.Trim + "'")
                End Select
            End If
        Catch ex As Exception
            MsgError("TextBox_StatusCodeAfterSent_TextChanged", ex.ToString)
        End Try
    End Sub

    Private Sub TextBox_StatusCodeNotSent_TextChanged(sender As Object, e As EventArgs) Handles TextBox_StatusCodeNotSent.TextChanged
        Try
            Label_StatusCodeNotSentDescription.Text = ""
            If TextBox_StatusCodeNotSent.Text.Trim.Length > 0 Then
                Select Case rkutils.ReadField(DataGridView_Clients, "CLIENTTYPE", Val(Label_ClientRow.Text))
                    Case "C"
                        Label_StatusCodeNotSentDescription.Text = rkutils.SQL_READ_FIELD(DataGridView3, "DB2", "STDESC", DB2SQLConnectionString, DB2SQLuser, "SELECT STDESC FROM ROIDATA.STATP WHERE STMTTP='C' AND STSTAT='" + TextBox_StatusCodeNotSent.Text.Trim + "'")
                    Case "S"
                        Label_StatusCodeNotSentDescription.Text = rkutils.SQL_READ_FIELD(DataGridView3, "DB2", "STDESC", DB2SQLConnectionString, DB2SQLuser, "SELECT STDESC FROM ROIDATA.STATP WHERE STMTTP='A' AND STSTAT='" + TextBox_StatusCodeNotSent.Text.Trim + "'")
                End Select
            End If
        Catch ex As Exception
            MsgError("TextBox1_TextChanged", ex.ToString)
        End Try
    End Sub

    Private Sub Button_Excel_Click(sender As Object, e As EventArgs) Handles Button_Excel.Click
        Try
            Dim sTempFile As String = dir_REPORTING + "Working_Reports\eLetters_" + rkutils.DateToday(16) + ".CSV"
            rkutils.DataGridView_to_CSV(DataGridView_Select, sTempFile, True, Label_AccountsRemaining)
            Shell("NOTEPAD.EXE " + sTempFile, AppWinStyle.NormalFocus)
        Catch ex As Exception
            MsgError("Button_Excel_Click", ex.ToString)
        End Try
    End Sub

    Private Function MultiAccountsLIST(ByVal iRow As Integer) As String
        Try
            '*************************************************************
            '* 2014-8-01 RFK: Give a list of all matched LOCX's
            '******************************************************************
            '* 2019-03-14 RFK:
            If swDTable Then
                MultiAccounts_sSysAccountMatched = rkutils.DataTable_ValueByColumnName(dTable_Select, "RAMLOCX", iRow).Trim
            Else
                MultiAccounts_sSysAccountMatched = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAMLOCX", iRow).Trim
            End If
            MultiAccounts_imLocxRow = rkutils.DataGridView_Contains(DataGridView_Multi, "RAMLOCX", sSysAccountMatched)
            MultiAccounts_iMatchedRow = MultiAccounts_imLocxRow
            tMultiMessage = ""
            Do While MultiAccounts_imLocxRow <= DataGridView_Multi.RowCount - 1 And sSysAccountMatched = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAMLOCX", MultiAccounts_imLocxRow).Trim
                If tMultiMessage.Length > 0 Then tMultiMessage += " "
                tMultiMessage += rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RALOCX", MultiAccounts_imLocxRow).Trim
                MultiAccounts_imLocxRow += 1
            Loop
            Return tMultiMessage
        Catch ex As Exception
            MsgError("MultiAccountsLIST", ex.ToString)
            Return ""
        End Try
    End Function

    Private Function MultiAccountsBalance(ByVal iRow As Integer) As String
        Try
            '******************************************************************
            '* 2013-02-04 RFK:
            '******************************************************************
            '* 2019-03-14 RFK:
            If swDTable Then
                MultiAccounts_sSysAccountMatched = rkutils.DataTable_ValueByColumnName(dTable_Select, "RAMLOCX", iRow).Trim
            Else
                MultiAccounts_sSysAccountMatched = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAMLOCX", iRow).Trim
            End If
            MultiAccounts_imLocxRow = rkutils.DataGridView_Contains(DataGridView_Multi, "RAMLOCX", sSysAccountMatched)
            MultiAccounts_iMatchedRow = MultiAccounts_imLocxRow
            Do While MultiAccounts_iMatchedRow <= DataGridView_Multi.RowCount - 1 And sSysAccountMatched = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAMLOCX", MultiAccounts_iMatchedRow).Trim
                tMultiMessage = Val(tMultiMessage) + Val(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RABALD", MultiAccounts_iMatchedRow).Trim)
                MultiAccounts_iMatchedRow += 1
            Loop
            Return tMultiMessage
        Catch ex As Exception
            MsgError("MultiAccountsBalance", ex.ToString)
            Return ""
        End Try
    End Function

    Private Function ListAllLocxByMatchingCriteria(ByVal iRow As Integer) As String
        Try
            '******************************************************************
            '* 2014-8-01 RFK: Give a list of all LOCX's for SSN (closed/not matched)
            tMultiMessage = ""
            '******************************************************************
            '* 2019-03-14 RFK:
            If swDTable Then
                tGSSN = rkutils.DataTable_ValueByColumnName(dTable_Select, "RAGSS#", iRow).Trim
                tMedRec = rkutils.DataTable_ValueByColumnName(dTable_Select, "RAMR#", iRow).Trim
            Else
                tGSSN = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGSS#", iRow).Trim
                tMedRec = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAMR#", iRow).Trim
            End If
            '*************************************************************
            Dim AllLocxBySSN_SQL = "SELECT RALOCX FROM ROIDATA.RACCTP"
            Select Case rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "MATCHBY", Val(Label_ClientRow.Text))
                Case "GSSN"
                    AllLocxBySSN_SQL += " WHERE RAGSS#='" + tGSSN + "'"
                    AllLocxBySSN_SQL += " AND RAGSS# > 0"
                Case "MEDREC"
                    AllLocxBySSN_SQL += " WHERE RAMR#='" + tMedRec + "'"
                    AllLocxBySSN_SQL += " AND RAMR# > 0"
            End Select
            '*************************************************************
            If rkutils.SQL_READ_DATAGRID(DataGridView3, "DB2", "*", DB2SQLConnectionString, DB2SQLuser, AllLocxBySSN_SQL) Then
                MultiAccounts_imLocxRow = 0
                Do While MultiAccounts_imLocxRow <= Me.DataGridView3.RowCount - 1
                    If rkutils.DataGridView_ValueByColumnName(DataGridView3, "RALOCX", MultiAccounts_imLocxRow).Trim.Length > 0 Then
                        If tMultiMessage.Length > 0 Then tMultiMessage += " "
                        tMultiMessage += rkutils.DataGridView_ValueByColumnName(DataGridView3, "RALOCX", MultiAccounts_imLocxRow).Trim
                    End If
                    MultiAccounts_imLocxRow += 1
                Loop
            End If
            '*************************************************************
            Return tMultiMessage
        Catch ex As Exception
            MsgError("AllLocxBySSN", ex.ToString)
            Return ""
        End Try
    End Function

    Protected Sub LettersSentLastMonths(ByVal iMonths As Integer)
        Try
            '******************************************************************
            '* 2016-08-31 RFK: 
            sSQL = "SELECT DISTINCT C.CLIENT"
            '******************************************************************
            For i1 = iMonths - 1 To 0 Step -1
                sTmpSTR = rkutils.STR_format(rkutils.STR_DATE_PLUS("TODAY", "-m", i1.ToString.Trim), "ccyy-mm")
                MsgStatus(sTmpSTR, True)
                sSQL += ", M" + i1.ToString.Trim + ".tCOUNT AS '" + sTmpSTR + "'"
            Next
            sSQL += ",'' AS Total"
            '******************************************************************
            sSQL += " FROM " + sDBO + ".dbo.letter_sent Z" + vbCrLf
            sSQL += " INNER JOIN" + vbCrLf
            sSQL += " (" + vbCrLf
            sSQL += " SELECT DISTINCT CLIENT"
            sSQL += " FROM " + sDBO + ".dbo.letter_sent"
            sSQL += " WHERE letter_matched='-'"
            sSQL += " AND CONVERT(DATE, LETTER_DATE) >= '" + rkutils.STR_format(rkutils.STR_DATE_PLUS("TODAY", "-m", iMonths.ToString.Trim), "ccyy-mm-dd") + "'"
            sSQL += " ) AS C ON C.CLIENT=Z.CLIENT" + vbCrLf
            '******************************************************************
            For i1 = 0 To iMonths - 1
                sSQL += " LEFT JOIN" + vbCrLf
                sSQL += " (" + vbCrLf
                sSQL += " SELECT CLIENT, COUNT(*) AS tCOUNT" + vbCrLf
                sSQL += " FROM " + sDBO + ".dbo.letter_sent" + vbCrLf
                sSQL += " WHERE letter_matched='-'" + vbCrLf
                sSQL += " AND YEAR(LETTER_DATE)='" + rkutils.STR_format(rkutils.STR_DATE_PLUS("TODAY", "-m", i1.ToString.Trim), "ccyy") + "'" + vbCrLf
                sSQL += " AND MONTH(LETTER_DATE)='" + rkutils.STR_format(rkutils.STR_DATE_PLUS("TODAY", "-m", i1.ToString.Trim), "mm") + "'" + vbCrLf
                sSQL += " GROUP BY client" + vbCrLf
                sSQL += " ) AS M" + i1.ToString.Trim + " ON M" + i1.ToString.Trim + ".CLIENT=Z.CLIENT" + vbCrLf
            Next
            sSQL += " ORDER BY client"
            '******************************************************************
            MsgStatus(sSQL, False)
            '******************************************************************
            '* 2016-08-31 RFK: too long of a query for DataGridView, use DTable
            '* 2019-03-14 RFK: LOOK INTO THIS AGAIN 
            Dim DT2 As New DataTable
            rkutils.SQL_READ_DATATABLE(DT2, "MSSQL", "*", msSQLConnectionString, msSQLuser, sSQL)
            DataGridView_Select.DataSource = DT2
            DataGridView_Select.Visible = True
        Catch ex As Exception
            MsgError("LettersSentLastMonths", ex.ToString)
        End Try
    End Sub

    Private Sub Printed(ByVal iLocxRow As Integer, ByVal tLetter As String, ByVal tCODE As String)
        Try
            '******************************************************************
            '* 2019-03-14 RFK:
            If swDTable Then
                If ReadFieldDataTable(dTable_Select, "ERRORCODE", iLocxRow) <> "READY" Then
                    If CheckBox_DEBUG.Checked Then MsgStatus("Printed NOT READY:" + iLocxRow.ToString + " " + ReadFieldDataTable(dTable_Select, "RALOCX", iLocxRow) + " " + tLetter + " " + tCODE + " [" + ReadField(DataGridView_Select, "ERRORCODE", iLocxRow) + "]", False)
                    Exit Sub
                End If
                '******************************************************************
                rkutils.DataTable_SetValueByColumnName(dTable_Select, "ERRORCODE", iLocxRow, tCODE)
                '******************************************************************
                rkutils.DataTable_SetValueByColumnName(dTable_Select, "LETTERPRINTED", iLocxRow, tLetter)
                rkutils.DataTable_SetValueByColumnName(dTable_Select, "LETTERPRINTEDDATE", iLocxRow, rkutils.STR_format("TODAY", "mm/dd/ccyy"))
                '******************************************************************
                '* 2016-11-09 RFK: corrected for Facility
                iRulesRow = DataGridView_Contains2Cols(DataGridView_RULES, "RRFACL", ReadFieldDataTable(dTable_Select, "RAFACL", iLocxRow), "RRACTV", tLetter)
            Else
                If ReadField(DataGridView_Select, "ERRORCODE", iLocxRow) <> "READY" Then
                    If CheckBox_DEBUG.Checked Then MsgStatus("Printed NOT READY:" + iLocxRow.ToString + " " + ReadField(DataGridView_Select, "RALOCX", iLocxRow) + " " + tLetter + " " + tCODE + " [" + ReadField(DataGridView_Select, "ERRORCODE", iLocxRow) + "]", False)
                    Exit Sub
                End If
                '******************************************************************
                rkutils.DataGridView_SetValueByColumnName(DataGridView_Select, "ERRORCODE", iLocxRow, tCODE)
                '******************************************************************
                rkutils.DataGridView_SetValueByColumnName(DataGridView_Select, "LETTERPRINTED", iLocxRow, tLetter)
                rkutils.DataGridView_SetValueByColumnName(DataGridView_Select, "LETTERPRINTEDDATE", iLocxRow, rkutils.STR_format("TODAY", "mm/dd/ccyy"))
                '******************************************************************
                '* 2016-11-09 RFK: corrected for Facility
                iRulesRow = DataGridView_Contains2Cols(DataGridView_RULES, "RRFACL", ReadField(DataGridView_Select, "RAFACL", iLocxRow), "RRACTV", tLetter)
            End If
            '******************************************************************
            If iRulesRow >= 0 Then
                '**************************************************************
                '* 2019-03-14 RFK:
                If swDTable Then
                    If CheckBox_DEBUG.Checked Then MsgStatus("Printed LOCXrow=" + iLocxRow.ToString + " Letter=" + tLetter + " Facility=" + ReadFieldDataTable(dTable_Select, "RAFACL", iLocxRow) + " RulesRow=" + iRulesRow.ToString + " tErrorCode=" + tCODE, False)
                Else
                    If CheckBox_DEBUG.Checked Then MsgStatus("Printed LOCXrow=" + iLocxRow.ToString + " Letter=" + tLetter + " Facility=" + ReadField(DataGridView_Select, "RAFACL", iLocxRow) + " RulesRow=" + iRulesRow.ToString + " tErrorCode=" + tCODE, False)
                End If
                '**************************************************************
                tLetterNext = rkutils.DataGridView_ValueByColumnName(DataGridView_RULES, "RRNACT", iRulesRow).Trim
                tLetterNextDays = rkutils.DataGridView_ValueByColumnName(DataGridView_RULES, "RRDAYS", iRulesRow).Trim
                tLetterNextDate = rkutils.STR_DATE_PLUS("TODAY", "+", tLetterNextDays)
                '******************************************************************
                '* 2019-03-14 RFK:
                If swDTable Then
                    rkutils.DataTable_SetValueByColumnName(dTable_Select, "LETTERNEXT", iLocxRow, tLetterNext)
                    If IsDate(tLetterNextDate) Then
                        rkutils.DataTable_SetValueByColumnName(dTable_Select, "LETTERNEXTDATE", iLocxRow, tLetterNextDate)
                    End If
                Else
                    rkutils.DataGridView_SetValueByColumnName(DataGridView_Select, "LETTERNEXT", iLocxRow, tLetterNext)
                    If IsDate(tLetterNextDate) Then
                        rkutils.DataGridView_SetValueByColumnName(DataGridView_Select, "LETTERNEXTDATE", iLocxRow, tLetterNextDate)
                    End If
                End If
                If CheckBox_DEBUG.Checked Then MsgStatus("Printed LetterNext=" + tLetterNext + " NextDays=" + tLetterNextDays + " NextDate=" + tLetterNextDate, False)
            Else
                '******************************************************************
                '* 2019-03-14 RFK:
                If swDTable Then
                    MsgStatus("ERROR/Printed LOCXrow=" + iLocxRow.ToString + " Letter=" + tLetter + " Facility=" + ReadFieldDataTable(dTable_Select, "RAFACL", iLocxRow) + " RulesRow=" + iRulesRow.ToString + " tErrorCode=" + tCODE, False)
                Else
                    MsgStatus("ERROR/Printed LOCXrow=" + iLocxRow.ToString + " Letter=" + tLetter + " Facility=" + ReadField(DataGridView_Select, "RAFACL", iLocxRow) + " RulesRow=" + iRulesRow.ToString + " tErrorCode=" + tCODE, False)
                End If
            End If
        Catch ex As Exception
            MsgError("Printed", ex.ToString)
        End Try
    End Sub

    Private Function JoinIRSURP() As String
        Try
            '******************************************************************
            '* 2019-08-20 RFK: CORRECT THIS TO AN OUTER JOIN BEFORE USING AGAIN
            Dim sJOIN As String = ""
            sJOIN += " LEFT JOIN ROIDATA.IRSURP I ON I.IRCL#=A.RACL#"
            sJOIN += " AND I.IRTOB=A.RATOB"
            sJOIN += " AND I.IRMTTP=A.RAMTTP"
            sJOIN += " AND I.IRGSS#=A.RAGSS#"
            sJOIN += " AND I.IRGLN5=LEFT(A.RAGLNM,5)"
            sJOIN += " AND I.IRGFN5=LEFT(A.RAGFNM,5)"
            sJOIN += " AND I.IRGZIP=A.RAGZIP"
            sJOIN += " AND I.IRIN#=1"
            sJOIN += " AND I.IRCIN#=1"
            Return sJOIN
        Catch ex As Exception
            MsgError("JoinIRSURP", ex.ToString)
        End Try
        Return ""
    End Function

    Protected Sub LettersSentLastDays()
        Try
            '******************************************************************
            '* 2016-08-30 RFK: 
            Dim iNumDays As Integer = 0, iSentTotalCol As Integer = 0
            sSQL = "SELECT DISTINCT C.CLIENT"
            '******************************************************************
            Select Case rkutils.STR_TRIM(Now.DayOfWeek.ToString, 3)
                Case "Mon"
                    iNumDays = 7
                Case "Tue"
                    iNumDays = 8
                Case "Wed"
                    iNumDays = 9
                Case "Thu"
                    iNumDays = 10
                Case "Fri"
                    iNumDays = 11
                Case "Sat"
                    iNumDays = 12
                Case Else
                    iNumDays = 13
            End Select
            For i1 = iNumDays To 0 Step -1
                sTmpSTR = rkutils.STR_TRIM(rkutils.STR_format(rkutils.STR_DATE_PLUS("TODAY", "-", i1.ToString.Trim), "DOW"), 3) + "-" + rkutils.STR_format(rkutils.STR_DATE_PLUS("TODAY", "-", i1.ToString.Trim), "mm-dd")
                If rkutils.STR_TRIM(sTmpSTR, 3) <> "Sat" And rkutils.STR_TRIM(sTmpSTR, 3) <> "Sun" Then
                    sSQL += ", D" + i1.ToString.Trim + ".DAY" + i1.ToString.Trim + " AS '" + sTmpSTR + "'"
                Else
                    If rkutils.STR_TRIM(sTmpSTR, 3) = "Sun" Then
                        sSQL += ", '0' as SentTotal"
                    End If
                End If
            Next
            sSQL += ", '0' as RunningTotal"
            '******************************************************************
            sSQL += " FROM " + sDBO + ".dbo.letter_sent Z"
            sSQL += " INNER JOIN"
            sSQL += " ("
            sSQL += " SELECT DISTINCT CLIENT"
            sSQL += " FROM " + sDBO + ".dbo.letter_sent"
            sSQL += " WHERE letter_matched='-'"
            sSQL += " AND CONVERT(DATE, LETTER_DATE) >= '" + rkutils.STR_format(rkutils.STR_DATE_PLUS("TODAY", "-", "7"), "ccyy-mm-dd") + "'"
            sSQL += " ) AS C ON C.CLIENT=Z.CLIENT"
            '******************************************************************
            For i1 = 0 To iNumDays
                sSQL += " LEFT JOIN"
                sSQL += " ("
                sSQL += " SELECT CLIENT, COUNT(*) AS DAY" + i1.ToString.Trim
                sSQL += " FROM " + sDBO + ".dbo.letter_sent"
                sSQL += " WHERE letter_matched='-'"
                sSQL += " AND CONVERT(DATE, LETTER_DATE) = '" + rkutils.STR_format(rkutils.STR_DATE_PLUS("TODAY", "-", i1.ToString.Trim), "ccyy-mm-dd") + "'"
                sSQL += " GROUP BY client"
                sSQL += " ) AS D" + i1.ToString.Trim + " ON D" + i1.ToString.Trim + ".CLIENT= Z.CLIENT"
            Next
            sSQL += " ORDER BY client"
            '******************************************************************
            MsgStatus(sSQL, False)
            '******************************************************************
            '* 2019-03-14 RFK:
            If swDTable Then
                '**************************************************************
                rkutils.SQL_READ_DATATABLE(dTable_Select, "MSSQL", "*", msSQLConnectionString, msSQLuser, sSQL)
                '**************************************************************
                If dTable_Select.Rows.Count >= 0 Then
                    '**********************************************************
                    '* 2016-09-12 RFK: Get the Total Column / Sum the Columns
                    MsgStatus("This needs to be converted", True)
                    'DataGridView_Select.Item(0, DataGridView_Select.RowCount - 1).Value = "Total:"
                    'For i1 = 1 To DataGridView_Select.ColumnCount - 1
                    '    If DataGridView_Select.Columns(i1).Name = "SentTotal" Then
                    '        iSentTotalCol = i1
                    '    Else
                    '        DataGridView_Select.Item(i1, DataGridView_Select.RowCount - 1).Value = DataGridView_SumColumn(DataGridView_Select, i1, False)
                    '    End If
                    'Next
                    ''**********************************************************
                    ''* 2016-09-12 RFK: Sum the ROW
                    'For iR = 0 To DataGridView_Select.RowCount - 1
                    '    DataGridView_Select.Item(iSentTotalCol, iR).Value = rkutils.STR_format(DataGridView_SumRow(DataGridView_Select, iR, 1, iSentTotalCol), "#,###")
                    '    DataGridView_Select.Item(iSentTotalCol, iR).Style.BackColor = Color.Gray
                    '    DataGridView_Select.Item(DataGridView_Select.ColumnCount - 1, iR).Value = rkutils.STR_format(DataGridView_SumRow(DataGridView_Select, iR, iSentTotalCol + 1, DataGridView_Select.ColumnCount - 1), "#,###")
                    '    DataGridView_Select.Item(DataGridView_Select.ColumnCount - 1, iR).Style.BackColor = Color.SteelBlue
                    'Next
                    '**********************************************************
                End If
                '**************************************************************
            Else
                '**************************************************************
                DataGridView_Select.Visible = rkutils.SQL_READ_DATAGRID(DataGridView_Select, "MSSQL", "*", msSQLConnectionString, msSQLuser, sSQL)
                '**************************************************************
                If DataGridView_Select.RowCount > 0 Then
                    '**********************************************************
                    '* 2016-09-12 RFK: Get the Total Column / Sum the Columns
                    DataGridView_Select.Item(0, DataGridView_Select.RowCount - 1).Value = "Total:"
                    For i1 = 1 To DataGridView_Select.ColumnCount - 1
                        If DataGridView_Select.Columns(i1).Name = "SentTotal" Then
                            iSentTotalCol = i1
                        Else
                            DataGridView_Select.Item(i1, DataGridView_Select.RowCount - 1).Value = DataGridView_SumColumn(DataGridView_Select, i1, False)
                        End If
                    Next
                    '**********************************************************
                    '* 2016-09-12 RFK: Sum the ROW
                    For iR = 0 To DataGridView_Select.RowCount - 1
                        DataGridView_Select.Item(iSentTotalCol, iR).Value = rkutils.STR_format(DataGridView_SumRow(DataGridView_Select, iR, 1, iSentTotalCol), "#,###")
                        DataGridView_Select.Item(iSentTotalCol, iR).Style.BackColor = Color.Gray
                        DataGridView_Select.Item(DataGridView_Select.ColumnCount - 1, iR).Value = rkutils.STR_format(DataGridView_SumRow(DataGridView_Select, iR, iSentTotalCol + 1, DataGridView_Select.ColumnCount - 1), "#,###")
                        DataGridView_Select.Item(DataGridView_Select.ColumnCount - 1, iR).Style.BackColor = Color.SteelBlue
                    Next
                    '**********************************************************
                End If
            End If
            '******************************************************************
        Catch ex As Exception
            MsgError("LettersSentLastDays", ex.ToString)
        End Try
    End Sub

    Private Sub Button_PreCalc_Click(sender As Object, e As EventArgs) Handles Button_Precalc.Click
        Try
            '******************************************************************
            '* 2021-04-02 RFK:
            '* 2021-08-20 RFK: HCLNTP
            '* 2021-08-20 RFK: C.HCFR30=P.SGROUP
            sSQL = "SELECT RALNAC, COUNT(*) AS tCount"
            sSQL += " FROM ROIDATA.RACCTP A"
            sSQL += " LEFT JOIN ROIDATA.HCLNTP C ON (A.RACL#=C.HCCL#)"
            sSQL += " LEFT JOIN ROIDATA.STATP S ON (A.RARSTA=S.STSTAT AND A.RAMTTP=S.STMTTP)"
            sSQL += " LEFT JOIN ROIDATA.STATEBLOCKING P ON ((RIGHT(TRIM(A.RAGCST),2)=P.POSTALCODE) AND (P.SGROUP=C.HCFR30))"
            sSQL += " LEFT JOIN ROIDATA.DILIGENCE D ON D.DICL# = A.RACL# AND A.RAMBAL >= D.BAL_GTE AND A.RAMBAL <= D.BAL_LTE"
            sSQL += AccountsWhere(sSQL, "PreCalc", rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)), "", 0, "")
            sSQL += " GROUP BY RALNAC"
            MsgStatus(sSQL, True)
            DataGridView_Select.Visible = rkutils.SQL_READ_DATAGRID(DataGridView_Select, "DB2", "*", DB2SQLConnectionString, DB2SQLuser, sSQL)
        Catch ex As Exception
            MsgError("PreCalc", ex.ToString)
        End Try
    End Sub

    Private Function AnnuityOne_MatchedCheck(ByVal iLine As Integer, ByVal imLine As Integer, ByVal sSysAccountMatched As String)
        Try
            '******************************************
            '* 2012-10-01 RFK: 
            Dim iLetterMonth As Integer = 0, iLetterDay As Integer = 0, iLetterYear As Integer = 0
            tLetterPrinted = ""
            If CheckBox_DEBUG.Checked Then MsgStatus("MatchedCheck iLine:" + iLine.ToString + " iMLine:" + imLine.ToString + " RamLocx:" + sSysAccountMatched + " [" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAMLOCX", imLine).Trim + "]", True)
            '******************************************
            '* 2012-10-01 RFK: Get OLDEST LETTER #
            'Dim tLetterNumber As String = "", tLetterDate As String = ""
            'Dim iMatchedRow As Integer = imLine
            AnnuityOne_MatchedCheck_tLetterNumber = ""
            AnnuityOne_MatchedCheck_tLetterDate = ""
            AnnuityOne_MatchedCheck_iMatchedRow = imLine
            Do While AnnuityOne_MatchedCheck_iMatchedRow <= DataGridView_Multi.RowCount - 1 And sSysAccountMatched = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAMLOCX", AnnuityOne_MatchedCheck_iMatchedRow).Trim
                If CheckBox_DEBUG.Checked Then MsgStatus(AnnuityOne_MatchedCheck_iMatchedRow.ToString + " " + DataGridView_Multi.RowCount.ToString + " " + sSysAccountMatched + " " + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAMLOCX", AnnuityOne_MatchedCheck_iMatchedRow).Trim + AnnuityOne_MatchedCheck_tLetterNumber, True)
                '******************************************
                '* 2012-11-06 RFK: check for MAX Letter Number
                Select Case ComboBox_MatchType.Text
                    Case "Sent Today"
                        AnnuityOne_MatchedCheck_tLetterNumber = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RALCAC", AnnuityOne_MatchedCheck_iMatchedRow).Trim
                    Case Else
                        AnnuityOne_MatchedCheck_tLetterNumber = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RALNAC", AnnuityOne_MatchedCheck_iMatchedRow).Trim
                End Select
                '******************************************************************************************
                '******************************************************************************************
                '* 2013-05-01 RFK:
                Select Case ComboBox_ClientActive.Text
                    Case "Testing"
                        If AnnuityOne_MatchedCheck_tLetterNumber = 0 Then
                            '******************************************************************************************
                            '* 2014-03-14 RFK:
                            '* 2014-12-03 RFK:
                            '* 2017-05-16 RFK: changed to ALL TEST clients
                            Select Case rkutils.ReadField(DataGridView_Multi, "RACL#", AnnuityOne_MatchedCheck_iMatchedRow)
                                'Case "UCA"
                                '    AnnuityOne_MatchedCheck_tLetterNumber = "510"
                                Case Else
                                    AnnuityOne_MatchedCheck_tLetterNumber = "310"
                            End Select
                            If CheckBox_DEBUG.Checked Then MsgStatus("AnnuityOne_MatchedCheck/DEBUG/using " + AnnuityOne_MatchedCheck_tLetterNumber + " instead of 0", True)
                        End If
                End Select
                '******************************************************************************************
                If CheckBox_DEBUG.Checked Then MsgStatus("MatchedCheck:" + iLine.ToString + " RamLocx:" + sSysAccountMatched + " LetterNumber:" + AnnuityOne_MatchedCheck_tLetterNumber, True)
                '******************************************************************************************
                Select Case Val(AnnuityOne_MatchedCheck_tLetterNumber)
                    Case 310, 910
                        '******************************************
                        '* 2012-11-09 RFK: 1st letters
                        If Val(AnnuityOne_MatchedCheck_tLetterNumber) > Val(tLetterPrinted) Then tLetterPrinted = AnnuityOne_MatchedCheck_tLetterNumber
                        '******************************************
                    Case Else
                        '******************************************
                        '* 2012-11-09 RFK:
                        If Val(AnnuityOne_MatchedCheck_tLetterNumber) > Val(tLetterPrinted) Then tLetterPrinted = AnnuityOne_MatchedCheck_tLetterNumber
                        '**********************************************************
                        '* 2012-11-09 RFK: Past 1st Letter 310/910 Check For Future
                        Select Case ComboBox_MatchType.Text
                            Case "Sent Today"
                                'Do Not Check for Future [Sent Today]
                            Case Else
                                '****************************************************************************
                                '* 2014-08-05 RFK: eLetters_FutureLetterDates
                                Select Case rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "ELETTERSFUTURELETTERDATES", Val(Label_ClientRow.Text))
                                    Case "N"
                                        'NO do not look at future letter dates, All Selected will get a letter
                                    Case Else
                                        '*******************************************************
                                        '* 2014-08-19 RFK: Only if valid letternumber (and date)
                                        If Val(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RALNAC", AnnuityOne_MatchedCheck_iMatchedRow).Trim) > 0 Then
                                            '****************************************************************
                                            iLetterMonth = Val(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RANLMO", AnnuityOne_MatchedCheck_iMatchedRow))
                                            iLetterDay = Val(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RANLDY", AnnuityOne_MatchedCheck_iMatchedRow))
                                            iLetterYear = Val(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RANLYR", AnnuityOne_MatchedCheck_iMatchedRow))
                                            AnnuityOne_MatchedCheck_tLetterDate = iLetterMonth.ToString.Trim
                                            AnnuityOne_MatchedCheck_tLetterDate += "/" + iLetterDay.ToString.Trim
                                            AnnuityOne_MatchedCheck_tLetterDate += "/" + iLetterYear.ToString.Trim
                                            If rkutils.STR_format(AnnuityOne_MatchedCheck_tLetterDate, "ccyymmdd") > rkutils.STR_format("TODAY", "ccyymmdd") Then
                                                '**************************************************
                                                '* 2012-11-09 RFK: Matched future date
                                                'MatchedFuture(iLine, imLine, sSysAccountMatched, tRALOCX, AnnuityOne_MatchedCheck_tLetterNumber, AnnuityOne_MatchedCheck_tLetterDate)
                                                tRALOCX = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RALOCX", iLine).Trim
                                                '******************************
                                                If CheckBox_DEBUG.Checked Then MsgStatus("AnnuityOne_MatchedCheck:" + tRALOCX + " " + " " + tLetterNext + " " + tLetterNextDate, True)
                                                MatchedFuture_SQLstring = "UPDATE ROIDATA.RACCTP"
                                                MatchedFuture_SQLstring += " SET RANLMO=" + iLetterMonth.ToString
                                                MatchedFuture_SQLstring += ",RANLDY=" + iLetterDay.ToString
                                                MatchedFuture_SQLstring += ",RANLYR=" + iLetterYear.ToString
                                                MatchedFuture_SQLstring += " WHERE RALOCX=" + tRALOCX
                                                '******************************
                                                If CheckBox_DEBUG.Checked Then MsgStatus(MatchedFuture_SQLstring, False)
                                                '******************************
                                                DB_COMMAND("DB2", DB2SQLConnectionString, DB2SQLuser, MatchedFuture_SQLstring)
                                                TRACKS_update("MSSQL", msSQLConnectionString, msSQLuser, "", tRALOCX, "", "L", "LETTER MATCHED DATE CHANGED [" + rkutils.STR_format(AnnuityOne_MatchedCheck_tLetterDate, "mm-dd-ccyy") + "]")
                                                '****************************************************************
                                                LettersPrinted("FUTURE", tLOCX, True)
                                                Return False
                                            End If
                                        End If
                                End Select
                        End Select
                        '******************************************
                End Select
                '******************************************
                AnnuityOne_MatchedCheck_iMatchedRow += 1
            Loop
            If CheckBox_DEBUG.Checked Then MsgStatus("MatchedCheck Exiting / RamLocx:" + sSysAccountMatched + " LetterPrinted:" + tLetterPrinted, True)
            If Val(tLetterPrinted) <= 0 Then Return False
            Return True
        Catch ex As Exception
            MsgError("MatchedCheck", ex.ToString)
        End Try
        Return False
    End Function

    Private Sub AccountsLoadForClient()
        Try
            '******************************************************************
            '* 2014-12-03 RFK:
            tSTR = STR_format(ReadField(DataGridView_Clients, "MatchComplete", Val(Label_ClientRow.Text)), "ccyymmdd")
            If Val(tSTR) < Val(STR_format(STR_DATE_PLUS("TODAY", "-", "3"), "ccyymmdd")) Then
                MsgStatus("Client last matched:" + tSTR, True)
                Select Case sSITE
                    Case "AnnuityOne"
                        rkutils.EMAILIT(msSQLConnectionString, msSQLuser, "DoNotReplyeLetterError@AnnuityHealth.com", "eLetters", "letters@AnnuityHealth.com", "IT", Me.Text, "eLetters " + ReadField(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)) + " NOT MATCHED CAN NOT CONTINUE", "Last Matched on:" + ReadField(DataGridView_Clients, "MatchComplete", Val(Label_ClientRow.Text)), "", "")
                        Return
                    Case Else
                        rkutils.EMAILIT(msSQLConnectionString, msSQLuser, "DoNotReplyeLetterError@fostertech.net", "eLetters", "letters@FosterTech.net", "IT", Me.Text, "eLetters " + ReadField(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)) + " NOT MATCHED CAN NOT CONTINUE", "Last Matched on:" + ReadField(DataGridView_Clients, "MatchComplete", Val(Label_ClientRow.Text)), "", "")
                End Select
                'Return
            End If
            '******************************************************************
            If Panel_RegF.Visible Then
                '**************************************************************
                MsgStatus("RegF LOAD NA (use RegF load buttons)", True)
                Exit Sub
            End If
            '******************************************************************
            Button_Run.Enabled = False
            Button_Run.Text = "Please wait"
            '******************************************************************
            Select Case sSITE
                Case "AnnuityOne"
                    AnnuityOne_AccountsLoad(rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "TOB", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "FACILITY", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTMATCHTYPE", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "MATCHBY", Val(Label_ClientRow.Text)), gtLetterType, "*", 0)
                Case Else
                    'iTCS_AccountsLoad(rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "TOB", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "FACILITY", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTMATCHTYPE", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "MATCHBY", Val(Label_ClientRow.Text)))
            End Select
            If DataGridView_Select.RowCount > 1 Or Me.CheckBox_Update.Checked = False Then
                Select Case sSITE
                    Case "AnnuityOne"
                        '*********************************************************************************************************
                        '* 2014-08-05 RFK: Don't load all accounts; instead of 1 big load for larger clients, single matched loads
                        If swReadAllMatched = True Then
                            AnnuityOne_AccountsLoad(rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "TOB", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "FACILITY", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTMATCHTYPE", Val(Label_ClientRow.Text)), "MATCHED", gtLetterType, "*", 0)
                        End If
                    Case Else
                        'iTCS_AccountsLoad(rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "TOB", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "FACILITY", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTMATCHTYPE", Val(Label_ClientRow.Text)), "MATCHED")
                End Select
            Else
                Label_Total.Text = "0"
            End If

            Button_Run.Text = "Run"
            Button_Run.Enabled = True

            '******************************************************************
            LettersPreCalc()
            '******************************************************************
            'MsgStatus("AccountsLoadForClient", True)
        Catch ex As Exception
            MsgError("AccountsLoadForClient", ex.ToString)
        End Try
    End Sub

    Private Function Letter_Types_Value(ByVal tLNumber As String, ByVal tField As String) As String
        Try
            For i1 = 0 To DataGridView_Letter_Types.RowCount - 1
                If rkutils.DataGridView_ValueByColumnName(DataGridView_Letter_Types, "LNUMBER", i1).Trim = tLNumber Then
                    Return rkutils.DataGridView_ValueByColumnName(DataGridView_Letter_Types, tField, i1).Trim
                End If
            Next
            Return ""
        Catch ex As Exception
            MsgError("Letter_Types_Value", ex.ToString)
        End Try
        Return ""
    End Function

    Private Sub Button_LoadRegF_Click(sender As Object, e As EventArgs) Handles Button_LoadRegF.Click
        Try
            Button_Run.Enabled = False
            Button_Run.Text = "Please wait"
            '******************************************************************
            Select Case sSITE
                Case "AnnuityOne"
                    AnnuityOne_AccountsLoad(rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "TOB", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "FACILITY", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTMATCHTYPE", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "MATCHBY", Val(Label_ClientRow.Text)), gtLetterType, "*", 1)
                Case Else
                    '**********************************************************
            End Select
            Button_Run.Text = "Run"
            Button_Run.Enabled = True
            '******************************************************************
            LettersPreCalc()
            '******************************************************************
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button_LoadRegFnot_Click(sender As Object, e As EventArgs) Handles Button_LoadRegFnot.Click
        Try
            If Panel_RegF.Visible Then
                Button_Run.Enabled = False
                Button_Run.Text = "Please wait"
                '******************************************************************
                Select Case sSITE
                    Case "AnnuityOne"
                        AnnuityOne_AccountsLoad(rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "TOB", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "FACILITY", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTMATCHTYPE", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "MATCHBY", Val(Label_ClientRow.Text)), gtLetterType, "*", 2)
                    Case Else
                        '**********************************************************
                        MsgStatus(sSITE, True)
                End Select
                Button_Run.Text = "Run"
                Button_Run.Enabled = True
                '******************************************************************
                LettersPreCalc()
                '******************************************************************
            End If
        Catch ex As Exception
            MsgError("Button_LoadRegFnot_Click", ex.ToString)
        End Try
    End Sub

    Private Sub Button_Run_Click(sender As System.Object, e As System.EventArgs) Handles Button_Run.Click
        If Me.Button_Run.Text = "STOP" Then
            Me.Button_Run.Text = "Run"
            Label_ClientRunning.Text = "Stopping"
            Label_RUNNING.Text = "Stopping"
            MsgStatus("Stopped", True)
            Exit Sub
        End If
        Run()
        Me.Button_Run.Text = "Run"
    End Sub

    Private Function LetterMultiLine(ByVal sLetterType As String, ByVal LetterMultiLine_iLine As Integer, ByVal LetterMultiLine_LOCX As String, bGhost As Boolean) As String
        Try
            '******************************************************************
            PrintLine_tLINE = ""
            MsgStatus("LetterMultiLine:" + MatchedPrint_iPrintedLines.ToString, CheckBox_DEBUG.Checked)
            '*************************************************************
            Select Case sLetterType
                Case "1"
                    '
                Case "2"
                    '*********************************************************
                    '* Account Number
                    PrintLine_tLINE += STR_RIGHT(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAACCT", LetterMultiLine_iLine).Trim, iField1Pad).PadRight(iField1Pad)
                    '*********************************************************
                    '* Patient
                    PrintLine_tLINE += STR_RIGHT(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAPNAM", LetterMultiLine_iLine).Trim, iField2Pad).PadRight(iField2Pad)
                    '*********************************************************
                    '* Admit Date
                    'PrintLine_tLINE  += rkutils.DataGridView_ValueByColumnName(DataGridView1, "RAADMM", LetterMultiLine_iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView1, "RAADMD", LetterMultiLine_iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView1, "RAADMY", LetterMultiLine_iLine).Trim
                    '*********************************************************
                    '* DOS
                    PrintLine_tLINE += STR_TRIM(STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISM", LetterMultiLine_iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISD", LetterMultiLine_iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISY", LetterMultiLine_iLine).Trim, "mm/dd/ccyy"), 12).PadRight(12)
                    '*********************************************************
                    '* Balance Due
                    sCurrentBalance = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RABALD", LetterMultiLine_iLine).Trim
                    dBalance += Val(sCurrentBalance)
                    PrintLine_tLINE += STR_format(sCurrentBalance, "$").PadLeft(iField3Pad)
                    '*********************************************************
                    '* LOCX
                    PrintLine_tLINE += "  [" + STR_RIGHT(tRALOCX.Trim + "]", 10).PadRight(10)
                    '**********************************************************
                    PrintLine_tLINE += vbCrLf
                Case "27"
                    '**************************************************
                    '* 2021-10-19 RFK: ROLL UP BY DOS
                    sDOS = STR_TRIM(STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISM", LetterMultiLine_iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISD", LetterMultiLine_iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISY", LetterMultiLine_iLine).Trim, "mm/dd/ccyy"), 12).PadRight(12)
                    dBalanceRU += Val(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RABALD", LetterMultiLine_iLine).Trim)
                    dBalance += Val(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RABALD", LetterMultiLine_iLine).Trim)
                    '**************************************************
                    bRollUp = False
                    If LetterMultiLine_iLine < DataGridView_Multi.Rows.Count - 1 Then
                        sDOSsave = STR_TRIM(STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISM", LetterMultiLine_iLine + 1).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISD", LetterMultiLine_iLine + 1).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISY", LetterMultiLine_iLine + 1).Trim, "mm/dd/ccyy"), 12).PadRight(12)
                        If sDOS = sDOSsave And sDOSsave <> "" Then
                            '**************************************************
                            bRollUp = True
                        End If
                    End If
                    '**********************************************************
                    If bRollUp Then
                        'PrintLine_tLINE += "RU:" + STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RALOCX", LetterMultiLine_iLine).Trim, iField1Pad).PadRight(iField1Pad)
                        ''******************************************************
                        ''* Balance Due
                        'PrintLine_tLINE += STR_format(dBalanceRU.ToString.Trim, "$").PadLeft(iField3Pad)
                        'sCurrentBalance = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RABALD", LetterMultiLine_iLine).Trim
                        'PrintLine_tLINE += STR_format(sCurrentBalance, "$").PadLeft(iField3Pad)
                        'PrintLine_tLINE += vbCrLf
                    Else
                        PrintLine_tLINE += STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAMR#", LetterMultiLine_iLine).Trim, iField1Pad).PadRight(iField1Pad)
                        PrintLine_tLINE += STR_RIGHT(STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAPNAM", LetterMultiLine_iLine).Trim, iField2Pad - 1), iField2Pad).PadRight(iField2Pad)
                        PrintLine_tLINE += STR_TRIM(STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISM", LetterMultiLine_iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISD", LetterMultiLine_iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISY", LetterMultiLine_iLine).Trim, "mm/dd/ccyy"), 12).PadRight(12)
                        PrintLine_tLINE += STR_format(dBalanceRU.ToString.Trim, "$").PadLeft(iField3Pad)
                        PrintLine_tLINE += vbCrLf
                        '******************************************************
                        dBalanceRU = 0
                        '******************************************************
                    End If
                    '**********************************************************
                Case "10"
                    MsgStatus("WARNING SHOULD NOT MAKE IT HERE (MULTI)", True)
                Case "11", "28"
                    '**********************************************************
                    '* 2021-10-14 RFK: 
                    Select Case sLetterType
                        Case "11"
                            '**************************************************
                            '* 2021-10-14 RFK: Account Number
                            PrintLine_tLINE += STR_RIGHT(STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAACCT", LetterMultiLine_iLine).Trim + "-" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RASUFX", LetterMultiLine_iLine).Trim, iField1Pad - 1), iField1Pad).PadRight(iField1Pad)
                        Case "28"
                            '**************************************************
                            '* 2021-10-14 RFK: ACCOUNT NUMBER NO SUFX
                            PrintLine_tLINE += STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAACCT", LetterMultiLine_iLine), iField1Pad).PadRight(iField1Pad)
                    End Select
                    '*********************************************************
                    '* Patient
                    PrintLine_tLINE += STR_RIGHT(STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAPNAM", LetterMultiLine_iLine).Trim, iField2Pad - 1), iField2Pad).PadRight(iField2Pad)
                    '*********************************************************
                    '* DOS
                    PrintLine_tLINE += STR_TRIM(STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISM", LetterMultiLine_iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISD", LetterMultiLine_iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISY", LetterMultiLine_iLine).Trim, "mm/dd/ccyy"), 12).PadRight(12)
                    '*********************************************************
                    '* Balance Due
                    sCurrentBalance = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RABALD", LetterMultiLine_iLine).Trim
                    dBalance += Val(sCurrentBalance)
                    PrintLine_tLINE += STR_format(sCurrentBalance, "$").PadLeft(iField3Pad)
                    '*********************************************************
                    '* TESTING DATA BELOW
                    'PrintLine_tLINE  += "  [" + STR_RIGHT(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RALOCX", LetterMultiLine_iLine).Trim + "]", 10).PadRight(10)
                    'PrintLine_tLINE  += "  [" + STR_RIGHT(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RALNAC", LetterMultiLine_iLine).Trim + "]", 10).PadRight(10)
                    '*********************************************************
                    PrintLine_tLINE += vbCrLf
                Case "12"
                    '*********************************************************
                    '* Facility
                    tFacility = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAFACL", LetterMultiLine_iLine).Trim
                    iFacilityRow = rkutils.DataGridView_Contains(DataGridView_Facilities, "FARFID", tFacility)
                    PrintLine_tLINE += STR_RIGHT(STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_Facilities, "HFNAME", iFacilityRow).Trim, iField1Pad * 2 - 1), iField1Pad).PadRight(iField1Pad * 2)
                    '*********************************************************
                    '* Account Number
                    PrintLine_tLINE += STR_RIGHT(STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAACCT", LetterMultiLine_iLine).Trim + "-" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RASUFX", LetterMultiLine_iLine).Trim, iField1Pad - 1), iField1Pad).PadRight(iField1Pad)
                    '*********************************************************
                    '* Patient
                    PrintLine_tLINE += STR_RIGHT(STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAPNAM", LetterMultiLine_iLine).Trim, iField2Pad - 1), iField2Pad).PadRight(iField2Pad)
                    '*********************************************************
                    '* Admit Date
                    'PrintLine_tLINE  += rkutils.DataGridView_ValueByColumnName(DataGridView1, "RAADMM", LetterMultiLine_iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView1, "RAADMD", LetterMultiLine_iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView1, "RAADMY", LetterMultiLine_iLine).Trim
                    '*********************************************************
                    '* DOS
                    PrintLine_tLINE += STR_TRIM(STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISM", LetterMultiLine_iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISD", LetterMultiLine_iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISY", LetterMultiLine_iLine).Trim, "mm/dd/ccyy"), 12).PadRight(12)
                    '*********************************************************
                    '* Balance Due
                    sCurrentBalance = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RABALD", LetterMultiLine_iLine).Trim
                    dBalance += Val(sCurrentBalance)
                    PrintLine_tLINE += STR_format(sCurrentBalance, "$").PadLeft(iField3Pad)
                    '*********************************************************
                    PrintLine_tLINE += vbCrLf
                    '*********************************************************
                    '* Doctor
                    PrintLine_tLINE += STR_RIGHT(STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RARDRN", LetterMultiLine_iLine).Trim, iField1Pad * 2 - 1), iField1Pad).PadRight(iField1Pad * 2)
                    PrintLine_tLINE += vbCrLf
                    '*********************************************************
                    '* Blank Line
                    PrintLine_tLINE += vbCrLf
                Case "14"
                    '*********************************************************
                    '* 2013-11-15 RFK: 1st Line DOS (DANTOM STRIPS LEADING SPACES
                    '* 2013-11-15 RFK: Item Descriptions
                    sSQL = "SELECT DISTINCT DIGITS(CHG.ERTRYR)||'/'||DIGITS(CHG.ERTRMO)||'/'||DIGITS(CHG.ERTRDY) AS Date"
                    sSQL += ",ERUNIT"
                    sSQL += ",ERFR30"
                    sSQL += ",ERCHRG"
                    sSQL += " FROM ROIDATA.ERCHGP CHG"
                    sSQL += " WHERE ERLOCX='" + ReadField(DataGridView_Multi, "RALOCX", LetterMultiLine_iLine) + "'"
                    sSQL += " ORDER BY DIGITS(CHG.ERTRYR)||'/'||DIGITS(CHG.ERTRMO)||'/'||DIGITS(CHG.ERTRDY) "
                    SQL_READ_FIELD(DataGridView3, "DB2", "*", DB2SQLConnectionString, DB2SQLuser, sSQL)
                    For i1 = 0 To Me.DataGridView3.RowCount - 1
                        If rkutils.DataGridView_ValueByColumnName(DataGridView3, "ERCHRG", i1).Length > 0 Then
                            PrintLine_tLINE += STR_TRIM(STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISM", LetterMultiLine_iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISD", LetterMultiLine_iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISY", LetterMultiLine_iLine).Trim, "mm/dd/yy"), 9).PadRight(9)
                            PrintLine_tLINE += STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView3, "ERUNIT", i1), 4).Trim.PadLeft(4)                            'QTY
                            PrintLine_tLINE += " "
                            PrintLine_tLINE += STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView3, "ERFR30", i1), 30).Trim.PadRight(30)                         'DESCRIPTION
                            PrintLine_tLINE += " "
                            PrintLine_tLINE += STR_TRIM(STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView3, "ERCHRG", i1), "0.00"), 9).Trim.PadLeft(9)         'AMOUNT
                            PrintLine_tLINE += vbCrLf
                        End If
                    Next
                    '*********************************************************
                    '* 2013-11-15 RFK: 
                    PrintLine_tLINE += "-------- "
                    PrintLine_tLINE += " ".PadRight(5)                'QTY
                    PrintLine_tLINE += " ".PadRight(30)               'DESCRIPTION
                    PrintLine_tLINE += " ---------"
                    PrintLine_tLINE += "  --------"
                    PrintLine_tLINE += " --------"
                    PrintLine_tLINE += " ---------"
                    PrintLine_tLINE += " ---------"
                    PrintLine_tLINE += vbCrLf
                    '*********************************************************
                    '* 2013-11-15 RFK: SubTotal Line (DANTOM STRIPS LEADING SPACES
                    PrintLine_tLINE += "SubTotal".PadRight(9)
                    PrintLine_tLINE += " ".PadRight(5)                'QTY
                    PrintLine_tLINE += STR_TRIM("Order:" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAACCT", LetterMultiLine_iLine).Trim, 30).PadRight(30)
                    '*********************************************************
                    '* Charges
                    sCharges = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAOBAL", LetterMultiLine_iLine).Trim
                    PrintLine_tLINE += STR_TRIM(STR_format(sCharges, "0.00"), 10).PadLeft(10)
                    '*********************************************************
                    '* Ins
                    '* 2013-10-01 RFK: PrintLine_tLINE  += STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RABLANK", LetterMultiLine_iLine).Trim, 10).PadLeft(10)
                    sSQL = "SELECT SUM(A.CXAMT) AS TSUM FROM"
                    sSQL += "("
                    sSQL += "SELECT DISTINCT LEFT(CXCODE,2),CXDATE,CXAMT,CXDESC"
                    sSQL += " FROM ROIDATA.CXTRNP"
                    sSQL += " WHERE CXLOCX='" + ReadField(DataGridView_Multi, "RALOCX", LetterMultiLine_iLine) + "'"
                    sSQL += " AND LEFT(CXCODE,2)='PI'"
                    sSQL += " GROUP BY LEFT(CXCODE,2),CXDATE,CXDESC,CXAMT"
                    sSQL += ") A"
                    sInsurance = SQL_READ_FIELD(DataGridView3, "DB2", "TSUM", DB2SQLConnectionString, DB2SQLuser, sSQL)
                    PrintLine_tLINE += STR_TRIM(STR_format(sInsurance, "0.00"), 10).PadLeft(10)
                    '*********************************************************
                    '* Self
                    '* 2013-10-01 RFK: PrintLine_tLINE  += STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RABLANK", LetterMultiLine_iLine).Trim, 10).PadLeft(10)
                    sSQL = "SELECT SUM(A.CXAMT) AS TSUM FROM"
                    sSQL += "("
                    sSQL += "SELECT DISTINCT LEFT(CXCODE,2),CXDATE,CXAMT,CXDESC"
                    sSQL += " FROM ROIDATA.CXTRNP"
                    sSQL += " WHERE CXLOCX='" + ReadField(DataGridView_Multi, "RALOCX", LetterMultiLine_iLine) + "'"
                    sSQL += " AND LEFT(CXCODE,2)='PP'"
                    sSQL += " GROUP BY LEFT(CXCODE,2),CXDATE,CXDESC,CXAMT"
                    sSQL += ") A"
                    sSelf = SQL_READ_FIELD(DataGridView3, "DB2", "TSUM", DB2SQLConnectionString, DB2SQLuser, sSQL)
                    PrintLine_tLINE += STR_TRIM(STR_format(sSelf, "0.00"), 9).PadLeft(9)

                    '*********************************************************
                    '* Credits
                    '* 2013-10-01 RFK: dTemp = Val(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAOPD", LetterMultiLine_iLine).Trim) + Val(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RATOTP", LetterMultiLine_iLine).Trim) + Val(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAOADJ", LetterMultiLine_iLine).Trim) + Val(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RATOTA", LetterMultiLine_iLine).Trim)
                    '* 2013-10-01 RFK: PrintLine_tLINE  += STR_format(Str(dTemp), "0.00").PadLeft(10)
                    sSQL = "SELECT SUM(A.CXAMT) AS TSUM FROM"
                    sSQL += "("
                    sSQL += "SELECT DISTINCT LEFT(CXCODE,1),CXDATE,CXAMT,CXDESC"
                    sSQL += " FROM ROIDATA.CXTRNP"
                    sSQL += " WHERE CXLOCX='" + ReadField(DataGridView_Multi, "RALOCX", LetterMultiLine_iLine) + "'"
                    sSQL += " AND LEFT(CXCODE,1)='A'"
                    sSQL += " GROUP BY LEFT(CXCODE,1),CXDATE,CXDESC,CXAMT"
                    sSQL += ") A"
                    sCredits = SQL_READ_FIELD(DataGridView3, "DB2", "TSUM", DB2SQLConnectionString, DB2SQLuser, sSQL)
                    '*********************************************************
                    '* 2013-10-10 RFK: Current Balance needed for Calculated
                    sCurrentBalance = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RABALD", LetterMultiLine_iLine).Trim
                    dBalance += Val(sCurrentBalance)
                    '*********************************************************
                    '* 2013-10-10 RFK: Calculated Balance
                    dTemp = Val(sCharges) + Val(sInsurance) + Val(sSelf) + Val(sCredits)
                    If STR_format(dTemp.ToString, "0.00").Trim = STR_format(sCurrentBalance, "0.00").Trim Then
                        PrintLine_tLINE += STR_TRIM(STR_format(sCredits, "0.00"), 10).PadLeft(10)
                    Else
                        dTemp2 = Val(sCredits) + Val(sCurrentBalance) - dTemp
                        PrintLine_tLINE += STR_TRIM(STR_format(dTemp2.ToString, "0.00"), 10).PadLeft(10)
                    End If
                    '*********************************************************
                    '* Balance Due
                    PrintLine_tLINE += STR_format(sCurrentBalance, "0.00").PadLeft(10)
                    PrintLine_tLINE += vbCrLf
                    PrintLine_tLINE += "." + vbCrLf
                Case "15"
                    '*********************************************************
                    '* DOS
                    PrintLine_tLINE += STR_TRIM(STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISM", LetterMultiLine_iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISD", LetterMultiLine_iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISY", LetterMultiLine_iLine).Trim, "mm/dd/yy"), 9).PadRight(9)
                    PrintLine_tLINE += " ".PadRight(5) 'QTY
                    '*********************************************************
                    '* Account Number
                    PrintLine_tLINE += STR_TRIM("Account:" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAACCT", LetterMultiLine_iLine).Trim, 29).PadRight(29)
                    '*********************************************************
                    '* Charges
                    sCharges = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAOBAL", LetterMultiLine_iLine).Trim
                    PrintLine_tLINE += STR_TRIM(STR_format(sCharges, "0.00"), 10).PadLeft(10)
                    '*********************************************************
                    '* Ins
                    '* 2013-10-01 RFK: PrintLine_tLINE  += STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RABLANK", LetterMultiLine_iLine).Trim, 10).PadLeft(10)
                    sSQL = "SELECT SUM(A.CXAMT) AS TSUM FROM"
                    sSQL += "("
                    sSQL += "SELECT DISTINCT LEFT(CXCODE,2),CXDATE,CXAMT,CXDESC"
                    sSQL += " FROM ROIDATA.CXTRNP"
                    sSQL += " WHERE CXLOCX='" + ReadField(DataGridView_Multi, "RALOCX", LetterMultiLine_iLine) + "'"
                    sSQL += " AND LEFT(CXCODE,2)='PI'"
                    sSQL += " GROUP BY LEFT(CXCODE,2),CXDATE,CXDESC,CXAMT"
                    sSQL += ") A"
                    sInsurance = SQL_READ_FIELD(DataGridView3, "DB2", "TSUM", DB2SQLConnectionString, DB2SQLuser, sSQL)
                    PrintLine_tLINE += STR_TRIM(STR_format(sInsurance, "0.00"), 10).PadLeft(10)
                    '*********************************************************
                    '* Self
                    '* 2013-10-01 RFK: PrintLine_tLINE  += STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RABLANK", LetterMultiLine_iLine).Trim, 10).PadLeft(10)
                    sSQL = "SELECT SUM(A.CXAMT) AS TSUM FROM"
                    sSQL += "("
                    sSQL += "SELECT DISTINCT LEFT(CXCODE,2),CXDATE,CXAMT,CXDESC"
                    sSQL += " FROM ROIDATA.CXTRNP"
                    sSQL += " WHERE CXLOCX='" + ReadField(DataGridView_Multi, "RALOCX", LetterMultiLine_iLine) + "'"
                    sSQL += " AND LEFT(CXCODE,2)='PP'"
                    sSQL += " GROUP BY LEFT(CXCODE,2),CXDATE,CXDESC,CXAMT"
                    sSQL += ") A"
                    sSelf = SQL_READ_FIELD(DataGridView3, "DB2", "TSUM", DB2SQLConnectionString, DB2SQLuser, sSQL)
                    PrintLine_tLINE += STR_TRIM(STR_format(sSelf, "0.00"), 10).PadLeft(10)
                    '*********************************************************
                    '* Credits
                    '* 2013-10-01 RFK: dTemp = Val(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAOPD", LetterMultiLine_iLine).Trim) + Val(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RATOTP", LetterMultiLine_iLine).Trim) + Val(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAOADJ", LetterMultiLine_iLine).Trim) + Val(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RATOTA", LetterMultiLine_iLine).Trim)
                    '* 2013-10-01 RFK: PrintLine_tLINE  += STR_format(Str(dTemp), "0.00").PadLeft(10)
                    sSQL = "SELECT SUM(A.CXAMT) AS TSUM FROM"
                    sSQL += "("
                    sSQL += "SELECT DISTINCT LEFT(CXCODE,1),CXDATE,CXAMT,CXDESC"
                    sSQL += " FROM ROIDATA.CXTRNP"
                    sSQL += " WHERE CXLOCX='" + ReadField(DataGridView_Multi, "RALOCX", LetterMultiLine_iLine) + "'"
                    sSQL += " AND LEFT(CXCODE,1)='A'"
                    sSQL += " GROUP BY LEFT(CXCODE,1),CXDATE,CXDESC,CXAMT"
                    sSQL += ") A"
                    sCredits = SQL_READ_FIELD(DataGridView3, "DB2", "TSUM", DB2SQLConnectionString, DB2SQLuser, sSQL)
                    '*********************************************************
                    '* 2013-10-10 RFK: Current Balance needed for Calculated
                    sCurrentBalance = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RABALD", LetterMultiLine_iLine).Trim
                    dBalance += Val(sCurrentBalance)
                    '*********************************************************
                    '* 2013-10-10 RFK: Calculated Balance
                    dTemp = Val(sCharges) + Val(sInsurance) + Val(sSelf) + Val(sCredits)
                    If STR_format(dTemp.ToString, "0.00").Trim = STR_format(sCurrentBalance, "0.00").Trim Then
                        PrintLine_tLINE += STR_TRIM(STR_format(sCredits, "0.00"), 10).PadLeft(10)
                    Else
                        dTemp2 = Val(sCredits) + Val(sCurrentBalance) - dTemp
                        PrintLine_tLINE += STR_TRIM(STR_format(dTemp2.ToString, "0.00"), 10).PadLeft(10)
                    End If
                    '*********************************************************
                    '* Balance Due
                    PrintLine_tLINE += STR_format(sCurrentBalance, "0.00").PadLeft(10)
                    '*********************************************************
                    '* 2013-10-10 RFK: (UnComment for Debugging) 
                    'If STR_format(dTemp.ToString, "0.00").Trim = STR_format(sCurrentBalance, "0.00").Trim Then
                    '    'Good To Go
                    'Else
                    '    PrintLine_tLINE  += " [" + sCurrentBalance + " - " + sCharges + " - " + sInsurance + " - " + sSelf + " - " + sCredits + " =" + STR_format(Str(dTemp), "0.00").PadLeft(10) + "]" + " {" + STR_format(Str(dTemp2), "0.00").PadLeft(10) + "}"
                    'End If
                    '*********************************************************
                    PrintLine_tLINE += vbCrLf
                Case "17", "18", "19", "20", "21", "22", "26"
                    '*********************************************************
                    '* DOS
                    PrintLine_tLINE += STR_TRIM(STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISM", LetterMultiLine_iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISD", LetterMultiLine_iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISY", LetterMultiLine_iLine).Trim, "mm/dd/yy"), 9).PadRight(9)
                    '*********************************************************
                    '* 2014-06-19 RFK: Specific Facility
                    Select Case sLetterType
                        Case "17"   'Facility
                            tFacility = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAFACL", LetterMultiLine_iLine).Trim
                            iFacilityRow = rkutils.DataGridView_Contains(DataGridView_Facilities, "FARFID", tFacility)
                            If iFacilityRow >= 0 Then
                                PrintLine_tLINE += STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_Facilities, "FANAME", iFacilityRow).Trim, 16).PadRight(16)
                            Else
                                MsgStatus(tFacility + " ERROR/", True)
                            End If
                        Case "18"   'Rendering Dr.  (Or Facility for Hospital) 
                            tFacility = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAFACL", LetterMultiLine_iLine).Trim
                            Select Case Val(tFacility)
                                Case 1
                                    PrintLine_tLINE += STR_TRIM("COOPER UNIV HOSP", 16).PadRight(16)
                                Case Else
                                    PrintLine_tLINE += STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAPRNM", LetterMultiLine_iLine).Trim, 16).PadRight(16)
                            End Select
                        Case "19", "20", "26"
                            tFacility = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAFACL", LetterMultiLine_iLine).Trim
                            iFacilityRow = rkutils.DataGridView_Contains(DataGridView_Facilities, "FARFID", tFacility)
                            If iFacilityRow >= 0 Then
                                PrintLine_tLINE += STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_Facilities, "FANAME", iFacilityRow).Trim, 16).PadRight(16)
                            Else
                                MsgStatus(tFacility + " ERROR/", True)
                            End If
                        Case "21", "22"   'Read the Charge Record for the description
                            sSQL = "SELECT DISTINCT DIGITS(CHG.ERTRYR)||'/'||DIGITS(CHG.ERTRMO)||'/'||DIGITS(CHG.ERTRDY) AS Date"
                            sSQL += ",ERUNIT"
                            sSQL += ",ERFR30"
                            sSQL += ",ERCHRG"
                            sSQL += " FROM ROIDATA.ERCHGP CHG"
                            sSQL += " WHERE ERLOCX='" + ReadField(DataGridView_Multi, "RALOCX", LetterMultiLine_iLine) + "'"
                            sSQL += " ORDER BY DIGITS(CHG.ERTRYR)||'/'||DIGITS(CHG.ERTRMO)||'/'||DIGITS(CHG.ERTRDY) "
                            SQL_READ_FIELD(DataGridView3, "DB2", "*", DB2SQLConnectionString, DB2SQLuser, sSQL)
                            PrintLine_tLINE += STR_TRIM(rkutils.ReadField(DataGridView3, "ERFR30", 0), 16).PadRight(16)
                    End Select
                    PrintLine_tLINE += " "
                    '*********************************************************
                    '* Patient
                    If bGhost Then
                        PrintLine_tLINE += STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_Ghosts, "FNAME", iGhosts_CTR).Trim + " " + rkutils.DataGridView_ValueByColumnName(DataGridView_Ghosts, "LNAME", iGhosts_CTR).Trim, 13).PadRight(13).ToUpper
                    Else
                        PrintLine_tLINE += STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAPNAM", LetterMultiLine_iLine).Trim, 13).PadRight(13)
                    End If
                    PrintLine_tLINE += " "
                    '*********************************************************
                    '* Account Number
                    '* 2017-07-14 RFK: Changed to 15 Account Number Length
                    If bGhost Then
                        PrintLine_tLINE += ("9999999999" + STR_RIGHT(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAACCT", LetterMultiLine_iLine).Trim, 4)).PadRight(15)
                    Else
                        PrintLine_tLINE += STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAACCT", LetterMultiLine_iLine).Trim, 15).PadRight(15)
                    End If
                    '* 2017-07-14 RFK: PrintLine_tLINE += " "
                    Select Case sLetterType
                        Case "21", "22"
                            '**************************************************
                            sCharges = rkutils.ReadField(DataGridView_Multi, "RAOBAL", LetterMultiLine_iLine).Trim
                            sCredits = rkutils.ReadField(DataGridView_Multi, "RAOADJ", LetterMultiLine_iLine).Trim
                            sSelf = rkutils.ReadField(DataGridView_Multi, "RAOPD", LetterMultiLine_iLine).Trim
                            sCurrentBalance = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RABALD", LetterMultiLine_iLine).Trim
                            '**************************************************
                        Case "26"
                            '**************************************************
                            '* 2021-08-31 RFK: RA FIELDS [not calculating from CXTRNP]
                            sCharges = rkutils.ReadField(DataGridView_Multi, "RAOBAL", LetterMultiLine_iLine).Trim
                            sCredits = rkutils.STR_format(Str(Val(rkutils.ReadField(DataGridView_Multi, "RAOADJ", LetterMultiLine_iLine).Trim) + Val(rkutils.ReadField(DataGridView_Multi, "RATOTA", LetterMultiLine_iLine).Trim)), "0.00")
                            sSelf = rkutils.STR_format(Str(Val(rkutils.ReadField(DataGridView_Multi, "RAOPD", LetterMultiLine_iLine).Trim) + Val(rkutils.ReadField(DataGridView_Multi, "RATOTP", LetterMultiLine_iLine).Trim)), "0.00")
                            sCurrentBalance = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RABALD", LetterMultiLine_iLine).Trim
                            '**************************************************
                        Case Else
                            '**************************************************
                            '* Charges
                            sCharges = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAOBAL", LetterMultiLine_iLine).Trim
                            '**************************************************
                            '* Adjustment
                            sSQL = "SELECT SUM(A.CXAMT) AS TSUM FROM"
                            sSQL += "("
                            sSQL += "SELECT DISTINCT LEFT(CXTYPE,1),CXDATE,CXAMT,CXDESC"
                            '*****************************************************************************************************************
                            '* 2014-06-19 RFK:
                            Select Case rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "ELETTERS", Val(Label_ClientRow.Text))
                                Case "t"
                                    sSQL += " FROM ROITEST.CXTRNP"
                                Case Else
                                    sSQL += " FROM ROIDATA.CXTRNP"
                            End Select
                            sSQL += " WHERE CXLOCX='" + ReadField(DataGridView_Multi, "RALOCX", LetterMultiLine_iLine) + "'"
                            sSQL += " AND LEFT(CXTYPE,1)='A'"
                            sSQL += " GROUP BY LEFT(CXTYPE,1),CXDATE,CXDESC,CXAMT"
                            sSQL += ") A"
                            sCredits = SQL_READ_FIELD(DataGridView3, "DB2", "TSUM", DB2SQLConnectionString, DB2SQLuser, sSQL)
                            '*********************************************************************************
                            '* 2014-07-16 RFK: 
                            dBalance = Val(sCredits)
                            dBalance += Val(ReadField(DataGridView_Multi, "RATOTA", LetterMultiLine_iLine))
                            'raoadj
                            sCredits = Trim(Str(dBalance))
                            '*********************************************************
                            '* Payments
                            sSQL = "SELECT SUM(A.CXAMT) AS TSUM FROM"
                            sSQL += "("
                            sSQL += "SELECT DISTINCT LEFT(CXTYPE,1),CXDATE,CXAMT,CXDESC"
                            '*****************************************************************************************************************
                            '* 2014-06-19 RFK:
                            Select Case rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "ELETTERS", Val(Label_ClientRow.Text))
                                Case "t"
                                    sSQL += " FROM ROITEST.CXTRNP"
                                Case Else
                                    sSQL += " FROM ROIDATA.CXTRNP"
                            End Select
                            sSQL += " WHERE CXLOCX='" + ReadField(DataGridView_Multi, "RALOCX", LetterMultiLine_iLine) + "'"
                            sSQL += " AND LEFT(CXTYPE,1)='P'"
                            sSQL += " GROUP BY LEFT(CXTYPE,1),CXDATE,CXDESC,CXAMT"
                            sSQL += ") A"
                            sSelf = SQL_READ_FIELD(DataGridView3, "DB2", "TSUM", DB2SQLConnectionString, DB2SQLuser, sSQL)
                            '**************************************************
                            '* 2014-07-16 RFK: 
                            dBalance = Val(sSelf)
                            dBalance += Val(ReadField(DataGridView_Multi, "RATOTP", LetterMultiLine_iLine))
                            sSelf = Trim(Str(dBalance))
                            '**************************************************
                            '* Balance Due
                            sCurrentBalance = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RABALD", LetterMultiLine_iLine).Trim
                            '**************************************************
                            '* Calculated Total
                            dTotal = Val(sCharges)
                            dTotal += Val(sCredits)
                            dTotal += Val(sSelf)
                            If STR_format(dTotal.ToString, "0.00").PadLeft(10) <> STR_format(sCurrentBalance, "0.00").PadLeft(10) Then
                                sCredits = (dTotal - Val(sCurrentBalance))
                            End If
                    End Select
                    '**********************************************************
                    PrintLine_tLINE += STR_format(sCharges, "0.00").PadLeft(10)
                    PrintLine_tLINE += STR_format(sCredits, "0.00").PadLeft(10)         'Adjust
                    PrintLine_tLINE += STR_format(sSelf, "0.00").PadLeft(10)            'Payments
                    PrintLine_tLINE += STR_format(sCurrentBalance, "0.00").PadLeft(10)
                    '**********************************************************
                    PrintLine_tLINE += vbCrLf
                    '**********************************************************
                    '* 2014-06-19 RFK:
                    Select Case sLetterType
                        Case "17"   '2nd Line
                            '**************************************************
                            '* 2014-09-30 RFK: 
                            sSQL = "SELECT DISTINCT DIGITS(CHG.ERTRMO)||'/'||DIGITS(CHG.ERTRDY)||'/'||DIGITS(CHG.ERTRYR) AS TDate"
                            sSQL += ",ERCHRG AS AMOUNT"
                            sSQL += ",ERLSRV AS LOCATION"
                            sSQL += ",ERTSRV"
                            sSQL += ",ERUNIT"
                            sSQL += ",ERFR30"
                            sSQL += " FROM ROIDATA.ERCHGP CHG"
                            sSQL += " WHERE CHG.ERLOCX='" + LetterMultiLine_LOCX + "'"
                            sSQL += " ORDER BY DIGITS(CHG.ERTRMO)||'/'||DIGITS(CHG.ERTRDY)||'/'||DIGITS(CHG.ERTRYR)"
                            SQL_READ_FIELD(DataGridView3, "DB2", "*", DB2SQLConnectionString, DB2SQLuser, sSQL)
                            For i1 = 0 To Me.DataGridView3.RowCount - 1
                                If rkutils.DataGridView_ValueByColumnName(DataGridView3, "ERUNIT", i1).Length > 0 Then
                                    '********************************************************
                                    '* Only 36 MULTI LINES
                                    If MatchedPrint_iPrintedLines < 36 Then
                                        PrintLine_tLINE += STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView3, "TDATE", i1), "mm/dd/yy")
                                        PrintLine_tLINE += " ".PadRight(3)
                                        PrintLine_tLINE += STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView3, "ERFR30", i1), 28).Trim.PadRight(28)
                                        PrintLine_tLINE += " QTY: "
                                        PrintLine_tLINE += STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView3, "ERUNIT", i1), 7).Trim.PadRight(7)
                                        PrintLine_tLINE += " "
                                        PrintLine_tLINE += STR_TRIM(STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView3, "AMOUNT", i1), "0.00"), 12).Trim.PadLeft(12)
                                        PrintLine_tLINE += vbCrLf
                                    End If
                                    '********************************************************
                                    '* IF more than 36 MULTI LINES
                                    MatchedPrint_iPrintedLines += 1
                                    If MatchedPrint_iPrintedLines >= 36 Then Exit For
                                End If
                            Next
                        Case "19", "20", "26" '2nd Line
                            '*********************************************************
                            '* 2015-08-05 RFK: 
                            '* 2015-08-06 RFK: changed to ERFRMO / ERFRDY / ERFRYR sSQL = "SELECT DISTINCT DIGITS(CHG.ERTRMO)||'/'||DIGITS(CHG.ERTRDY)||'/'||DIGITS(CHG.ERTRYR) AS TDate"
                            sSQL = "SELECT DISTINCT DIGITS(CHG.ERFRMO)||'/'||DIGITS(CHG.ERFRDY)||'/'||DIGITS(CHG.ERFRYR) AS TDate"
                            sSQL += ",CHG.ERCHRG AS AMOUNT"
                            sSQL += ",CHG.ERLSRV AS LOCATION"
                            sSQL += ",CHG.ERTSRV"
                            sSQL += ",CHG.ERUNIT"
                            sSQL += ",CHG.ERFR30"
                            sSQL += ",CHG.ERCPT#"
                            sSQL += ",CPT.CPDESC"
                            sSQL += ",CPT.CPLDESC"
                            sSQL += " FROM ROIDATA.ERCHGP CHG"
                            sSQL += " LEFT JOIN ROIDATA.CPTTRP CPT ON CHG.ERCPT#=CPT.CPCPT#"
                            sSQL += " WHERE CHG.ERLOCX='" + LetterMultiLine_LOCX + "'"
                            sSQL += " ORDER BY DIGITS(CHG.ERFRMO)||'/'||DIGITS(CHG.ERFRDY)||'/'||DIGITS(CHG.ERFRYR)"
                            If rkutils.SQL_READ_DATAGRID(DataGridView3, "DB2", "*", DB2SQLConnectionString, DB2SQLuser, sSQL) Then
                                For i1 = 0 To DataGridView3.RowCount - 1
                                    '**********************************************
                                    '* 2015-08-05 RFK: 
                                    If rkutils.ReadField(DataGridView3, "ERUNIT", i1).Length > 0 Then
                                        '********************************************************
                                        '* Only 36 MULTI LINES
                                        If MatchedPrint_iPrintedLines < 36 Then
                                            PrintLine_tLINE += STR_TRIM(rkutils.ReadField(DataGridView3, "ERCPT#", i1), 9).Trim.PadRight(8)
                                            Select Case sLetterType
                                                Case "20"
                                                    PrintLine_tLINE += " "
                                                Case Else
                                                    PrintLine_tLINE += "QTY:"
                                                    PrintLine_tLINE += STR_TRIM(rkutils.ReadField(DataGridView3, "ERUNIT", i1), 7).Trim.PadRight(5)
                                            End Select
                                            sTmpSTR = rkutils.ReadField(DataGridView3, "CPLDESC", i1).Trim
                                            If sTmpSTR.Length > 0 Then
                                                PrintLine_tLINE += sTmpSTR
                                            Else
                                                PrintLine_tLINE += STR_TRIM(rkutils.ReadField(DataGridView3, "ERFR30", i1), 28).Trim.PadRight(31)
                                            End If
                                            PrintLine_tLINE += vbCrLf
                                        End If
                                        '********************************************************
                                        '* IF more than 36 MULTI LINES
                                        MatchedPrint_iPrintedLines += 1
                                        If MatchedPrint_iPrintedLines >= 36 Then Exit For
                                    End If
                                Next
                            End If
                        Case "21"   '2nd Line
                            'PrintLine_tLINE += vbCrLf
                    End Select
                Case "23", "24", "25", "29"
                    'Should Never Make it HERE
            End Select
            Return PrintLine_tLINE
        Catch ex As Exception
            MsgError("LetterMultLetterMultiLine_iLine", ex.ToString)
        End Try
        Return ""
    End Function

    Private Function LettersPrinted(ByVal sLetter As String, ByVal sLOCX As String, ByVal bError As Boolean) As Boolean
        Try
            '****************************************************************
            '* 2013-11-26 RFK:
            '* 2014-08-01 RFK:
            If bError Then
                iLetterRow = Listbox_Contains(ListBox_Letters, "ERR_" + sLetter, False)
                File.AppendAllText(dir_REPORTS + "LETTERS\" + rkutils.STR_format("TODAY", "ccyymmdd") + "_" + rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)) + "_ERR_" + sLetter + ".XLS", sLOCX + vbTab + "http:\\enhance\display_account.aspx?IDR=" + sLOCX + "" + vbCrLf)
            Else
                iLetterRow = Listbox_Contains(ListBox_Letters, sLetter, False)
            End If
            If iLetterRow >= 0 Then
                ListBox_Letters.Items(iLetterRow) = STR_BREAK(ListBox_Letters.Items(iLetterRow).ToString, 1) + " " + Str(Val(STR_BREAK(ListBox_Letters.Items(iLetterRow).ToString, 2)) + 1).Trim
            Else
                If bError Then
                    ListBox_Letters.Items.Add("ERR_" + sLetter + " 1")
                    MsgStatus("ERROR:" + sLetter + " " + sLOCX, True)
                Else
                    ListBox_Letters.Items.Add(sLetter + " 1")
                End If
            End If
            Return True
        Catch ex As Exception
            MsgError("LettersError", ex.ToString)
        End Try
        Return False
    End Function

    Private Sub ReadLine(ByVal iLine As Integer)
        Try
            If CheckBox_DEBUG.Checked Then MsgStatus("ReadLine:" + iLine.ToString, True)
            '******************************************************************
            '* 2019-03-14 RFK:
            If swDTable Then
                TextBox_Locx.Text = rkutils.DataTable_ValueByColumnName(dTable_Select, "RALOCX", iLine)
                tLOCX = rkutils.DataTable_ValueByColumnName(dTable_Select, "RALOCX", iLine)
                tRamLOCX = rkutils.DataTable_ValueByColumnName(dTable_Select, "RAMLOCX", iLine)
                sRAACCT = rkutils.DataTable_ValueByColumnName(dTable_Select, "RAACCT", iLine)
                tGSSN = rkutils.DataTable_ValueByColumnName(dTable_Select, "RAGSS#", iLine)
                tMedRec = rkutils.DataTable_ValueByColumnName(dTable_Select, "RAMR#", iLine)
                tGNAMEL = rkutils.DataTable_ValueByColumnName(dTable_Select, "RAGLNM", iLine)
                tGNAMEF = rkutils.DataTable_ValueByColumnName(dTable_Select, "RAGFNM", iLine)
                tGNAMEM = rkutils.DataTable_ValueByColumnName(dTable_Select, "RAGMI", iLine)
                tGADDR = rkutils.DataTable_ValueByColumnName(dTable_Select, "RAGADD", iLine)
                tGCITY = rkutils.DataTable_ValueByColumnName(dTable_Select, "RAGCITY", iLine)
                tGZIP = rkutils.DataTable_ValueByColumnName(dTable_Select, "RAGZIP", iLine).PadLeft(5, "0")
                tPSSN = rkutils.DataTable_ValueByColumnName(dTable_Select, "RAGSS#", iLine)
                tPNAMEL = rkutils.DataTable_ValueByColumnName(dTable_Select, "RAPNAM", iLine)
                sRalNAC = rkutils.DataTable_ValueByColumnName(dTable_Select, "RALNAC", iLine)
                sDOS = STR_TRIM(STR_format(rkutils.DataTable_ValueByColumnName(dTable_Select, "RADISM", iLine).Trim + "/" + rkutils.DataTable_ValueByColumnName(dTable_Select, "RADISD", iLine).Trim + "/" + rkutils.DataTable_ValueByColumnName(dTable_Select, "RADISY", iLine).Trim, "mm/dd/ccyy"), 12).PadRight(12)
                sRABALD = rkutils.DataTable_ValueByColumnName(dTable_Select, "RABALD", iLine)
                sPlacementDate = rkutils.DataTable_ValueByColumnName(dTable_Select, "PlacementDate", iLine)
                sPlacementDate1 = rkutils.DataTable_ValueByColumnName(dTable_Select, "PlacementDate1", iLine)
            Else
                TextBox_Locx.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RALOCX", iLine)
                tLOCX = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RALOCX", iLine)
                tRamLOCX = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAMLOCX", iLine)
                sRAACCT = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAACCT", iLine)
                tGSSN = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGSS#", iLine)
                tMedRec = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAMR#", iLine)
                tGNAMEL = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGLNM", iLine)
                tGNAMEF = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGFNM", iLine)
                tGNAMEM = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGMI", iLine)
                tGADDR = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGADD", iLine)
                tGCITY = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGCITY", iLine)
                tGZIP = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGZIP", iLine).PadLeft(5, "0")
                tPSSN = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGSS#", iLine)
                tPNAMEL = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAPNAM", iLine)
                sRalNAC = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RALNAC", iLine)
                sDOS = STR_TRIM(STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RADISM", iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RADISD", iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RADISY", iLine).Trim, "mm/dd/ccyy"), 12).PadRight(12)
                sRABALD = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RABALD", iLine)
                sPlacementDate = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "PlacementDate", iLine)
                sPlacementDate1 = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "PlacementDate1", iLine)
            End If
            '**************************************************************
            '* 2021-07-12 RFK: RALNAC
            gtLetterType = Letter_Types_Value(sRalNAC, "LTYPE")
        Catch ex As Exception
            MsgError("ReadLine", ex.ToString)
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Function MatchedPrint(ByVal iLine As Integer, ByVal tLetterVendor As String, ByVal tLetterNumber As String, ByVal sLetterType As String, ByVal tLOCX As String, ByVal bGhost As Boolean) As String
        Try
            '************************************************
            '* 2012-06-21 RFK:
            If CheckBox_DEBUG.Checked Or bGhost Then MsgStatus("MatchedPrint:" + iLine.ToString, True)
            PrintLine_tLINE = ""
            '*************************************************
            '* 2012-08-09 RFK:
            Select Case tLetterVendor
                Case "APEX"
                    '
                Case "ACCUDOC", "DANTOM", "REVSPRING"
                    sLetterType = Letter_Types_Value(tLetterNumber, "LTYPE")
                    dBalance = 0    'Reset for Each Matched Letter
                    Select Case sLetterType
                        Case "10"
                            '**************************************************
                            '* 2021-12-07 RFK:
                            iField1Pad = 20
                            iField2Pad = 20
                            iField3Pad = 20
                            PrintLine_tLINE += LetterMultiHeader(sLetterType)
                            sRAMLOCX = tRamLOCX
                            sPlacementDate1S = sPlacementDate1
                            Do While tRamLOCX = sRAMLOCX And sPlacementDate1 = sPlacementDate1S
                                '**************************************************
                                '* 2021-11-26 RFK: REG F LETTER 1
                                PrintLine_tLINE += STR_TRIM(sRAACCT, iField1Pad).PadRight(iField1Pad)
                                '**************************************************
                                '* DOS
                                PrintLine_tLINE += STR_TRIM(sDOS, iField2Pad).PadRight(iField2Pad)
                                '**************************************************
                                '* Balance Due
                                PrintLine_tLINE += STR_TRIM(STR_format(Str(Val(sRABALD)).Trim, "$"), iField3Pad).PadLeft(iField3Pad)
                                '**************************************************
                                '* Interest
                                PrintLine_tLINE += STR_TRIM(STR_format(Str(Val("0")).Trim, "$"), iField3Pad).PadLeft(iField3Pad)
                                '**************************************************
                                '* Fees
                                PrintLine_tLINE += STR_TRIM(STR_format(Str(Val("0")).Trim, "$"), iField3Pad).PadLeft(iField3Pad)
                                '**************************************************
                                '* Payments/Credits
                                PrintLine_tLINE += STR_TRIM(STR_format(Str(Val("0")).Trim, "$"), iField3Pad).PadLeft(iField3Pad)
                                '**************************************************
                                '* Amount Due
                                PrintLine_tLINE += STR_TRIM(STR_format(Str(Val(sRABALD)).Trim, "$"), iField3Pad).PadLeft(iField3Pad)
                                dBalance += Val(sRABALD)
                                '**************************************************
                                '*2021-12-13 RFK: FOR TESTING
                                'PrintLine_tLINE += " " + STR_TRIM(sPlacementDate1, iField3Pad).PadLeft(iField3Pad)
                                '**************************************************
                                PrintLine_tLINE += vbCrLf
                                '**************************************************
                                '* 2021-12-13 RFK:
                                sRAMLOCX = ""
                                If swDTable Then
                                    If iLine + 1 <= dTable_Select.Rows.Count Then
                                        sRAMLOCX = rkutils.DataTable_ValueByColumnName(dTable_Select, "RAMLOCX", iLine + 1)
                                        sPlacementDate1S = rkutils.DataTable_ValueByColumnName(dTable_Select, "PlacementDate1", iLine + 1)
                                        If tRamLOCX = sRAMLOCX And sPlacementDate1 = sPlacementDate1S Then
                                            iLine += 1
                                            ReadLine(iLine)
                                            rkutils.DataTable_SetValueByColumnName(dTable_Select, "ERRORCODE", iLine, "MULTI")
                                            letter_sent("MatchedPrint." + sLetterType, False, iLine, tLetterNumber, sSentBalance, True)
                                        End If
                                    End If
                                Else
                                    If iLine + 1 <= DataGridView_Select.Rows.Count Then
                                        sRAMLOCX = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAMLOCX", iLine + 1)
                                        sPlacementDate1S = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "PlacementDate1", iLine + 1)
                                        If tRamLOCX = sRAMLOCX And sPlacementDate1 = sPlacementDate1S Then
                                            iLine += 1
                                            ReadLine(iLine)
                                            rkutils.DataGridView_SetValueByColumnName(DataGridView_Select, "ERRORCODE", iLine, "MULTI")
                                            letter_sent("MatchedPrint." + sLetterType, False, iLine, tLetterNumber, sSentBalance, True)
                                        End If
                                    End If
                                End If
                            Loop
                            '**************************************************
                            PrintLine_tLINE += LetterMultiTrailer(sLetterType)
                            '**************************************************
                            Return PrintLine_tLINE
                        Case "3"
                            iField1Pad = 25
                            iField2Pad = 12
                            sSQL = "SELECT SUBSTRING(F.FRREVC,2,2) AS FRREVC, COUNT(F.FRREVC) AS COUNT, SUM(F.FRCHRG) AS SUM, '' AS DESCRIPTION FROM ROITEST.FREVCP F"
                            sSQL += " WHERE F.FRLOCX=" + tLOCX
                            sSQL += " GROUP BY SUBSTRING(F.FRREVC,2,2)"
                            sSQL += " ORDER BY SUBSTRING(F.FRREVC,2,2)"
                            If SQL_READ_DATAGRID(DataGridView2, "DB2", "*", DB2SQLConnectionString, DB2SQLuser, sSQL) Then
                                'Me.DataGridView2.Visible = True
                                '**********************************************
                                For i1 = 0 To Me.DataGridView2.RowCount - 1
                                    tSTR = rkutils.DataGridView_ValueByColumnName(DataGridView2, "FRREVC", i1).Trim
                                    If tSTR.Length > 0 Then
                                        tSUM = rkutils.DataGridView_ValueByColumnName(DataGridView2, "SUM", i1).Trim
                                        tDESCR = SQL_READ_FIELD(DataGridView3, "MSSQL", "DESCRIPTION", msSQLConnectionString, msSQLuser, "SELECT DESCRIPTION FROM " + sDBO + ".dbo.RevenueCodes WHERE SUBSTRING(REV_CODE, 2, 2)='" + tSTR + "'")
                                        If tDESCR.Length = 0 Then tDESCR = tSTR + "-Other"
                                        '*************************************************
                                        '* 2012-08-09 RFK:
                                        Select Case tLetterNumber
                                            Case Else
                                                PrintLine_tLINE += STR_RIGHT(tDESCR, iField1Pad).PadRight(iField1Pad) + " " + STR_format(tSUM, "$").PadLeft(iField2Pad) + vbCrLf
                                        End Select
                                    End If
                                Next
                            End If
                        Case Else
                            '**************************************************
                            iField1Pad = 20
                            iField2Pad = 20
                            iField3Pad = 20
                            If DataGridView_Multi.RowCount > 0 Then
                                '**********************************************
                                '* 2012-10-01 RFK: 
                                sSysAccountMatched = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAMLOCX", iLine).Trim
                                imLocxRow = rkutils.DataGridView_Contains(DataGridView_Multi, "RAMLOCX", sSysAccountMatched)
                                MatchedPrint_iMatchedRow = imLocxRow
                                MatchedPrint_iPrintedLines = 0
                                dSummaryLine = 0
                                If bGhost Then MsgStatus("GHOST/MatchedPrint:" + sSysAccountMatched + "][" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAMLOCX", MatchedPrint_iMatchedRow).Trim + "]", True)
                                Do While MatchedPrint_iMatchedRow <= DataGridView_Multi.RowCount - 1 And sSysAccountMatched = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAMLOCX", MatchedPrint_iMatchedRow).Trim
                                    tRALOCX = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RALOCX", MatchedPrint_iMatchedRow).Trim
                                    If bGhost Then MsgStatus("RALOCX:" + tRALOCX, True)
                                    '******************************************
                                    '* 2015-08-05 RFK: 
                                    '* 2013-01-07 RFK:
                                    '******************************************
                                    If CheckBox_DEBUG.Checked Then MsgStatus("MatchedPrint MatchedPrint_iMatchedRow:" + MatchedPrint_iMatchedRow.ToString + " mRLocx:" + tRALOCX + " mRamLocx" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAMLOCX", MatchedPrint_iMatchedRow).Trim, True)
                                    '******************************************
                                    If Val(tRALOCX) > 0 Then
                                        '**************************************
                                        '* Header Row
                                        If MatchedPrint_iMatchedRow = imLocxRow Then PrintLine_tLINE += LetterMultiHeader(sLetterType)
                                        '**************************************
                                        Select Case sLetterType
                                            Case "15"
                                                '* Only Print 37 MULTI LINES  (2 for title bar)
                                                If MatchedPrint_iPrintedLines < 36 Then
                                                    '**************************
                                                    '* 2016-05-10 RFK: NO 0 BALANCE
                                                    If Val(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RABALD", MatchedPrint_iMatchedRow).Trim) <> 0 Then
                                                        If Listbox_Contains(ListBox_Printed, tLOCX, False) < 0 Then
                                                            PrintLine_tLINE += LetterMultiLine(sLetterType, MatchedPrint_iMatchedRow, tLOCX, bGhost)
                                                            '******************
                                                            MatchedPrint_iPrintedLines += 1
                                                            '******************
                                                        End If
                                                    End If
                                                End If
                                            Case "17", "19", "20", "21", "22", "26"
                                                '* Only Print 37 MULTI LINES  (2 for title bar)
                                                If MatchedPrint_iPrintedLines < 36 Then
                                                    '**************************
                                                    '* 2016-04-05 RFK: added 17 to the tRALOCX
                                                    If Listbox_Contains(ListBox_Printed, tLOCX, False) < 0 Then
                                                        '**********************
                                                        PrintLine_tLINE += LetterMultiLine(sLetterType, MatchedPrint_iMatchedRow, tRALOCX, bGhost)
                                                        MatchedPrint_iPrintedLines += 1
                                                        '**********************
                                                    End If
                                                End If
                                            Case "23", "24", "25", "29"
                                                '******************************
                                                '* 2021-07-19 RFK: 
                                                '* NOTHING
                                            Case Else
                                                '******************************
                                                '* 2021-07-19 RFK: 
                                                '******************************************************************
                                                '* 2021-11-29 RFK:
                                                If Listbox_Contains(ListBox_Printed, tRALOCX, True) < 0 Then
                                                    PrintLine_tLINE += LetterMultiLine(sLetterType, MatchedPrint_iMatchedRow, tRALOCX, bGhost)
                                                End If
                                                '******************************
                                                If swDTable Then
                                                    'Select Case ReadFieldDataTable(dTable_Select, "ERRORCODE", iLocxRow)
                                                    '    Case "MULTI"
                                                    '        MsgStatus("MULTI_ALREADY/" + ReadFieldDataTable(dTable_Select, "RALOCX", iLocxRow), False)
                                                    '    Case Else
                                                    '        MsgStatus("FAILED/" + ReadFieldDataTable(dTable_Select, "ERRORCODE", iLocxRow) + "/" + ReadFieldDataTable(dTable_Select, "RALOCX", iLocxRow), False)
                                                    'End Select
                                                Else
                                                    '****************************************************************
                                                    '* 2012-08-21 RFK: so does not print again
                                                    iLocxMultiRow = rkutils.DataGridView_Contains(DataGridView_Select, "RALOCX", tRALOCX)
                                                    If iLocxMultiRow >= 0 Then
                                                        If iLocxMultiRow = MatchedPrint_iMatchedRow Then
                                                            Printed(iLocxMultiRow, tLetterPrinted, "PRINTED")
                                                        Else
                                                            Printed(iLocxMultiRow, tLetterPrinted, "MULTI")
                                                        End If
                                                    End If
                                                End If
                                                '******************************
                                                MatchedPrint_iPrintedLines += 1
                                                '******************************
                                        End Select
                                        '**************************************
                                        '* 2012-11-12 RFK:
                                        '* 2018-03-07 RFK: DO FOR ALL ACCOUNTS
                                        If bGhost = False Then letter_sent("MatchedPrint.2", True, MatchedPrint_iMatchedRow, tLetterNumber, sSentBalance, True)
                                    End If
                                    MatchedPrint_iMatchedRow += 1
                                Loop
                                '**********************************************
                                '* 2021-07-19 RFK:
                                Select Case sLetterType
                                    Case "29"
                                        '**********************************************************
                                        '* 2022-03-09 RFK: 
                                        sSQL = "SELECT DIGITS(A.RAADMM)||'/'||DIGITS(A.RAADMD)||'/'||DIGITS(A.RAADMY) AS DOS"
                                        sSQL += ", C.CXDESC AS DESCRIPTION"
                                        sSQL += ", C.CXAMT AS AMOUNT"
                                        'sSQL += ", SUM(C.CXAMT) AS Total"
                                        sSQL += " FROM ROIDATA.CXTRNP C"
                                        sSQL += " LEFT JOIN ROIDATA.RACCTP A ON C.CXLOCX = A.RALOCX"
                                        sSQL += " WHERE C.CXLOCX IN('" + MultiAccountsLIST(iLine).Replace(" ", "','") + "')"
                                        sSQL += " GROUP BY DIGITS(A.RAADMM)||'/'||DIGITS(A.RAADMD)||'/'||DIGITS(A.RAADMY)"
                                        sSQL += ",C.CXDESC, C.CXAMT"
                                        sSQL += " ORDER BY DIGITS(A.RAADMM)||'/'||DIGITS(A.RAADMD)||'/'||DIGITS(A.RAADMY)"
                                        MsgStatus(sSQL, False)
                                        '**********************************************************
                                        If rkutils.SQL_READ_DATAGRID(DataGridView3, "DB2", "*", DB2SQLConnectionString, DB2SQLuser, sSQL) Then
                                            dSummaryLine = 0
                                            '******************************************************
                                            For i1 = 0 To DataGridView3.RowCount - 1
                                                '**************************************************
                                                If rkutils.ReadField(DataGridView3, "Description", i1).Length > 1 Then
                                                    If MatchedPrint_iPrintedLines < 36 Then
                                                        PrintLine_tLINE += STR_format(rkutils.ReadField(DataGridView3, "DOS", i1), "mm/dd/ccyy")
                                                        PrintLine_tLINE += " "
                                                        PrintLine_tLINE += STR_TRIM(rkutils.ReadField(DataGridView3, "Description", i1), 30).Trim.PadRight(30)
                                                        PrintLine_tLINE += " "
                                                        PrintLine_tLINE += STR_format(rkutils.ReadField(DataGridView3, "Amount", i1), "$")
                                                        PrintLine_tLINE += vbCrLf
                                                        dSummaryLine = 0
                                                    Else
                                                        'PrintLine_tLINE += "" + STR_format(rkutils.ReadField(DataGridView3, "Total", i1), "$")
                                                        'PrintLine_tLINE += vbCrLf
                                                        dSummaryLine += Val(rkutils.ReadField(DataGridView3, "Amount", i1))
                                                    End If
                                                    MatchedPrint_iPrintedLines += 1
                                                End If
                                            Next
                                        End If
                                        '**********************************************************
                                        '* 2022-03-09 RFK: 
                                        sSQL = "SELECT DIGITS(A.RAADMM)||'/'||DIGITS(A.RAADMD)||'/'||DIGITS(A.RAADMY) AS DOS"
                                        sSQL += ", F.FRFR30 AS DESCRIPTION"
                                        sSQL += ", F.FRCHRG AS AMOUNT"
                                        sSQL += " FROM ROIDATA.FREVCP F"
                                        sSQL += " LEFT JOIN ROIDATA.RACCTP A ON F.FRLOCX = A.RALOCX"
                                        sSQL += " WHERE F.FRLOCX IN('" + MultiAccountsLIST(iLine).Replace(" ", "','") + "')"
                                        sSQL += " GROUP BY DIGITS(A.RAADMM)||'/'||DIGITS(A.RAADMD)||'/'||DIGITS(A.RAADMY)"
                                        sSQL += ", F.FRFR30, F.FRCHRG"
                                        sSQL += " ORDER BY DIGITS(A.RAADMM)||'/'||DIGITS(A.RAADMD)||'/'||DIGITS(A.RAADMY)"
                                        sSQL += ", F.FRFR30, F.FRCHRG"
                                        MsgStatus(sSQL, False)
                                        '**********************************************************
                                        If rkutils.SQL_READ_DATAGRID(DataGridView3, "DB2", "*", DB2SQLConnectionString, DB2SQLuser, sSQL) Then
                                            dSummaryLine = 0
                                            '******************************************************
                                            For i1 = 0 To DataGridView3.RowCount - 1
                                                '**************************************************
                                                If rkutils.ReadField(DataGridView3, "Description", i1).Length > 1 Then
                                                    If MatchedPrint_iPrintedLines < 36 Then
                                                        PrintLine_tLINE += STR_format(rkutils.ReadField(DataGridView3, "DOS", i1), "mm/dd/ccyy")
                                                        PrintLine_tLINE += " "
                                                        PrintLine_tLINE += STR_TRIM(rkutils.ReadField(DataGridView3, "Description", i1), 30).Trim.PadRight(30)
                                                        PrintLine_tLINE += " "
                                                        PrintLine_tLINE += STR_format(rkutils.ReadField(DataGridView3, "Amount", i1), "$")
                                                        PrintLine_tLINE += vbCrLf
                                                        dSummaryLine = 0
                                                    Else
                                                        'PrintLine_tLINE += "" + STR_format(rkutils.ReadField(DataGridView3, "Total", i1), "$")
                                                        'PrintLine_tLINE += vbCrLf
                                                        dSummaryLine += Val(rkutils.ReadField(DataGridView3, "Amount", i1))
                                                    End If
                                                    MatchedPrint_iPrintedLines += 1
                                                End If
                                            Next
                                        End If
                                        If MatchedPrint_iPrintedLines >= 24 And dSummaryLine > 0 Then
                                            PrintLine_tLINE += "00/00/0000"
                                            PrintLine_tLINE += " "
                                            PrintLine_tLINE += "OTHER CHARGES".Trim.PadRight(30)
                                            PrintLine_tLINE += " "
                                            PrintLine_tLINE += STR_format(dSummaryLine.ToString, "$")
                                            PrintLine_tLINE += vbCrLf
                                        End If
                                    Case "23"
                                        '**************************************
                                        '* 2021-05-27 RFK: 
                                        sSQL = "SELECT DIGITS(RAADMM)||'/'||DIGITS(RAADMD)||'/'||DIGITS(RAADMY) AS DOS"
                                        sSQL += ", FRFR30 AS DESCRIPTION"
                                        sSQL += ", SUM(FRCHRG) AS Total"
                                        sSQL += " FROM ROIDATA.FREVCP F"
                                        sSQL += " LEFT JOIN ROIDATA.RACCTP A ON F.FRLOCX = A.RALOCX"
                                        sSQL += " WHERE FRLOCX IN('" + MultiAccountsLIST(iLine).Replace(" ", "','") + "')"
                                        sSQL += " GROUP BY DIGITS(RAADMM)||'/'||DIGITS(RAADMD)||'/'||DIGITS(RAADMY), FRFR30"
                                        sSQL += " ORDER BY DIGITS(RAADMM)||'/'||DIGITS(RAADMD)||'/'||DIGITS(RAADMY), FRFR30"
                                        MsgStatus(sSQL, False)
                                        If rkutils.SQL_READ_DATAGRID(DataGridView3, "DB2", "*", DB2SQLConnectionString, DB2SQLuser, sSQL) Then
                                            dSummaryLine = 0
                                            For i1 = 0 To DataGridView3.RowCount - 1
                                                '**************************************************
                                                If rkutils.ReadField(DataGridView3, "Description", i1).Length > 1 Then
                                                    If MatchedPrint_iPrintedLines < 36 Then
                                                        PrintLine_tLINE += STR_format(rkutils.ReadField(DataGridView3, "DOS", i1), "mm/dd/ccyy")
                                                        PrintLine_tLINE += " "
                                                        PrintLine_tLINE += STR_TRIM(rkutils.ReadField(DataGridView3, "Description", i1), 30).Trim.PadRight(30)
                                                        PrintLine_tLINE += " "
                                                        PrintLine_tLINE += STR_format(rkutils.ReadField(DataGridView3, "Total", i1), "$")
                                                        PrintLine_tLINE += vbCrLf
                                                        dSummaryLine = 0
                                                    Else
                                                        'PrintLine_tLINE += "" + STR_format(rkutils.ReadField(DataGridView3, "Total", i1), "$")
                                                        'PrintLine_tLINE += vbCrLf
                                                        dSummaryLine += Val(rkutils.ReadField(DataGridView3, "Total", i1))
                                                    End If
                                                    MatchedPrint_iPrintedLines += 1
                                                End If
                                            Next
                                        End If
                                        If MatchedPrint_iPrintedLines >= 36 And dSummaryLine > 0 Then
                                            PrintLine_tLINE += "00/00/0000"
                                            PrintLine_tLINE += " "
                                            PrintLine_tLINE += "OTHER CHARGES".Trim.PadRight(30)
                                            PrintLine_tLINE += " "
                                            PrintLine_tLINE += STR_format(dSummaryLine.ToString, "$")
                                            PrintLine_tLINE += vbCrLf
                                        End If
                                    Case "24", "25"
                                        '**************************************
                                        '* 2021-05-27 RFK: 
                                        sSQL = "SELECT FRFR30 AS DESCRIPTION"
                                        sSQL += ", SUM(FRCHRG) AS Total"
                                        sSQL += " FROM ROIDATA.FREVCP F"
                                        sSQL += " WHERE FRLOCX IN('" + MultiAccountsLIST(iLine).Replace(" ", "','") + "')"
                                        sSQL += " GROUP BY FRFR30"
                                        sSQL += " ORDER BY SUM(FRCHRG) DESC, FRFR30"
                                        MsgStatus(sSQL, False)
                                        If rkutils.SQL_READ_DATAGRID(DataGridView3, "DB2", "*", DB2SQLConnectionString, DB2SQLuser, sSQL) Then
                                            dSummaryLine = 0
                                            For i1 = 0 To DataGridView3.RowCount - 1
                                                '**************************************************
                                                If rkutils.ReadField(DataGridView3, "Description", i1).Length > 1 Then
                                                    If MatchedPrint_iPrintedLines < 9 Then
                                                        'PrintLine_tLINE += STR_format(rkutils.ReadField(DataGridView3, "DOS", i1), "mm/dd/ccyy")
                                                        PrintLine_tLINE += "00/00/0000"
                                                        PrintLine_tLINE += " "
                                                        PrintLine_tLINE += STR_TRIM(rkutils.ReadField(DataGridView3, "Description", i1), 30).Trim.PadRight(30)
                                                        PrintLine_tLINE += " "
                                                        PrintLine_tLINE += STR_format(rkutils.ReadField(DataGridView3, "Total", i1), "$")
                                                        PrintLine_tLINE += vbCrLf
                                                        dSummaryLine = 0
                                                    Else
                                                        'PrintLine_tLINE += "" + STR_format(rkutils.ReadField(DataGridView3, "Total", i1), "$")
                                                        'PrintLine_tLINE += vbCrLf
                                                        dSummaryLine += Val(rkutils.ReadField(DataGridView3, "Total", i1))
                                                    End If
                                                    MatchedPrint_iPrintedLines += 1
                                                End If
                                            Next
                                        End If
                                        If MatchedPrint_iPrintedLines >= 9 And dSummaryLine > 0 Then
                                            PrintLine_tLINE += "00/00/0000"
                                            PrintLine_tLINE += " "
                                            PrintLine_tLINE += "OTHER CHARGES".Trim.PadRight(30)
                                            PrintLine_tLINE += " "
                                            PrintLine_tLINE += STR_format(dSummaryLine.ToString, "$")
                                            PrintLine_tLINE += vbCrLf
                                        End If
                                    Case Else
                                        '* IF more than 36 MULTI LINES
                                        If MatchedPrint_iPrintedLines >= 36 Then
                                            PrintLine_tLINE += "PLEASE CONTACT US FOR INFORMATION ON ADDITIONAL ACCOUNTS" + vbCrLf
                                            MsgStatus(tLOCX + " contained " + MatchedPrint_iPrintedLines.ToString + " multilines; (more than 36)", True)
                                        End If
                                End Select
                            Else
                                MsgStatus("No Matched Accounts", True)
                            End If
                            If dBalance > 0 Then
                                PrintLine_tLINE += LetterMultiTrailer(sLetterType)
                            End If
                    End Select
            End Select
            '*******************************************************
            If CheckBox_DEBUG.Checked Then MsgStatus("MatchedPrint Done", True)
            Return PrintLine_tLINE
            '*******************************************************
        Catch ex As Exception
            MsgError("MatchedPrint", ex.ToString)
        End Try
        Return ""
    End Function

    Private Sub letter_sent(ByVal sModule As String, ByVal bMatched As Boolean, ByVal iRow As Integer, ByVal tLetter As String, ByVal sBalanceSent As String, ByVal bMatchedCounter As Boolean)
        Try
            '******************************************************************
            '* 2021-12-13 RFK: 
            If bMatched Then
                tRALOCX = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RALOCX", iRow).Trim
                tRamLOCX = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAMLOCX", iRow).Trim
                letter_sent_tLetterCounter = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RALETC", iRow).Trim
                tRABALD = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RABALD", iRow).Trim
                tRAMBAL = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAMBAL", iRow).Trim
                tState = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAGSTATE", iRow).Trim
                tFacility = rkutils.ReadField(DataGridView_Multi, "RAFACL", iRow)
                tAccountNumber = rkutils.ReadField(DataGridView_Multi, "RAACCT", iRow)
            Else
                tRALOCX = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RALOCX", iRow).Trim
                tRamLOCX = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAMLOCX", iRow).Trim
                letter_sent_tLetterCounter = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RALETC", iRow).Trim
                tRABALD = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RABALD", iRow).Trim
                tRAMBAL = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAMBAL", iRow).Trim
                tState = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGSTATE", iRow).Trim
                tFacility = rkutils.ReadField(DataGridView_Select, "RAFACL", iRow)
                tAccountNumber = rkutils.ReadField(DataGridView_Select, "RAACCT", iRow)
            End If
            If CheckBox_DEBUG.Checked Then MsgStatus("letter_sent/" + sModule + "]Letter:" + tLetter + "]RALOCX:" + tRALOCX + "]RamLOCX:" + tRamLOCX + "]" + tAccountNumber + "]Matched:" + bMatched.ToString, CheckBox_DEBUG.Checked)
            '******************************************************************
            '* 2016-11-09 RFK: corrected for Facility
            '* 2017-07-17 RFK: corrected for bMatched Facility
            iRulesRow = DataGridView_Contains2Cols(DataGridView_RULES, "RRFACL", tFacility, "RRACTV", tLetter)
            '******************************************************************
            If iRulesRow >= 0 Then
                If CheckBox_DEBUG.Checked Then MsgStatus("letter_sent row=" + iRow.ToString + " Letter=" + tLetter + " Facility=" + tFacility + " RulesRow=" + iRulesRow.ToString, CheckBox_DEBUG.Checked)
                tLetterNext = rkutils.DataGridView_ValueByColumnName(DataGridView_RULES, "RRNACT", iRulesRow).Trim
                tLetterNextDays = rkutils.DataGridView_ValueByColumnName(DataGridView_RULES, "RRDAYS", iRulesRow).Trim
                tLetterNextDate = rkutils.STR_DATE_PLUS("TODAY", "+", tLetterNextDays)
            Else
                If CheckBox_DEBUG.Checked Then MsgStatus("ERROR/letter_sent/" + sModule + " row=" + iRow.ToString + " Letter=" + tLetter + " Facility=" + tFacility + " RulesRow=" + iRulesRow.ToString, CheckBox_DEBUG.Checked)
                Exit Sub
            End If
            '******************************************************************
            If CheckBox_DEBUG.Checked Then MsgStatus("letter_sent LetterNext=" + tLetterNext + " NextDays=" + tLetterNextDays + " NextDate=" + tLetterNextDate, CheckBox_DEBUG.Checked)
            '******************************************************************
            sSQL = "INSERT INTO " + sDBO + ".dbo.letter_sent"
            sSQL += " (client,tob,facility,locx,ramlocx"
            sSQL += ",letter_matched"
            sSQL += ",letter_date"
            sSQL += ",letter_type"
            sSQL += ",letter_num"
            sSQL += ",letter_next"
            sSQL += ",letter_nextdate"
            sSQL += ",state"
            sSQL += ",rabald"
            sSQL += ",rambal"
            sSQL += ",modified_date"
            sSQL += ",modified_by"
            sSQL += ") values("
            '******************************************************************
            If bMatched Then
                sSQL += "'" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RACL#", iRow).Trim + "'"
                sSQL += ",'" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RATOB", iRow).Trim + "'"
                sSQL += ",'" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAFACL", iRow).Trim + "'"
                sSQL += ",'" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RALOCX", iRow).Trim + "'"
                sSQL += ",'" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAMLOCX", iRow).Trim + "'"
                sSQL += ",'S'"
            Else
                sSQL += "'" + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RACL#", iRow).Trim + "'"
                sSQL += ",'" + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RATOB", iRow).Trim + "'"
                sSQL += ",'" + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAFACL", iRow).Trim + "'"
                sSQL += ",'" + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RALOCX", iRow).Trim + "'"
                sSQL += ",'" + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAMLOCX", iRow).Trim + "'"
                sSQL += ",'-'"
            End If
            sSQL += ",'" + STR_format("TODAY", "mm/dd/ccyy HH:MM:SS") + "'"
            sSQL += ",'" + Letter_Types_Value(tLetter, "LTYPE") + "'"
            sSQL += ",'" + tLetter + "'"
            sSQL += ",'" + tLetterNext + "'"
            sSQL += ",'" + STR_format(tLetterNextDate, "mm/dd/ccyy") + "'"
            sSQL += ",'" + tState + "'"
            sSQL += "," + tRABALD + ""
            sSQL += "," + tRAMBAL + ""
            sSQL += ",'" + STR_format("TODAY", "mm/dd/ccyy HH:MM:SS") + "'"
            sSQL += ",'" + WhoAmI() + "'"
            sSQL += ")"
            '******************************************************************
            File.AppendAllText(FileNameRPT(), sSQL + vbCrLf)
            If CheckBox_Update.Checked = True Then DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, sSQL)
            '******************************************************************
            '* 2012-08-21 RFK: so does not print again
            iLocxMultiRow = rkutils.DataGridView_Contains(DataGridView_Select, "RALOCX", tRALOCX)
            If rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "ERRORCODE", iLocxMultiRow).Trim = "READY" Then
                If bMatched Then
                    Printed(iLocxMultiRow, tLetterPrinted, "MULTI")
                Else
                    Printed(iLocxMultiRow, tLetterPrinted, "PRINTED")
                End If
            End If
            '******************************************************************
            '* 2012-08-21 RFK: UPDATE
            '******************************************************************
            Dim bScreen As Boolean = False
            If tLetterNextDate <> rkutils.STR_DATE_PLUS("TODAY", "+", tLetterNextDays) And rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAFACL", iRow).Trim > 21 And Val(rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAFACL", iRow).Trim) < 16 And Val(rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAFACL", iRow).Trim) <> 6 Then
                bScreen = True
                If bMatched Then
                    MsgStatus("***Matched:  " + tLOCX + " " + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAFACL", iRow).Trim + " " + tLetter + " " + tLetterNext + " " + tLetterNextDays + " " + tLetterNextDate + " " + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RALETC", iRow).Trim, bScreen)
                Else
                    MsgStatus("***Select : " + tRALOCX + " " + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAFACL", iRow).Trim + " " + tLetter + " " + tLetterNext + " " + tLetterNextDays + " " + tLetterNextDate + " " + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RALETC", iRow).Trim, bScreen)
                End If
            End If
            sSQL = "UPDATE ROIDATA.RACCTP"
            sSQL += " SET RALCAC=" + tLetter
            sSQL += ",RAVERI=" + STR_format("TODAY", "mmddccyy")
            sSQL += ",RASDATE='" + STR_format("TODAY", "ccyymmdd") + "'"              '2014-06-05 RFK:
            sSQL += ",RASBAL='" + sBalanceSent + "'"                                  '2014-06-05 RFK:
            sSQL += ",RALETC=" + Trim(Str(Val(letter_sent_tLetterCounter) + 1))
            sSQL += ",RALNAC=" + Trim(Str(Val(tLetterNext)))                          '2014-06-05 RFK: 
            sSQL += ",RANLMO=" + STR_format(tLetterNextDate, "mm")
            sSQL += ",RANLDY=" + STR_format(tLetterNextDate, "dd")
            sSQL += ",RANLYR=" + STR_format(tLetterNextDate, "ccyy")
            '**************************************************************************************
            '* 2018-03-05 RFK:
            '* 2021-12-29 RFK:
            'If rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "SentBadDebtNotification", iRow).Trim = "Y" Then
            If Letter_Types_Value(tLetter, "SentBadDebtNotification") = "Y" Then
                sSQL += ",RALCAD=" + STR_format("TODAY", "ccyymmdd")
            End If
            '**************************************************************************************
            '* 2021-12-29 RFK: First Letter
            If Letter_Types_Value(tLetter, "FirstLetter") = "Y" Then
                sSQL += ",RATZ12='Y'"
            End If
            '**************************************************************************************
            '* 2021-12-29 RFK: Final Letter
            If Letter_Types_Value(tLetter, "FinalLetter") = "Y" Then
                sSQL += ",RATZ13='Y'"
            End If
            '**************************************************************************************
            '* 2021-12-29 RFK: PPL Letter
            If Letter_Types_Value(tLetter, "PaymentPlanLetter") = "Y" Then
                sSQL += ",RATZ14='Y'"
            End If
            '**************************************************************************************
            Select Case tLetter
                Case "960"
                    sSQL += ",RATZ19='S'"
                    sSQL += ",RATZ20='L'"
                Case Else
                    '
            End Select
            '**************************************************************************************
            sSQL += " WHERE RALOCX=" + tLOCX
            '**************************************************************************************
            File.AppendAllText(FileNameRPT(), sSQL + vbCrLf)
            If CheckBox_Update.Checked = True Then
                DB_COMMAND("DB2", DB2SQLConnectionString, DB2SQLuser, sSQL)
                If CheckBox_SentNote.Checked Then
                    If bMatched Then
                        NOTES_ADD("DB2", DB2SQLConnectionString, DB2SQLuser, "LETTER", Me.DataGridView3, tLOCX, "1", "LS", rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RARSTA", iRow).Trim, "", "SENT LETTER _" + tLetterPrinted)
                    Else
                        NOTES_ADD("DB2", DB2SQLConnectionString, DB2SQLuser, "LETTER", Me.DataGridView3, tLOCX, "1", "LS", rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RARSTA", iRow).Trim, "", "SENT LETTER -" + tLetterPrinted)
                    End If
                End If
                '**************************************************************
                '* 2014-10-07 RFK:
                TRACKS_update("MSSQL", msSQLConnectionString, msSQLuser, "", tLOCX, "", "L", "SENT LETTER:" + tLetterPrinted + "] BAL:" + Me.tRABALD + "] RAMBAL:" + Me.tRAMBAL + "] RamLOCX:" + Me.tRamLOCX)
                '**************************************************************
                '* 2018-01-25 RFK:
                If TextBox_StatusCodeAfterSent.Text.Trim.Length > 0 And Label_StatusCodeSentDescription.Text.Trim.Length > 0 Then rkutils.COMMAND_STATUS(msSQLConnectionString, msSQLuser, tLOCX, TextBox_StatusCodeAfterSent.Text, "", "", "")
            End If
            '******************************************************************
            '* 2018-03-07 RFK:
            If Listbox_Contains(ListBox_Noted, tLOCX, True) >= 0 Then
                MsgStatus("ALREADY NOTED" + tLOCX, CheckBox_DEBUG.Checked)
            Else
                MsgStatus("NOTED" + tRALOCX, CheckBox_DEBUG.Checked)
            End If
            '******************************************************************
            '* 2017-07-19 RFK:
            sSQL = "SELECT LOCX FROM ROIDATA.LETTERSLAST"
            sSQL += rkutils.WhereAnd(sSQL, "LOCX='" + tLOCX + "'")
            sTemp = rkutils.SQL_READ_FIELD(DataGridView3, "DB2", "LOCX", DB2SQLConnectionString, DB2SQLuser, sSQL).Trim
            If sTemp.Length > 0 Then
                sSQL = "UPDATE ROIDATA.LETTERSLAST"
                Select Case Val(tLetterPrinted)
                    Case 310, 910
                        sSQL += " SET LETTER10SENTDATE='" + rkutils.STR_format("TODAY", "ccyymmdd") + "'"
                        sSQL += ",LETTER10BALANCE='" + rkutils.STR_format(sSentBalance, "0.00") + "'"
                        sSQL += ",LETTER10RAMLOCX='" + tRamLOCX + "'"
                        sSQL += ",LETTER10VENDOR='" + rkutils.STR_TRIM(tCurrentLetterVendor, 1) + "'"
                    Case 311, 911
                        sSQL += " SET LETTER11SENTDATE='" + rkutils.STR_format("TODAY", "ccyymmdd") + "'"
                        sSQL += ",LETTER11BALANCE='" + rkutils.STR_format(sSentBalance, "0.00") + "'"
                        sSQL += ",LETTER11RAMLOCX" + tRamLOCX + "'"
                        sSQL += ",LETTER11VENDOR" + rkutils.STR_TRIM(tCurrentLetterVendor, 1) + "'"
                    Case 315, 915
                        sSQL += " SET LETTER15SENTDATE='" + rkutils.STR_format("TODAY", "ccyymmdd") + "'"
                        sSQL += ",LETTER15BALANCE='" + rkutils.STR_format(sSentBalance, "0.00") + "'"
                        sSQL += ",LETTER15RAMLOCX='" + tRamLOCX + "'"
                        sSQL += ",LETTER15VENDOR='" + rkutils.STR_TRIM(tCurrentLetterVendor, 1) + "'"
                    Case 320, 920
                        sSQL += " SET LETTER20SENTDATE='" + rkutils.STR_format("TODAY", "ccyymmdd") + "'"
                        sSQL += ",LETTER20BALANCE='" + rkutils.STR_format(sSentBalance, "0.00") + "'"
                        sSQL += ",LETTER20RAMLOCX='" + tRamLOCX + "'"
                        sSQL += ",LETTER20VENDOR='" + rkutils.STR_TRIM(tCurrentLetterVendor, 1) + "'"
                    Case 330, 930
                        sSQL += " SET LETTER30SENTDATE='" + rkutils.STR_format("TODAY", "ccyymmdd") + "'"
                        sSQL += ",LETTER30BALANCE='" + rkutils.STR_format(sSentBalance, "0.00") + "'"
                        sSQL += ",LETTER30RAMLOCX='" + tRamLOCX + "'"
                        sSQL += ",LETTER30VENDOR='" + rkutils.STR_TRIM(tCurrentLetterVendor, 1) + "'"
                    Case 340, 940
                        sSQL += " SET LETTER40SENTDATE='" + rkutils.STR_format("TODAY", "ccyymmdd") + "'"
                        sSQL += ",LETTER40BALANCE='" + rkutils.STR_format(sSentBalance, "0.00") + "'"
                        sSQL += ",LETTER40RAMLOCX='" + tRamLOCX + "'"
                        sSQL += ",LETTER40VENDOR='" + rkutils.STR_TRIM(tCurrentLetterVendor, 1) + "'"
                    Case 350, 950
                        sSQL += " SET LETTER50SENTDATE='" + rkutils.STR_format("TODAY", "ccyymmdd") + "'"
                        sSQL += ",LETTER50BALANCE='" + rkutils.STR_format(sSentBalance, "0.00") + "'"
                        sSQL += ",LETTER50RAMLOCX='" + tRamLOCX + "'"
                        sSQL += ",LETTER50VENDOR='" + rkutils.STR_TRIM(tCurrentLetterVendor, 1) + "'"
                    Case 351, 951
                        sSQL += " SET LETTER51SENTDATE='" + rkutils.STR_format("TODAY", "ccyymmdd") + "'"
                        sSQL += ",LETTER51BALANCE='" + rkutils.STR_format(sSentBalance, "0.00") + "'"
                        sSQL += ",LETTER51RAMLOCX='" + tRamLOCX + "'"
                        sSQL += ",LETTER51VENDOR='" + rkutils.STR_TRIM(tCurrentLetterVendor, 1) + "'"
                    Case 960
                        sSQL += " SET LETTER60SENTDATE='" + rkutils.STR_format("TODAY", "ccyymmdd") + "'"
                        sSQL += ",LETTER60BALANCE='" + rkutils.STR_format(sSentBalance, "0.00") + "'"
                        sSQL += ",LETTER60RAMLOCX='" + tRamLOCX + "'"
                        sSQL += ",LETTER60VENDOR='" + rkutils.STR_TRIM(tCurrentLetterVendor, 1) + "'"
                    Case Else
                        MsgStatus("You MUST define this letter type in ROIDATA.LETTERSLAST", True)
                        Exit Sub
                End Select
                sSQL += rkutils.WhereAnd(sSQL, "LOCX='" + tLOCX + "'")
            Else
                sSQL = "INSERT INTO ROIDATA.LETTERSLAST"
                sSQL += " (LOCX"
                Select Case Val(tLetterPrinted)
                    Case 310, 910
                        sSQL += ",LETTER10SENTDATE"
                        sSQL += ",LETTER10BALANCE"
                        sSQL += ",LETTER10RAMLOCX"
                        sSQL += ",LETTER10VENDOR"
                    Case 311, 911
                        sSQL += ",LETTER11SENTDATE"
                        sSQL += ",LETTER11BALANCE"
                        sSQL += ",LETTER11RAMLOCX"
                        sSQL += ",LETTER11VENDOR"
                    Case 315, 915
                        sSQL += ",LETTER15SENTDATE"
                        sSQL += ",LETTER15BALANCE"
                        sSQL += ",LETTER15RAMLOCX"
                        sSQL += ",LETTER15VENDOR"
                    Case 320, 920
                        sSQL += ",LETTER20SENTDATE"
                        sSQL += ",LETTER20BALANCE"
                        sSQL += ",LETTER20RAMLOCX"
                        sSQL += ",LETTER20VENDOR"
                    Case 330, 930
                        sSQL += ",LETTER30SENTDATE"
                        sSQL += ",LETTER30BALANCE"
                        sSQL += ",LETTER30RAMLOCX"
                        sSQL += ",LETTER30VENDOR"
                    Case 340, 940
                        sSQL += ",LETTER40SENTDATE"
                        sSQL += ",LETTER40BALANCE"
                        sSQL += ",LETTER40RAMLOCX"
                        sSQL += ",LETTER40VENDOR"
                    Case 350, 950
                        sSQL += ",LETTER50SENTDATE"
                        sSQL += ",LETTER50BALANCE"
                        sSQL += ",LETTER50RAMLOCX"
                        sSQL += ",LETTER50VENDOR"
                    Case 351, 951
                        sSQL += ",LETTER51SENTDATE"
                        sSQL += ",LETTER51BALANCE"
                        sSQL += ",LETTER51RAMLOCX"
                        sSQL += ",LETTER51VENDOR"
                    Case 960
                        sSQL += ",LETTER60SENTDATE"
                        sSQL += ",LETTER60BALANCE"
                        sSQL += ",LETTER60RAMLOCX"
                        sSQL += ",LETTER60VENDOR"
                End Select
                sSQL += ") VALUES('" + tLOCX + "'"
                Select Case Val(tLetterPrinted)
                    Case 310, 311, 315, 320, 330, 340, 350, 351, 910, 911, 915, 920, 930, 940, 950, 951, 960
                        sSQL += ",'" + rkutils.STR_format("TODAY", "ccyymmdd") + "'"
                        sSQL += ",'" + rkutils.STR_format(sSentBalance, "0.00") + "'"
                        sSQL += ",'" + tRamLOCX + "'"
                        sSQL += ",'" + rkutils.STR_TRIM(tCurrentLetterVendor, 1) + "'"
                End Select
                sSQL += ")"
            End If
            If CheckBox_Update.Checked Then DB_COMMAND("DB2", DB2SQLConnectionString, DB2SQLuser, sSQL)
            '******************************************************************
            '* 2012-11-09 RFK: Screen Counters
            If bMatched Then
                '**************************************************************
            Else
                '**************************************************************
                '* 2012-11-13 RFK:
                Label_Printed.Text = Str(Val(Label_Printed.Text) + 1).Trim
                LettersPrinted(tLetter, tLOCX, False)
            End If
            '******************************************************************
            '* 2012-11-09 RFK: Screen Counters
            If bMatchedCounter Then
                Label_PrintedM.Text = Str(Val(Label_PrintedM.Text) + 1).Trim
            End If
            '******************************************************************
        Catch ex As Exception
            MsgError("letter_sent/ERROR:", ex.ToString)
        End Try
    End Sub

    Private Function MultiAccounts(ByVal iRow As Integer, ByVal sLetterType As String) As String
        Try
            '******************************************************************
            '* 2013-02-04 RFK:
            tMultiMessage = ""
            '******************************************************************
            '* 2021-12-15 RFK:
            Select Case sLetterType
                Case "10"
                    '**********************************************************
                    Dim sMA_RAMLOCX As String = "", sMAS_RAMLOCX As String = ""
                    Dim sMA_PlacementDate1 As String = "", sMAS_PlacementDate1 As String = ""
                    Dim iCount As Integer = 0
                    MultiAccounts_imLocxRow = iRow
                    If swDTable Then
                        sMA_RAMLOCX = rkutils.DataTable_ValueByColumnName(dTable_Select, "RAMLOCX", MultiAccounts_imLocxRow)
                        sMA_PlacementDate1 = rkutils.STR_format(rkutils.DataTable_ValueByColumnName(dTable_Select, "PlacementDate1", MultiAccounts_imLocxRow), "ccyymmdd")
                        iCount = dTable_Select.Rows.Count - 1
                    Else
                        sMA_RAMLOCX = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAMLOCX", MultiAccounts_imLocxRow)
                        sMA_PlacementDate1 = rkutils.STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "PlacementDate1", MultiAccounts_imLocxRow), "ccyymmdd")
                        iCount = DataGridView_Select.Rows.Count - 1
                    End If
                    sMAS_RAMLOCX = sMA_RAMLOCX
                    sMAS_PlacementDate1 = sMA_PlacementDate1
                    '**********************************************************
                    Do While sMA_RAMLOCX = sMAS_RAMLOCX And sMA_PlacementDate1 = sMAS_PlacementDate1
                        tMultiMessage += "["
                        If swDTable Then
                            tMultiMessage += rkutils.DataTable_ValueByColumnName(dTable_Select, "RAACCT", MultiAccounts_imLocxRow).Trim
                            tMultiMessage += " " + STR_format(rkutils.DataTable_ValueByColumnName(dTable_Select, "RABALD", MultiAccounts_imLocxRow).Trim, "$")
                        Else
                            tMultiMessage += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAACCT", MultiAccounts_imLocxRow).Trim
                            tMultiMessage += " " + STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RABALD", MultiAccounts_imLocxRow).Trim, "$")
                        End If
                        tMultiMessage += "]"
                        '******************************************************
                        If MultiAccounts_imLocxRow + 1 <= iCount Then
                            MultiAccounts_imLocxRow += 1
                            If swDTable Then
                                sMA_RAMLOCX = rkutils.DataTable_ValueByColumnName(dTable_Select, "RAMLOCX", MultiAccounts_imLocxRow)
                                sMA_PlacementDate1 = rkutils.STR_format(rkutils.DataTable_ValueByColumnName(dTable_Select, "PlacementDate1", MultiAccounts_imLocxRow), "ccyymmdd")
                            Else
                                sMA_RAMLOCX = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAMLOCX", MultiAccounts_imLocxRow)
                                sMA_PlacementDate1 = rkutils.STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "PlacementDate1", MultiAccounts_imLocxRow), "ccyymmdd")
                            End If
                        Else
                            sMA_RAMLOCX = ""
                            sMA_PlacementDate1 = ""
                        End If
                    Loop
                    Return tMultiMessage
                Case Else
                    'Keep Going
            End Select
            '******************************************************************
            '* 2019-03-14 RFK:
            If swDTable Then
                MultiAccounts_sSysAccountMatched = rkutils.DataTable_ValueByColumnName(dTable_Select, "RAMLOCX", iRow).Trim
            Else
                MultiAccounts_sSysAccountMatched = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAMLOCX", iRow).Trim
            End If
            MultiAccounts_imLocxRow = rkutils.DataGridView_Contains(DataGridView_Multi, "RAMLOCX", sSysAccountMatched)
            MultiAccounts_iMatchedRow = MultiAccounts_imLocxRow
            Do While MultiAccounts_imLocxRow <= DataGridView_Multi.RowCount - 1 And sSysAccountMatched = rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAMLOCX", MultiAccounts_imLocxRow).Trim
                Select Case sLetterType
                    Case "25"
                        '******************************************************
                        '* 2021-07-22 RFK: 
                        tMultiMessage += "["
                        '* Account Number
                        tMultiMessage += rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAACCT", MultiAccounts_imLocxRow).Trim
                        '* Patient Name
                        tMultiMessage += "|" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAPNAM", MultiAccounts_imLocxRow).Trim
                        '* DOS
                        tMultiMessage += "|" + STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISM", MultiAccounts_imLocxRow).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISD", MultiAccounts_imLocxRow).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RADISY", MultiAccounts_imLocxRow), "mm/dd/ccyy")
                        '* Balance
                        tMultiMessage += "|" + STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RABALD", MultiAccounts_imLocxRow).Trim, "$")
                        tMultiMessage += "]"
                        '******************************************************
                    Case Else
                        '******************************************************
                        tMultiMessage += "[" + rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RAACCT", MultiAccounts_imLocxRow).Trim
                        tMultiMessage += " " + STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Multi, "RABALD", MultiAccounts_imLocxRow).Trim, "$")
                        tMultiMessage += "]"
                        '******************************************************
                End Select
                MultiAccounts_imLocxRow += 1
            Loop
            Return tMultiMessage
        Catch ex As Exception
            MsgError("MultiAccounts", ex.ToString)
            Return ""
        End Try
    End Function

    Private Function CalculateSentBalance(ByVal sMatchedAccount As String, ByVal sLetterVendor As String, ByVal sLetterType As String, ByVal sLetterNumber As String, ByVal CSB_irow As Integer) As String
        Try
            '******************************************************************
            '* 2015-08-20 RFK: added bucket dollars
            Dim CalculateSentBalance_iLocxRow As Integer = 0

            SentBalanceDollars30 = 0
            SentBalanceDollars60 = 0
            SentBalanceDollars90 = 0
            SentBalanceDollars120 = 0
            SentBalanceDollars121 = 0
            dBalance = 0
            SentBalance = 0
            sAccountSave = ""
            Select Case sSITE
                Case "AnnuityOne"
                    '**********************************************************
                    '* 2021-12-15 RFK:
                    Select Case sLetterType
                        Case "10"
                            '**************************************************
                            Dim sMA_RAMLOCX As String = "", sMAS_RAMLOCX As String = ""
                            Dim sMA_PlacementDate1 As String = "", sMAS_PlacementDate1 As String = ""
                            Dim iCount As Integer = 0
                            MultiAccounts_imLocxRow = CSB_irow
                            If swDTable Then
                                sMA_RAMLOCX = rkutils.DataTable_ValueByColumnName(dTable_Select, "RAMLOCX", MultiAccounts_imLocxRow)
                                sMA_PlacementDate1 = rkutils.STR_format(rkutils.DataTable_ValueByColumnName(dTable_Select, "PlacementDate1", MultiAccounts_imLocxRow), "ccyymmdd")
                                iCount = dTable_Select.Rows.Count - 1
                            Else
                                sMA_RAMLOCX = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAMLOCX", MultiAccounts_imLocxRow)
                                sMA_PlacementDate1 = rkutils.STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "PlacementDate1", MultiAccounts_imLocxRow), "ccyymmdd")
                                iCount = DataGridView_Select.Rows.Count - 1
                            End If
                            sMAS_RAMLOCX = sMA_RAMLOCX
                            sMAS_PlacementDate1 = sMA_PlacementDate1
                            '**********************************************************************
                            Do While sMA_RAMLOCX = sMAS_RAMLOCX And sMA_PlacementDate1 = sMAS_PlacementDate1
                                '******************************************************************
                                If swDTable Then
                                    SentBalance = Val(rkutils.DataTable_ValueByColumnName(dTable_Select, "RABALD", MultiAccounts_imLocxRow))
                                    sDate = rkutils.DataTable_ValueByColumnName(dTable_Select, "RADISM", MultiAccounts_imLocxRow).Trim + "/" + rkutils.DataTable_ValueByColumnName(dTable_Select, "RADISD", MultiAccounts_imLocxRow).Trim + "/" + rkutils.DataTable_ValueByColumnName(dTable_Select, "RADISY", MultiAccounts_imLocxRow).Trim
                                Else
                                    SentBalance = Val(ReadField(DataGridView_Select, "RABALD", MultiAccounts_imLocxRow))
                                    sDate = rkutils.ReadField(DataGridView_Select, "RADISM", MultiAccounts_imLocxRow).Trim + "/" + rkutils.ReadField(DataGridView_Select, "RADISD", MultiAccounts_imLocxRow).Trim + "/" + rkutils.ReadField(DataGridView_Select, "RADISY", MultiAccounts_imLocxRow).Trim
                                End If
                                dBalance += SentBalance
                                '******************************************************************
                                '* 2015-08-20 RFK: added bucket dollars
                                If IsDate(sDate) Then
                                    Select Case Val(rkutils.STR_format(sDate, "AGE_D"))
                                        Case 0 To 30
                                            SentBalanceDollars30 += SentBalance
                                        Case 31 To 60
                                            SentBalanceDollars60 += SentBalance
                                        Case 61 To 90
                                            SentBalanceDollars90 += SentBalance
                                        Case 91 To 120
                                            SentBalanceDollars120 += SentBalance
                                        Case Else
                                            SentBalanceDollars121 += SentBalance
                                    End Select
                                Else
                                    MsgStatus("Calculate Sent SDate Failure:" + sAccountSave + "[" + sDate + "]", True)
                                End If
                                '**********************************************
                                '**********************************************
                                If MultiAccounts_imLocxRow + 1 <= iCount Then
                                    MultiAccounts_imLocxRow += 1
                                    If swDTable Then
                                        sMA_RAMLOCX = rkutils.DataTable_ValueByColumnName(dTable_Select, "RAMLOCX", MultiAccounts_imLocxRow)
                                        sMA_PlacementDate1 = rkutils.STR_format(rkutils.DataTable_ValueByColumnName(dTable_Select, "PlacementDate1", MultiAccounts_imLocxRow), "ccyymmdd")
                                    Else
                                        sMA_RAMLOCX = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAMLOCX", MultiAccounts_imLocxRow)
                                        sMA_PlacementDate1 = rkutils.STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "PlacementDate1", MultiAccounts_imLocxRow), "ccyymmdd")
                                    End If
                                Else
                                    sMA_RAMLOCX = ""
                                    sMA_PlacementDate1 = ""
                                End If
                            Loop
                            Return rkutils.STR_format(dBalance.ToString, "0.00")
                            '**************************************************
                        Case Else
                            '**********************************************************
                            '* 2015-01-15 RFK:
                            CalculateSentBalance_iLocxRow = DataGridView_Contains(DataGridView_Multi, "RAMLOCX", sMatchedAccount)
                            If CalculateSentBalance_iLocxRow >= 0 Then
                                sAccountSave = ReadField(DataGridView_Multi, "RAMLOCX", CalculateSentBalance_iLocxRow)
                                Do While CalculateSentBalance_iLocxRow < DataGridView_Multi.Rows.Count - 1 And Label_RUNNING.Text = "Running" And sAccountSave = sMatchedAccount
                                    SentBalance = Val(ReadField(DataGridView_Multi, "RABALD", CalculateSentBalance_iLocxRow))
                                    dBalance += SentBalance
                                    '**************************************************
                                    '* 2016-01-27 RFK: 20 RAFR30 contains SelfPayDate
                                    '* 2016-02-09 RFK: 19 using placement date
                                    '* 2021-08-31 RFK: 26 similar to 19
                                    Select Case sLetterType
                                        Case "19", "26"   'Placement Date
                                            sDate = rkutils.ReadField(DataGridView_Multi, "RAAMON", CalculateSentBalance_iLocxRow).Trim + "/" + rkutils.ReadField(DataGridView_Multi, "RAADAY", CalculateSentBalance_iLocxRow).Trim + "/" + rkutils.ReadField(DataGridView_Multi, "RAAYR", CalculateSentBalance_iLocxRow).Trim
                                            If IsDate(sDate) = False Then
                                                MsgStatus("Calculate SDate Failure Type:" + sLetterType + "]Account:" + sAccountSave + "]sDate:" + sDate + "]", True)
                                                sDate = rkutils.ReadField(DataGridView_Multi, "RADISM", CalculateSentBalance_iLocxRow).Trim + "/" + rkutils.ReadField(DataGridView_Multi, "RADISD", CalculateSentBalance_iLocxRow).Trim + "/" + rkutils.ReadField(DataGridView_Multi, "RADISY", CalculateSentBalance_iLocxRow).Trim
                                                MsgStatus("SDate Now DOS:" + sDate + "]", True)
                                            End If
                                        Case "20"
                                            sDate = rkutils.ReadField(DataGridView_Multi, "RAFR30", CalculateSentBalance_iLocxRow).Trim
                                            If IsDate(sDate) = False Then
                                                MsgStatus("Calculate SDate Failure Type:" + sLetterType + "]Account:" + sAccountSave + "]sDate:" + sDate + "]", True)
                                                sDate = rkutils.ReadField(DataGridView_Multi, "RADISM", CalculateSentBalance_iLocxRow).Trim + "/" + rkutils.ReadField(DataGridView_Multi, "RADISD", CalculateSentBalance_iLocxRow).Trim + "/" + rkutils.ReadField(DataGridView_Multi, "RADISY", CalculateSentBalance_iLocxRow).Trim
                                                MsgStatus("SDate Now DOS:" + sDate + "]", True)
                                            End If
                                        Case Else
                                            sDate = rkutils.ReadField(DataGridView_Multi, "RADISM", CalculateSentBalance_iLocxRow).Trim + "/" + rkutils.ReadField(DataGridView_Multi, "RADISD", CalculateSentBalance_iLocxRow).Trim + "/" + rkutils.ReadField(DataGridView_Multi, "RADISY", CalculateSentBalance_iLocxRow).Trim
                                    End Select
                                    '**************************************************
                                    '* 2016-02-01 RFK: 
                                    If IsDate(sDate) = False Then
                                        MsgStatus("Calculate SDate Failure Type:" + sLetterType + "]Account:" + sAccountSave + "]sDate:" + sDate + "]", True)
                                        sDate = rkutils.ReadField(DataGridView_Multi, "RAAMON", CalculateSentBalance_iLocxRow).Trim + "/" + rkutils.ReadField(DataGridView_Multi, "RAADAY", CalculateSentBalance_iLocxRow).Trim + "/" + rkutils.ReadField(DataGridView_Multi, "RAAYR", CalculateSentBalance_iLocxRow).Trim
                                        MsgStatus("SDate Now PLACEMENT:" + sDate + "]", True)
                                    End If
                                    '**************************************************
                                    '* 2015-08-20 RFK: added bucket dollars
                                    If IsDate(sDate) Then
                                        Select Case Val(rkutils.STR_format(sDate, "AGE_D"))
                                            Case 0 To 30
                                                SentBalanceDollars30 += SentBalance
                                            Case 31 To 60
                                                SentBalanceDollars60 += SentBalance
                                            Case 61 To 90
                                                SentBalanceDollars90 += SentBalance
                                            Case 91 To 120
                                                SentBalanceDollars120 += SentBalance
                                            Case Else
                                                SentBalanceDollars121 += SentBalance
                                        End Select
                                    Else
                                        MsgStatus("Calculate Sent SDate Failure:" + sAccountSave + "[" + sDate + "]", True)
                                    End If
                                    '**************************************************
                                    CalculateSentBalance_iLocxRow += 1
                                    sAccountSave = ReadField(DataGridView_Multi, "RAMLOCX", CalculateSentBalance_iLocxRow)
                                Loop
                            End If
                    End Select
                    '**********************************************************
                Case "iTeleCollect"
                    'REMOVED 
            End Select
            Return dBalance.ToString
        Catch ex As Exception
            MsgError("CalculateSentBalance", ex.ToString)
            Return ""
        End Try
    End Function

    Private Sub Run()
        Try
            '**************************************************************************************
            '* 
            Dim tLastPayDate As String = "", sFacility As String = "", sSQL As String = ""
            Dim tStatus As String = "", sPPLstatus As String = ""
            Dim swROW As Boolean = False
            Dim iLocxRow As Integer = 0, imLocxRow As Integer = 0
            Dim tSysAccountMatched As String = ""
            Dim tSQLstring As String = ""
            Dim iRulesRow As Integer = 0, iRulesRowDB2 As Integer = 0
            Dim iCount As Integer = 0
            Dim tState As String = ""
            Dim sLastLetterDate As String = "", sOKdate As String = ""
            '**************************************************************************************
            '* 
            MsgStatus("Run", True)
            CountersClear()
            '**************************************************************************************
            If swDTable Then
                If dTable_Select.Rows.Count - 1 <= 0 Then
                    MsgStatus("Can NOT Run(" + (dTable_Select.Rows.Count - 1).ToString + ")", True)
                    Exit Sub
                End If
            Else
                If DataGridView_Select.Rows.Count - 1 <= 0 Then
                    MsgStatus("Can NOT Run(" + (DataGridView_Select.Rows.Count - 1).ToString + ")", True)
                    Exit Sub
                End If
            End If
            '**************************************************************************************
            RunStart()
            iGhosts_CTR = 0
            '**************************************************************************************
            '* 2014-08-04 RFK:
            sSQL = "UPDATE " + sDBO + ".dbo.clients"
            sSQL += " SET LettersStartDate='" + rkutils.STR_format("TODAY", "mm/dd/ccyy HH:MM:SS") + "'"
            sSQL += ",LettersRunDate=NULL"
            sSQL += " WHERE ClientName='" + rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)) + "'"
            rkutils.DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, sSQL)
            '**************************************************************************************
            Me.Button_Run.Text = "STOP"
            '**************************************************************************************
            Label_AccountsRemaining.Text = Label_NumberAccounts.Text
            '**************************************************************************************
            Select Case sSITE
                Case "AnnuityOne"
                    tCurrentLetterVendor = rkutils.DataGridView_ValueByColumnName(DataGridView_Letter_Types, "VENDOR", 0).Trim()
                    Select Case tCurrentLetterVendor
                        Case "ACCUDOC"
                            '**********************************************************************
                            If Me.CheckBox_Update.Checked = True Then
                                Label_LetterFile.Text = dir_LETTERS
                            Else
                                Label_LetterFile.Text = dir_TEST
                            End If
                            '**********************************************************************
                            Label_LetterFile.Text += rkutils.STR_format("TODAY", "ccyymmdd") + "_" + rkutils.STR_format("TODAY", "HH") + "_"
                            '**********************************************************************
                            If Me.CheckBox_Update.Checked = True Then
                                '******************************************************************
                            Else
                                Label_LetterFile.Text += "TEST_"
                            End If
                            '**********************************************************************
                            If Panel_RegF.Visible Then
                                If rkutils.DataGridView_Contains(DataGridView_Select, "RALNAC", "910") >= 0 Then
                                    Label_LetterFile.Text += "910_"
                                End If
                            End If
                            '******************************************************************
                            If Panel_RegF.Visible Then
                                If rkutils.DataGridView_Contains(DataGridView_Select, "RARSTA", "PPL") >= 0 Then
                                    Label_LetterFile.Text += "PP_"
                                Else
                                    If rkutils.DataGridView_Contains(DataGridView_Select, "RARSTA", "PPC") >= 0 Then
                                        Label_LetterFile.Text += "PP_"
                                    End If
                                End If
                            End If
                            '******************************************************************
                            Label_LetterFile.Text += rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text))
                            Label_LetterFile.Text += ".ACC"
                            '******************************************************************
                        Case "APEX"
                            If Me.CheckBox_Update.Checked = True Then
                                Label_LetterFile.Text = dir_LETTERS + rkutils.STR_format("TODAY", "ccyymmdd") + "_" + rkutils.STR_format("TODAY", "HH") + "_" + rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text))
                            Else
                                Label_LetterFile.Text = dir_TEST + rkutils.STR_format("TODAY", "ccyymmdd") + "_TEST_" + rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text))
                            End If
                            Label_LetterFile.Text += ".APX"
                        Case "DANTOM", "REVSPRING"
                            If Me.CheckBox_Update.Checked = True Then
                                Label_LetterFile.Text = dir_LETTERS + rkutils.STR_format("TODAY", "ccyymmdd") + "_" + rkutils.STR_format("TODAY", "HH") + "_" + rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)) + ".DAN"
                            Else
                                Label_LetterFile.Text = dir_TEST + rkutils.STR_format("TODAY", "ccyymmdd") + "_TEST_" + rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)) + ".DAN"
                            End If
                        Case "DIAMOND"
                            If Me.CheckBox_Update.Checked = True Then
                                Label_LetterFile.Text = dir_LETTERS + rkutils.STR_format("TODAY", "mmddccyy") + "_" + rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)) + "_patinfo.DAT"
                            Else
                                Label_LetterFile.Text = dir_TEST + rkutils.STR_format("TODAY", "mmddccyy") + "_TEST_" + rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)) + "_patinfo.DAT"
                            End If
                        Case Else
                            MsgStatus("Invalid Letter Vendor [" + tCurrentLetterVendor + "]", True)
                            Exit Sub
                    End Select
                Case "iTeleCollect"
                    '*
            End Select
            '******************************************************************
            MsgStatus(tCurrentLetterVendor + " " + Label_LetterFile.Text, True)
            '******************************************************************
            If File.Exists(Label_LetterFile.Text) Then
                If CheckBox_Update.Checked = True Then
                    MsgStatus("CAN NOT CREATE [ALREADY EXISTS]" + Label_LetterFile.Text, True)
                    Exit Sub
                End If
                '**************************************************************
                '* 2012-10-01 RFK: In Test Mode, so delete it
                MsgStatus("DELETED [ALREADY EXISTS]" + Label_LetterFile.Text, True)
                File.Delete(Label_LetterFile.Text)
            End If
            '******************************************************************
            MsgStatus("Starting", True)
            Application.DoEvents()
            '******************************************************************
            '* 2012-11-20 RFK: Reset ALL ERRORCODE to READY
            '* 2019-03-14 RFK:
            If swDTable Then
                iCount = dTable_Select.Rows.Count - 1
            Else
                iCount = DataGridView_Select.Rows.Count - 1
            End If
            '******************************************************************
            '* 2019-03-14 RFK:
            Do While iLocxRow < iCount And Label_RUNNING.Text = "Running"
                If swDTable Then
                    rkutils.DataTable_SetValueByColumnName(dTable_Select, "ERRORCODE", iLocxRow, "READY")
                Else
                    rkutils.DataGridView_SetValueByColumnName(DataGridView_Select, "ERRORCODE", iLocxRow, "READY")
                End If
                iLocxRow += 1
            Loop
            '******************************************************************
            MsgStatus("Running", True)
            PrintHeader(tCurrentLetterVendor, Label_LetterFile.Text)
            '******************************************************************
            iLocxRow = 0
            Do While iLocxRow < iCount And Label_RUNNING.Text = "Running"
                '**************************************************************
                If CheckBox_DEBUG.Checked Then MsgStatus("Run:" + iLocxRow.ToString, True)
                Label_AccountsRemaining.Text = (iCount - iLocxRow).ToString
                '**************************************************************
                'If iLocxRow Mod 10 = 0 Then Application.DoEvents()
                Application.DoEvents()
                '**************************************************************
                '* 2017-06-02 RFK: Ghosts
                If rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "eLettersGhosts", Val(Label_ClientRow.Text)) = "Y" And iGhosts_CTR < DataGridView_Ghosts.Rows.Count Then
                    '******************************************************************
                    '* 2019-03-14 RFK:
                    If swDTable Then
                        'Need To Convert 
                    Else
                        If ReadField(DataGridView_Select, "RAGLNM", iLocxRow) > ReadField(DataGridView_Ghosts, "LName", iGhosts_CTR) Then
                            MsgStatus("GHOST:" + ReadField(DataGridView_Ghosts, "LName", iGhosts_CTR), True)
                            ReadLine(iLocxRow)
                            AnnuityOne_AccountsLoad(rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "TOB", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "FACILITY", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTMATCHTYPE", Val(Label_ClientRow.Text)), "RAMLOCX", gtLetterType, tRamLOCX, Letter_Types_Value(rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RALNAC", Val(Label_AccountRow.Text)), "ChargeOffDate"))
                            PrintLine("run", iLocxRow, tCurrentLetterVendor, Letter_Types_Value(tLetterPrinted, "LTYPE"), tLetterPrinted, Label_LetterFile.Text, True)
                            MsgStatus("GHOST/DONE", True)
                            iGhosts_CTR += 1
                        End If
                    End If
                End If
                '*************************************************************
                '* 2015-07-28 RFK:
                Select Case sSITE
                    Case "AnnuityOne"
                        '******************************************************
                        '* 2014-03-07 RFK:
                        '* 2019-03-14 RFK:
                        If swDTable Then
                            tRALOCX = rkutils.DataTable_ValueByColumnName(dTable_Select, "RALOCX", iLocxRow).Trim
                            If ReadFieldDataTable(dTable_Select, "RAGLNM", iLocxRow).Trim.Length = 0 Or ReadFieldDataTable(dTable_Select, "RAGFNM", iLocxRow).Trim.Length = 0 Then
                                Select Case rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "ELETTERS", Val(Label_ClientRow.Text))
                                    Case "t"
                                        '**************************************
                                        '* 2015-07-21 RFK: 
                                    Case Else
                                        '**************************************
                                        '* 2019-03-14 RFK:
                                        If swDTable Then
                                            rkutils.DataTable_SetValueByColumnName(dTable_Select, "ERRORCODE", iLocxRow, "INVALID_GUARANTOR")
                                        Else
                                            rkutils.DataGridView_SetValueByColumnName(DataGridView_Select, "ERRORCODE", iLocxRow, "INVALID_GUARANTOR")
                                        End If
                                        LettersPrinted("GUAR", tRALOCX, True)
                                        If CheckBox_Update.Checked Then rkutils.COMMAND_STATUS(msSQLConnectionString, msSQLuser, tRALOCX, sStatusCodeBadGuarantorName, "", "", "")
                                End Select
                            End If
                        Else
                            tRALOCX = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RALOCX", iLocxRow).Trim
                            If ReadField(DataGridView_Select, "RAGLNM", iLocxRow).Trim.Length = 0 Or ReadField(DataGridView_Select, "RAGFNM", iLocxRow).Trim.Length = 0 Then
                                Select Case rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "ELETTERS", Val(Label_ClientRow.Text))
                                    Case "t"
                                        '**************************************
                                        '* 2015-07-21 RFK: 
                                    Case Else
                                        '**************************************
                                        '* 2019-03-14 RFK:
                                        If swDTable Then
                                            rkutils.DataTable_SetValueByColumnName(dTable_Select, "ERRORCODE", iLocxRow, "INVALID_GUARANTOR")
                                        Else
                                            rkutils.DataGridView_SetValueByColumnName(DataGridView_Select, "ERRORCODE", iLocxRow, "INVALID_GUARANTOR")
                                        End If
                                        LettersPrinted("GUAR", tRALOCX, True)
                                        If CheckBox_Update.Checked Then rkutils.COMMAND_STATUS(msSQLConnectionString, msSQLuser, tRALOCX, sStatusCodeBadGuarantorName, "", "", "")
                                End Select
                            End If
                        End If
                    Case Else
                        '******************************************************
                        '* 2015-07-28 RFK:
                End Select
                '**************************************************************
                '* 2014-08-05 RFK:
                '* 2019-03-14 RFK:
                If swDTable Then
                    sTemp = ReadFieldDataTable(dTable_Select, "ERRORCODE", iLocxRow)
                Else
                    sTemp = ReadField(DataGridView_Select, "ERRORCODE", iLocxRow)
                End If
                If sTemp = "READY" Then
                    '**********************************************************
                    Select Case sSITE
                        Case "AnnuityOne"
                            ReadLine(iLocxRow)
                            '**************************************************
                            If CheckBox_DEBUG.Checked Then MsgStatus("iLocxRow:" + iLocxRow.ToString + " [" + tRALOCX + "][" + tRamLOCX + "]", False)
                            '*********************************************************************************************************
                            '* 2014-08-05 RFK:
                            If swReadAllMatched = False Then
                                Label_AccountRow.Text = iLocxRow.ToString.Trim
                                '**********************************************
                                '* 2021-11-26 RFK:
                                AnnuityOne_AccountsLoad(rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "TOB", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "FACILITY", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTMATCHTYPE", Val(Label_ClientRow.Text)), "RAMLOCX", gtLetterType, tRamLOCX, Letter_Types_Value(rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RALNAC", Val(Label_AccountRow.Text)), "ChargeOffDate"))
                            End If
                            '*********************************************************************************************************
                        Case Else
                            'iTCS_ReadLine(iLocxRow)
                            'tRALOCX = tSysAccount
                    End Select
                    '**********************************************************
                    '* 2019-03-14 RFK:
                    If swDTable Then
                        sTemp = ReadFieldDataTable(dTable_Select, "ERRORCODE", iLocxRow)
                    Else
                        sTemp = ReadField(DataGridView_Select, "ERRORCODE", iLocxRow)
                    End If
                    If sTemp = "READY" Then
                        '******************************************************
                        '* 2012-11-06 RFK:
                        '******************************************************
                        Select Case sSITE
                            Case "AnnuityOne"
                                '**********************************************
                                '* 2019-03-14 RFK:
                                If swDTable Then
                                    sSysAccountMatched = rkutils.DataTable_ValueByColumnName(dTable_Select, "RAMLOCX", iLocxRow).Trim
                                Else
                                    sSysAccountMatched = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAMLOCX", iLocxRow).Trim
                                End If
                            Case Else
                                'sSysAccountMatched = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "SysAccountMatched", iLocxRow).Trim
                        End Select
                        If Val(sSysAccountMatched) > 0 Then
                            Select Case sSITE
                                Case "AnnuityOne"
                                    '******************************************
                                    '* 2021-12-07 RFK: 10 = NO MATCHED
                                    Select Case gtLetterType
                                        Case "10"
                                            imLocxRow = -1
                                        Case Else
                                            imLocxRow = rkutils.DataGridView_Contains(DataGridView_Multi, "RAMLOCX", sSysAccountMatched)
                                    End Select
                                    '******************************************
                                Case Else
                                    'imLocxRow = rkutils.DataGridView_Contains(DataGridView_Multi, "SYSACCOUNTMATCHED", sSysAccountMatched)
                            End Select
                            '**************************************************
                            If CheckBox_DEBUG.Checked Then MsgStatus("DEBUG/LOCX " + " " + tRALOCX + " RAMLOCX:" + sSysAccountMatched + " LOCXROW:" + imLocxRow.ToString, CheckBox_DEBUG.Checked)
                            '**************************************************
                            '* 2021-12-07 RFK: 10 = NO MATCHED
                            Select Case gtLetterType
                                Case "10"
                                    bOK = True
                                Case Else
                                    bOK = MatchedCheck(iLocxRow, imLocxRow, sSysAccountMatched)
                            End Select
                            '**************************************************
                            If bOK Then
                                Select Case sSITE
                                    Case "AnnuityOne"
                                        '**************************************
                                        '* 2012-10-15 RFK:
                                        Select Case ComboBox_MatchType.Text
                                            Case "Sent Today"
                                                '******************************
                                                '* 2019-03-14 RFK:
                                                If swDTable Then
                                                    tLetterPrinted = rkutils.DataTable_ValueByColumnName(dTable_Select, "RALCAC", iLocxRow).Trim
                                                Else
                                                    tLetterPrinted = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RALCAC", iLocxRow).Trim
                                                End If
                                            Case Else
                                                '******************************
                                                '* 2019-03-14 RFK:
                                                If CheckBox_DEBUG.Checked Then MsgStatus(sSysAccountMatched + " iLocxRow" + iLocxRow.ToString + " [" + tRALOCX + "][" + tRamLOCX + "]", True)
                                                If swDTable Then
                                                    tLetterPrinted = rkutils.DataTable_ValueByColumnName(dTable_Select, "RALNAC", iLocxRow).Trim
                                                Else
                                                    tLetterPrinted = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RALNAC", iLocxRow).Trim
                                                End If
                                                '******************************
                                                '* 2015-08-06 RFK:
                                                Select Case ComboBox_MatchType.Text
                                                    Case "Testing Clients"
                                                        If rkutils.STR_TRIM(ComboBox_TestAll.Text, 1) = "Y" Then
                                                            '**********************************************************
                                                            '* 2019-03-14 RFK:
                                                            '* 2021-11-24 RFK: changed to Valid Letter Types
                                                            iTestLetter += 1
                                                            If iTestLetter >= DataGridView_Letter_Types.RowCount - 1 Then iTestLetter = 0
                                                            tLetterPrinted = rkutils.DataGridView_ValueByColumnName(DataGridView_Letter_Types, "LNumber", iTestLetter).Trim
                                                        Else
                                                            '**********************************************************
                                                            '* 2021-11-26 RFK: Use RALNAC
                                                        End If
                                                End Select
                                        End Select
                                        If CheckBox_DEBUG.Checked Then MsgStatus("DEBUG/iLocxRow=" + iLocxRow.ToString + " " + tRALOCX + " " + sSysAccountMatched + " " + tLetterPrinted, True)
                                        '**************************************
                                        '* : If in PP 
                                        '**********************************************************
                                        '* 2019-03-14 RFK:
                                        If swDTable Then
                                            tStatus = ReadFieldDataTable(dTable_Select, "RARSTA", iLocxRow)
                                        Else
                                            tStatus = ReadField(DataGridView_Select, "RARSTA", iLocxRow)
                                        End If
                                        If STR_LEFT(tStatus, 2) = "PP" Then
                                            Select Case tLetterPrinted
                                                Case "350", "351", "950", "951"
                                                    'Good
                                                Case Else
                                                    '**********************************************************
                                                    '* 2019-03-14 RFK:
                                                    If swDTable Then
                                                        rkutils.DataTable_SetValueByColumnName(dTable_Select, "ERRORCODE", iLocxRow, "PPL_INVALID_LETTER")
                                                    Else
                                                        rkutils.DataGridView_SetValueByColumnName(DataGridView_Select, "ERRORCODE", iLocxRow, "PPL_INVALID_LETTER")
                                                    End If
                                                    '**************************
                                                    '* 2013-10-10 RFK:
                                                    LettersPrinted("PPL_" + tLetterPrinted, tLOCX, True)
                                                    '**************************
                                                    '* 2014-06-09 RFK: 
                                                    If CheckBox_Update.Checked = True Then
                                                        tSQLstring = "UPDATE ROIDATA.RACCTP"
                                                        tSQLstring += " SET RALNAC=0"
                                                        tSQLstring += " WHERE RALOCX=" + tRALOCX
                                                        DB_COMMAND("DB2", DB2SQLConnectionString, DB2SQLuser, tSQLstring)
                                                        '**********************
                                                        TRACKS_update("MSSQL", msSQLConnectionString, msSQLuser, "", tRALOCX, "", "L", "PPL WITHOUT CORRECT LETTER UNABLE [" + tLetterPrinted + "]")
                                                    End If
                                                    '**************************
                                            End Select
                                        Else
                                            '**********************************
                                            '* 2018-01-19 RFK: If NOT in PP but in a PP letter
                                            '* 2018-06-25 RFK: 
                                            Select Case tLetterPrinted
                                                Case "350", "351", "950", "951"
                                                    '**************************
                                                    '* 2018-06-25 RFK: 
                                                    If ComboBox_MatchType.Text = "Testing Clients" Then
                                                        '**********************
                                                        '* 2018-06-25 RFK: Good to Go
                                                    Else
                                                        MsgStatus(ComboBox_MatchType.Text, True)
                                                        '**********************
                                                        '* 2019-03-14 RFK:
                                                        If swDTable Then
                                                            rkutils.DataTable_SetValueByColumnName(dTable_Select, "ERRORCODE", iLocxRow, "PPL_INVALID_LETTER")
                                                        Else
                                                            rkutils.DataGridView_SetValueByColumnName(DataGridView_Select, "ERRORCODE", iLocxRow, "PPL_INVALID_LETTER")
                                                        End If
                                                        '**********************
                                                        '2013-10-10 RFK:
                                                        LettersPrinted("PPL_" + tLetterPrinted, tLOCX, True)
                                                        '**********************
                                                        '* 2014-06-09 RFK: 
                                                        If CheckBox_Update.Checked = True Then
                                                            tSQLstring = "UPDATE ROIDATA.RACCTP"
                                                            tSQLstring += " SET RALNAC=0"
                                                            tSQLstring += " WHERE RALOCX=" + tRALOCX
                                                            DB_COMMAND("DB2", DB2SQLConnectionString, DB2SQLuser, tSQLstring)
                                                            '******************
                                                            TRACKS_update("MSSQL", msSQLConnectionString, msSQLuser, "", tRALOCX, "", "L", "PPL WITHOUT CORRECT LETTER UNABLE [" + tLetterPrinted + "]")
                                                        End If
                                                        '**********************
                                                    End If
                                                    '**************************
                                                Case Else
                                                    'Good
                                            End Select
                                        End If
                                        '**************************************
                                        '* 2019-03-14 RFK:
                                        If swDTable Then
                                            tState = rkutils.DataTable_ValueByColumnName(dTable_Select, "RAGSTATE", iLocxRow).Trim
                                            tStateAllow = rkutils.DataTable_ValueByColumnName(dTable_Select, "STATEALLOW", iLocxRow).Trim
                                            If tState = "XX" Then
                                                rkutils.DataTable_SetValueByColumnName(dTable_Select, "ERRORCODE", iLocxRow, "STATE_BLOCKED")
                                            End If
                                        Else
                                            tState = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGSTATE", iLocxRow).Trim
                                            tStateAllow = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "STATEALLOW", iLocxRow).Trim
                                            If tState = "XX" Then
                                                rkutils.DataGridView_SetValueByColumnName(DataGridView_Select, "ERRORCODE", iLocxRow, "STATE_BLOCKED")
                                            End If
                                        End If
                                        If tState = "XX" Then
                                            LettersPrinted("STATE", tLOCX, True)
                                            '******************************************************************************************
                                            tSQLstring = "UPDATE ROIDATA.RACCTP"
                                            tSQLstring += " SET RALNAC=0, RAGADI='B'"
                                            tSQLstring += " WHERE RAMLOCX='" + sSysAccountMatched + "'"
                                            MsgStatus(tSQLstring, False)
                                            If CheckBox_Update.Checked = True Then
                                                DB_COMMAND("DB2", DB2SQLConnectionString, DB2SQLuser, tSQLstring)
                                                '****************************************************************
                                                TRACKS_update("MSSQL", msSQLConnectionString, msSQLuser, "", tRALOCX, "", "L", "LETTER UNABLE [BAD STATE]")
                                            End If
                                        Else
                                            '**********************************************************************************************
                                            '* 2012-10-02 RFK: State Blocking
                                            If swDTable Then
                                                sTemp = ReadFieldDataTable(dTable_Select, "LETTERS", iLocxRow)
                                            Else
                                                sTemp = ReadField(DataGridView_Select, "LETTERS", iLocxRow)
                                            End If
                                            If State_Block(tLetterPrinted, sTemp, tState) Then
                                                '******************************************************************
                                                '* 2019-03-14 RFK:
                                                If swDTable Then
                                                    rkutils.DataTable_SetValueByColumnName(dTable_Select, "ERRORCODE", iLocxRow, "STATE_BLOCKED")
                                                Else
                                                    rkutils.DataGridView_SetValueByColumnName(DataGridView_Select, "ERRORCODE", iLocxRow, "STATE_BLOCKED")
                                                End If
                                                'Label_StateBlocked.Text = Str(Val(Label_StateBlocked.Text) + 1).Trim
                                                LettersPrinted("STATE", tLOCX, True)
                                                '******************************************************************************************
                                                tSQLstring = "UPDATE ROIDATA.RACCTP"
                                                tSQLstring += " SET RALNAC=0, RAGADI='B'"
                                                tSQLstring += " WHERE RAMLOCX='" + sSysAccountMatched + "'"
                                                MsgStatus(tSQLstring, False)
                                                If CheckBox_Update.Checked = True Then
                                                    DB_COMMAND("DB2", DB2SQLConnectionString, DB2SQLuser, tSQLstring)
                                                    '****************************************************************
                                                    TRACKS_update("MSSQL", msSQLConnectionString, msSQLuser, "", tRALOCX, "", "L", "LETTER UNABLE [STATE BLOCKED]")
                                                End If
                                            Else
                                                '******************************
                                                '* 2012-12-11 RFK:
                                                iRulesRow = DataGridView_Contains(DataGridView_Letter_Types, "LNUMBER", tLetterPrinted)
                                                '******************************
                                                '* 2014-06-05 RFK:
                                                '* 2019-03-14 RFK:
                                                If swDTable Then
                                                    sFacility = rkutils.DataTable_ValueByColumnName(dTable_Select, "RAFACL", iLocxRow).Trim
                                                Else
                                                    sFacility = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAFACL", iLocxRow).Trim
                                                End If
                                                '******************************
                                                iRulesRowDB2 = rkutils.DataGridView_Contains2Cols(DataGridView_RULES, "RRFACL", sFacility, "RRACTV", tLetterPrinted)
                                                If CheckBox_DEBUG.Checked Then MsgStatus("[" + tLetterPrinted + "[" + iRulesRow.ToString + "][" + iRulesRowDB2.ToString + "]", True)
                                                '******************************
                                                If iRulesRow >= 0 And iRulesRowDB2 >= 0 Then
                                                    '**************************
                                                    '* 2014-05-19 RFK: 
                                                    Select Case tLetterPrinted
                                                        Case "350", "950"
                                                            '******************
                                                            Select Case rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text))
                                                                Case Else
                                                                    '**********
                                                                    '* 2014-08-05 RFK: Do Not Look for Broken 1st time
                                                                    If swDTable Then
                                                                        sTemp = ReadFieldDataTable(dTable_Select, "RALCAC", iLocxRow)
                                                                    Else
                                                                        sTemp = ReadField(DataGridView_Select, "RALCAC", iLocxRow)
                                                                    End If
                                                                    Select Case sTemp
                                                                        '******
                                                                        '* 2014-08-05 RFK: already sent a PPL Letter, check to see if they have broken
                                                                        '* 2014-09-04 RFK: SEND ALL PPL
                                                                        'Case "350", "351", "950", "951"
                                                                        Case "XXX"
                                                                            sSQL = "SELECT SUM(RPAMTP) AS SSUM FROM ROIDATA.RPAYP"
                                                                            sSQL += " WHERE RPLOCX IN('" + ListAllLocxByMatchingCriteria(iLocxRow).Replace(" ", "','") + "')"
                                                                            sSQL += " AND DIGITS(RPRYR)||DIGITS(RPRMON)||DIGITS(RPRDAY) > " + rkutils.STR_format(rkutils.STR_DATE_PLUS("TODAY", "-", "30"), "ccyymmdd")
                                                                            sSQL += " AND RPTTYP='PMT'"   'Payment
                                                                            sPPLpaid = rkutils.SQL_READ_FIELD(DataGridView_RPAYP, "DB2", "SSUM", DB2SQLConnectionString, DB2SQLuser, sSQL)
                                                                            sPPLamount = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RATERM", iLocxRow).Trim
                                                                            If (Val(sPPLpaid) * -1) >= Val(sPPLamount) Then
                                                                                MsgStatus("GOOD[" + sPPLamount + "]" + sPPLpaid + "]" + sSQL, True)
                                                                                PrintLine("run.2", iLocxRow, tCurrentLetterVendor, Letter_Types_Value(tLetterPrinted, "LTYPE"), tLetterPrinted, Label_LetterFile.Text, False)
                                                                            Else
                                                                                MsgStatus("BROKEN RALCAC=" + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RALCAC", iLocxRow).Trim + "]PPLamount=" + sPPLamount + "]PPLpaid=" + sPPLpaid + "]" + sSQL, True)
                                                                                sPPLstatus = rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "PPL_BROKEN_STATUS", Val(Label_ClientRow.Text)).Trim
                                                                                If sPPLstatus.Length > 0 Then
                                                                                    MsgStatus(sPPLstatus + " STATUS", True)
                                                                                    If CheckBox_Update.Checked Then rkutils.COMMAND_STATUS(msSQLConnectionString, msSQLuser, tLOCX, sPPLstatus, "", "", "")
                                                                                End If
                                                                                '**************************************************************************************************
                                                                                '* 2014-08-05 RFK: Rules for 351/951
                                                                                '* 2014-08-05 RFK: Only if MinimumLetterAmount setup
                                                                                If Val(ReadField(DataGridView_Clients, "PPL_Broken_MinimumLetterAmount", Val(Label_ClientRow.Text))) > 0 Then
                                                                                    If Val(ReadField(DataGridView_Select, "RAMBAL", iLocxRow)) > Val(ReadField(DataGridView_Clients, "PPL_Broken_MinimumLetterAmount", Val(Label_ClientRow.Text))) Then
                                                                                        '******************************************************************************************
                                                                                        '* 2016-11-09 RFK: corrected for Facility
                                                                                        If swDTable Then
                                                                                            sTemp = ReadFieldDataTable(dTable_Select, "RAFACL", iLocxRow)
                                                                                        Else
                                                                                            sTemp = ReadField(DataGridView_Select, "RAFACL", iLocxRow)
                                                                                        End If
                                                                                        iRulesRow = DataGridView_Contains2Cols(DataGridView_RULES, "RRFACL", sTemp, "RRACTV", tLetterPrinted)
                                                                                        tLetterNext = rkutils.DataGridView_ValueByColumnName(DataGridView_RULES, "RRNACT", iRulesRow).Trim
                                                                                        Select Case tLetterNext
                                                                                            Case "351", "951"
                                                                                                tLetterPrinted = tLetterNext
                                                                                                MsgStatus("Change to " + tLetterPrinted + "", True)
                                                                                                PrintLine("run.3", iLocxRow, tCurrentLetterVendor, Letter_Types_Value(tLetterPrinted, "LTYPE"), tLetterPrinted, Label_LetterFile.Text, False)
                                                                                        End Select
                                                                                    End If
                                                                                End If
                                                                                '**************************************************************************************************
                                                                            End If
                                                                        Case Else
                                                                            If CheckBox_DEBUG.Checked Then MsgStatus("INITIAL PPL LETTER [" + tLetterPrinted + "][" + tLOCX + "]", False)
                                                                            PrintLine("run.4", iLocxRow, tCurrentLetterVendor, Letter_Types_Value(tLetterPrinted, "LTYPE"), tLetterPrinted, Label_LetterFile.Text, False)
                                                                    End Select
                                                            End Select
                                                        Case Else
                                                            If swDTable Then
                                                                sLastLetterDate = rkutils.ReadFieldDataTable(dTable_Select, "RASDATE", iLocxRow)
                                                            Else
                                                                sLastLetterDate = rkutils.ReadField(DataGridView_Select, "RASDATE", iLocxRow)
                                                            End If
                                                            '**********************************************************************************************************************
                                                            '* 2016-11-09 RFK: corrected for Facility
                                                            If swDTable Then
                                                                sTemp = ReadFieldDataTable(dTable_Select, "RAFACL", iLocxRow)
                                                            Else
                                                                sTemp = ReadField(DataGridView_Select, "RAFCL", iLocxRow)
                                                            End If
                                                            iRulesRow = DataGridView_Contains2Cols(DataGridView_RULES, "RRFACL", sTemp, "RRACTV", tLetterPrinted)
                                                            '**********************************************************************************************************************
                                                            tLetterNextDays = rkutils.DataGridView_ValueByColumnName(DataGridView_RULES, "RRDAYS", iRulesRow).Trim
                                                            If CheckBox_DEBUG.Checked Then MsgStatus("tLetterPrinted=" + tLetterPrinted + "] iRulesRow=" + iRulesRow.ToString + "] iRulesRowDB2=" + iRulesRowDB2.ToString + "] sLastLetterDate=" + sLastLetterDate + "] tLetterNextDays=" + tLetterNextDays + "]", True)
                                                            If sLastLetterDate.Trim.Length >= 8 Then
                                                                sOKdate = rkutils.STR_format(rkutils.STR_DATE_PLUS(sLastLetterDate.Substring(4, 2) + "/" + sLastLetterDate.Substring(6, 2) + "/" + sLastLetterDate.Substring(0, 4), "+", tLetterNextDays), "ccyymmdd")
                                                                If sOKdate.Trim.Length >= 8 And sOKdate <= rkutils.STR_format("TODAY", "ccyymmdd") Then
                                                                    PrintLine("run.5", iLocxRow, tCurrentLetterVendor, Letter_Types_Value(tLetterPrinted, "LTYPE"), tLetterPrinted, Label_LetterFile.Text, False)
                                                                Else
                                                                    '**************************************************************************************************************
                                                                    '* 2016-11-02 RFK:
                                                                    '* 2017-01-18 RFK: CheckBox_BypassDate
                                                                    If CheckBox_BypassDate.Checked Then
                                                                        MsgStatus("NOT OK DATE ALLOWED TO BYPASSRULESDATE LOCX:" + tLOCX + "LETTER [" + tLetterPrinted + "] ByPassRulesDate:" + Letter_Types_Value(tLetterPrinted, "BypassRulesDate") + " LastLetterDate:" + sLastLetterDate, True)
                                                                        PrintLine("run.6", iLocxRow, tCurrentLetterVendor, Letter_Types_Value(tLetterPrinted, "LTYPE"), tLetterPrinted, Label_LetterFile.Text, False)
                                                                        If CheckBox_Update.Checked Then TRACKS_update("MSSQL", msSQLConnectionString, msSQLuser, "", tLOCX, "", "L", "NOT OK DATE ALLOWED TO BYPASSRULESDATE LOCX:" + tLOCX + "LETTER [" + tLetterPrinted + "] ByPassRulesDate:" + Letter_Types_Value(tLetterPrinted, "BypassRulesDate") + " LastLetterDate:" + sLastLetterDate)
                                                                    Else
                                                                        Select Case Letter_Types_Value(tLetterPrinted, "BypassRulesDate")
                                                                            Case "Y"
                                                                                MsgStatus("NOT OK DATE ALLOWED TO BYPASSRULESDATE LOCX:" + tLOCX + "LETTER [" + tLetterPrinted + "] ByPassRulesDate:" + Letter_Types_Value(tLetterPrinted, "BypassRulesDate") + " LastLetterDate:" + sLastLetterDate, True)
                                                                                PrintLine("run.7", iLocxRow, tCurrentLetterVendor, Letter_Types_Value(tLetterPrinted, "LTYPE"), tLetterPrinted, Label_LetterFile.Text, False)
                                                                                If CheckBox_Update.Checked Then TRACKS_update("MSSQL", msSQLConnectionString, msSQLuser, "", tLOCX, "", "L", "NOT OK DATE ALLOWED TO BYPASSRULESDATE LOCX:" + tLOCX + "LETTER [" + tLetterPrinted + "] ByPassRulesDate:" + Letter_Types_Value(tLetterPrinted, "BypassRulesDate") + " LastLetterDate:" + sLastLetterDate)
                                                                            Case Else
                                                                                '******************************************************************************************************
                                                                                '* 2016-11-28 RFK: 
                                                                                Select Case ComboBox_MatchType.Text
                                                                                    Case "Sent Today"
                                                                                        '**********************************************************************************************
                                                                                        '* 2016-11-28 RFK: ONLY TODAYS DATE SHOULD BE SELECTED
                                                                                        PrintLine("run.8", iLocxRow, tCurrentLetterVendor, Letter_Types_Value(tLetterPrinted, "LTYPE"), tLetterPrinted, Label_LetterFile.Text, False)
                                                                                        If CheckBox_Update.Checked Then TRACKS_update("MSSQL", msSQLConnectionString, msSQLuser, "", tLOCX, "", "L", "NOT OK DATE ALLOWED TO BYPASSRULESDATE LOCX:" + tLOCX + "LETTER [" + tLetterPrinted + "] ByPassRulesDate:" + Letter_Types_Value(tLetterPrinted, "BypassRulesDate") + " LastLetterDate:" + sLastLetterDate)
                                                                                    Case Else
                                                                                        MsgStatus("NOT OK DATE [" + tLetterPrinted + "] Locx:[" + tLOCX + "] LastLetterDate:" + sLastLetterDate + " LetterNextDays:" + tLetterNextDays + " OKDate:" + sOKdate, True)
                                                                                        sSQL = "UPDATE ROIDATA.RACCTP SET RANLMO=" + sOKdate.Substring(4, 2)
                                                                                        sSQL += ",RANLDY=" + sOKdate.Substring(6, 2)
                                                                                        sSQL += ",RANLYR=" + sOKdate.Substring(0, 4)
                                                                                        sSQL += " WHERE RALOCX=" + tLOCX
                                                                                        MsgStatus(sSQL, False)
                                                                                        If CheckBox_Update.Checked Then rkutils.DB_COMMAND("DB2", DB2SQLConnectionString, DB2SQLuser, sSQL)
                                                                                        If CheckBox_Update.Checked Then TRACKS_update("MSSQL", msSQLConnectionString, msSQLuser, "", tLOCX, "", "L", "LETTER DATE NOT OK [" + tLetterPrinted + "] Locx:[" + tLOCX + "] LastLetterDate:" + sLastLetterDate + " LetterNextDays:" + tLetterNextDays + " OKDate:" + sOKdate)
                                                                                        '**********************************************************************************************
                                                                                End Select
                                                                        End Select
                                                                    End If
                                                                End If
                                                            Else
                                                                'MsgStatus(tLOCX + " LastLetterDate:" + sLastLetterDate + "] send it", True)
                                                                PrintLine("run.9", iLocxRow, tCurrentLetterVendor, Letter_Types_Value(tLetterPrinted, "LTYPE"), tLetterPrinted, Label_LetterFile.Text, False)
                                                            End If
                                                    End Select
                                                Else
                                                    '**************************
                                                    '* 2013-12-26 RFK: Substitute Number
                                                    '******************************************************************
                                                    '* 2019-03-14 RFK:
                                                    If swDTable Then
                                                        rkutils.DataTable_SetValueByColumnName(dTable_Select, "ERRORCODE", iLocxRow, "INVALID_LETTER [" + iRulesRow.ToString + "][" + iRulesRowDB2.ToString + "]Facility=" + sFacility + "][RRACTV=" + tLetterPrinted + "]")
                                                    Else
                                                        rkutils.DataGridView_SetValueByColumnName(DataGridView_Select, "ERRORCODE", iLocxRow, "INVALID_LETTER [" + iRulesRow.ToString + "][" + iRulesRowDB2.ToString + "]Facility=" + sFacility + "][RRACTV=" + tLetterPrinted + "]")
                                                    End If
                                                    '****************************************************************
                                                    '2013-11-26 RFK:
                                                    MsgStatus("INVALID_LETTER [" + iRulesRow.ToString + "][" + iRulesRowDB2.ToString + "]Facility=" + sFacility + "][RRACTV=" + tLetterPrinted + "]", True)
                                                    LettersPrinted(tLetterPrinted, tLOCX, True)
                                                    '************************************************************************************
                                                    '* 2014-06-09 RFK: 
                                                    If CheckBox_Update.Checked = True Then
                                                        tSQLstring = "UPDATE ROIDATA.RACCTP"
                                                        tSQLstring += " SET RALNAC=0"
                                                        tSQLstring += " WHERE RALOCX=" + tRALOCX
                                                        DB_COMMAND("DB2", DB2SQLConnectionString, DB2SQLuser, tSQLstring)
                                                        '********************************************************************************
                                                        TRACKS_update("MSSQL", msSQLConnectionString, msSQLuser, "", tRALOCX, "", "L", "LETTER UNABLE [INVALID LETTER " + tLetterPrinted + "]")
                                                    End If
                                                    '************************************************************************************
                                                End If
                                                '******************************
                                                '* 2012-11-12 RFK:
                                            End If
                                        End If
                                        '**************************************
                                    Case "iTeleCollect"
                                        '**************************************
                                        '* 2012-10-15 RFK:
                                        '* 2015-08-14 RFK:
                                        Select Case TextBox_TCodes.Text
                                            Case "L0"   '1st Letter
                                                tLetterPrinted = "001"
                                            Case "L2"   '2nd Letter
                                                tLetterPrinted = "005"
                                            Case "L3"   'Denial
                                                tLetterPrinted = "002"
                                            Case "L5"   'Final
                                                tLetterPrinted = "009"
                                            Case "L6"   'No Phone Number
                                                tLetterPrinted = "004"
                                            Case "L7"   'Follow Up
                                                tLetterPrinted = "003"
                                            Case "L8"   'Patient Guarantor
                                                tLetterPrinted = "008"
                                            Case Else
                                                MsgStatus("Invalid TCode", True)
                                        End Select
                                        'MsgStatus(tLetterPrinted, True)
                                        '**************************************
                                        '* 2019-03-14 RFK:
                                        If swDTable Then
                                            tState = rkutils.ReadFieldDataTable(dTable_Select, "GSTATE", iLocxRow)
                                        Else
                                            tState = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "GSTATE", iLocxRow).Trim
                                        End If
                                        If tState = "XX" Then
                                            '******************************************************************
                                            '* 2019-03-14 RFK:
                                            If swDTable Then
                                                rkutils.DataTable_SetValueByColumnName(dTable_Select, "ERRORCODE", iLocxRow, "STATE_BLOCKED")
                                            Else
                                                rkutils.DataGridView_SetValueByColumnName(DataGridView_Select, "ERRORCODE", iLocxRow, "STATE_BLOCKED")
                                            End If
                                            '**********************************
                                            tSQLstring = "UPDATE ROIDATA.RACCTP"
                                            tSQLstring += " SET RALNAC=0, RAGADI='B'"
                                            tSQLstring += " WHERE RAMLOCX='" + sSysAccountMatched + "'"
                                            If CheckBox_Update.Checked = True Then
                                                'DB_COMMAND("DB2", DB2SQLConnectionString, DB2SQLuser, tSQLstring)
                                                '******************************
                                                TRACKS_update("MSSQL", msSQLConnectionString, msSQLuser, "", tRALOCX, "", "L", "LETTER UNABLE [BAD STATE]")
                                            Else
                                                MsgStatus(tSQLstring, False)
                                            End If
                                        Else
                                            '**********************************
                                            '* 2012-10-02 RFK: State Blocking
                                            If State_Block(tLetterPrinted, tStateAllow, tState) Then
                                                '******************************
                                                '* 2019-03-14 RFK:
                                                If swDTable Then
                                                    rkutils.DataTable_SetValueByColumnName(dTable_Select, "ERRORCODE", iLocxRow, "STATE_BLOCKED")
                                                Else
                                                    rkutils.DataGridView_SetValueByColumnName(DataGridView_Select, "ERRORCODE", iLocxRow, "STATE_BLOCKED")
                                                End If
                                                '******************************
                                                tSQLstring = "UPDATE ROIDATA.RACCTP"
                                                tSQLstring += " SET RALNAC=0, RAGADI='B'"
                                                tSQLstring += " WHERE RAMLOCX='" + sSysAccountMatched + "'"
                                                If CheckBox_Update.Checked = True Then
                                                    'DB_COMMAND("DB2", DB2SQLConnectionString, DB2SQLuser, tSQLstring)
                                                    '****************************************************************
                                                    TRACKS_update("MSSQL", msSQLConnectionString, msSQLuser, "", tRALOCX, "", "L", "LETTER UNABLE [STATE BLOCKED]")
                                                Else
                                                    MsgStatus(tSQLstring, False)
                                                End If
                                            Else
                                                '******************************************************************************************
                                                '* 2012-12-11 RFK:
                                                iRulesRow = DataGridView_Contains(DataGridView_Letter_Types, "LNUMBER", tLetterPrinted)
                                                If iRulesRow >= 0 Then
                                                    MsgStatus("LETTER [" + tLetterPrinted + "][" + tLOCX + "]", True)
                                                    PrintLine("run.10", iLocxRow, tCurrentLetterVendor, Letter_Types_Value(tLetterPrinted, "LTYPE"), tLetterPrinted, Label_LetterFile.Text, False)
                                                Else
                                                    MsgStatus("Invalid Letter [" + tLetterPrinted + "]", True)
                                                    '******************************************************************
                                                    '* 2019-03-14 RFK:
                                                    If swDTable Then
                                                        rkutils.DataTable_SetValueByColumnName(dTable_Select, "ERRORCODE", iLocxRow, "INVALID_LETTER")
                                                    Else
                                                        rkutils.DataGridView_SetValueByColumnName(DataGridView_Select, "ERRORCODE", iLocxRow, "INVALID_LETTER")
                                                    End If
                                                    '****************************************************************
                                                    '2013-11-26 RFK:
                                                    LettersPrinted(tLetterPrinted, tLOCX, True)
                                                    '************************************************************************************
                                                    '* 2014-06-09 RFK: 
                                                    If CheckBox_Update.Checked = True Then
                                                        tSQLstring = "UPDATE ROIDATA.RACCTP"
                                                        tSQLstring += " SET RALNAC=0"
                                                        tSQLstring += " WHERE RALOCX=" + tRALOCX
                                                        DB_COMMAND("DB2", DB2SQLConnectionString, DB2SQLuser, tSQLstring)
                                                        '********************************************************************************
                                                        TRACKS_update("MSSQL", msSQLConnectionString, msSQLuser, "", tRALOCX, "", "L", "LETTER UNABLE [INVALID LETTER " + tLetterPrinted + "]")
                                                    End If
                                                    '************************************************************************************
                                                End If
                                                '******************************************************************************************
                                                '* 2012-11-12 RFK:
                                            End If
                                        End If
                                    Case Else
                                        MsgStatus("Run/Default/Site", True)
                                End Select
                            Else
                                '**********************************************
                                '* 2019-03-14 RFK:
                                If swDTable Then
                                    rkutils.DataTable_SetValueByColumnName(dTable_Select, "ERRORCODE", iLocxRow, "FAILED_MATCHCHECK")
                                Else
                                    rkutils.DataGridView_SetValueByColumnName(DataGridView_Select, "ERRORCODE", iLocxRow, "FAILED_MATCHCHECK")
                                End If
                                MsgStatus("FAILED_MATCHCHECK", CheckBox_DEBUG.Checked)
                            End If
                        Else
                            '**********************************************
                            '* 2019-03-14 RFK:
                            If swDTable Then
                                rkutils.DataTable_SetValueByColumnName(dTable_Select, "ERRORCODE", iLocxRow, "FAILED_RAMLOCX")
                            Else
                                rkutils.DataGridView_SetValueByColumnName(DataGridView_Select, "ERRORCODE", iLocxRow, "FAILED_RAMLOCX")
                            End If
                            MsgStatus("FAILED_RAMLOCX", CheckBox_DEBUG.Checked)
                        End If
                    Else
                        'MsgStatus("debug:" + sTemp, True)
                    End If
                    '**********************************************************
                Else
                    '******************************************************************
                    '* 2019-03-14 RFK:
                    If swDTable Then
                        Select Case ReadFieldDataTable(dTable_Select, "ERRORCODE", iLocxRow)
                            Case "MULTI"
                                MsgStatus("MULTI_ALREADY/" + ReadFieldDataTable(dTable_Select, "RALOCX", iLocxRow), False)
                            Case Else
                                MsgStatus("FAILED/" + ReadFieldDataTable(dTable_Select, "ERRORCODE", iLocxRow) + "/" + ReadFieldDataTable(dTable_Select, "RALOCX", iLocxRow), False)
                        End Select
                    Else
                        Select Case ReadField(DataGridView_Select, "ERRORCODE", iLocxRow)
                            Case "MULTI"
                                MsgStatus("MULTI_ALREADY/" + ReadField(DataGridView_Select, "RALOCX", iLocxRow), False)
                            Case Else
                                MsgStatus("FAILED/" + ReadField(DataGridView_Select, "ERRORCODE", iLocxRow) + "/" + ReadField(DataGridView_Select, "RALOCX", iLocxRow), False)
                        End Select
                    End If
                End If
                '**************************************************************
                iLocxRow += 1
                '**************************************************************
                If Val(Label_Printed.Text) >= Val(TextBox_MaxLetters.Text) Then
                    MsgStatus("Limited to " + TextBox_MaxLetters.Text, True)
                    Label_RUNNING.Text = "Limited"
                End If
                '**************************************************************
            Loop
            PrintTrailer(tCurrentLetterVendor, Label_LetterFile.Text)
            Label_AccountsRemaining.Text = "0"
            '******************************************************************
            If ListBox_Letters.Items.Count > 0 And Val(Label_Printed.Text) > 0 Then
                If CheckBox_Update.Checked = True Then
                    If Val(Label_Printed.Text) > 0 Then FTP_put()
                    '**********************************************************
                    '* 2014-09-25 RFK:
                    sSQL = "INSERT INTO " + sDBO + ".dbo.LetterVendorFile"
                    sSQL += " (RevMDFile"
                    sSQL += ",RevMDSentDate"
                    sSQL += ",RevMDLetterVendor"
                    sSQL += ",RevMDRecords"
                    sSQL += ",RevMDRecordsMatched"
                    sSQL += ",RevMDRecordsBadAddress"
                    'sSQL += ",[ReceivedDate]"
                    'sSQL += ",[VendorDate]"
                    'sSQL += ",[VendorFile]"
                    'sSQL += ",[RecordsProcessed]"
                    'sSQL += ",[RecordsRejected]"
                    'sSQL += ",[RecordsLoaded]"
                    'sSQL += ",[RecordsInTest]"
                    sSQL += ")"
                    sSQL += " VALUES("
                    sSQL += "'" + Path.GetFileName(Label_LetterFile.Text) + "'"
                    sSQL += ",'" + STR_format("TODAY", "mm/dd/ccyy HH:MM:SS") + "'"
                    sSQL += ",'" + tCurrentLetterVendor + "'"
                    sSQL += "," + Label_Printed.Text
                    sSQL += "," + Label_PrintedM.Text
                    sSQL += "," + Label_BadAddress.Text
                    sSQL += ")"
                    rkutils.DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, sSQL)
                End If
                '**************************************************************
            End If
            '******************************************************************
            EMail_Summary()
            '******************************************************************
            '* 2014-08-04 RFK:
            sSQL = "UPDATE " + sDBO + ".dbo.clients"
            sSQL += " SET LettersRunDate='" + rkutils.STR_format("TODAY", "mm/dd/ccyy HH:MM:SS") + "'"
            sSQL += ",LettersRunTotal='" + Label_Printed.Text + "'"
            sSQL += " WHERE ClientName='" + rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)) + "'"
            rkutils.DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, sSQL)
            '******************************************************************
            '* 2018-03-07 RFK:
            Dim sMessage As String = ""
            For i1 = 0 To ListBox_Printed.Items.Count - 1
                '**************************************************************
                If Listbox_Contains(ListBox_Noted, ListBox_Printed.Items(i1).ToString, False) < 0 Then
                    sMessage += ListBox_Printed.Items(i1).ToString + vbCrLf
                    MsgStatus("WARNING/NOT_NOTED/" + ListBox_Printed.Items(i1).ToString, True)
                End If
            Next
            If sMessage.Length > 0 Then rkutils.EMAILIT(msSQLConnectionString, msSQLuser, "eLetters_DoNotReply@AnnuityHealth.com", "eLetters", "RYAN@AnnuityHealth.com", "IT", "", "eLetters/ERROR/NOT_NOTED", sMessage, "", "")
            '******************************************************************
            RunReady()
            MsgStatus("Completed:" + Label_Printed.Text, True)
            '******************************************************************
        Catch ex As Exception
            MsgError("Run", ex.ToString)
        End Try
    End Sub

    Private Sub RunALL()
        Try
            Dim iClientRow As Integer = 0
            '******************************************************************
            If Me.CheckBox_Update.Checked = True Then
                Label_SummaryFile.Text = dir_LETTERS + rkutils.STR_format("TODAY", "ccyymmdd") + "_" + rkutils.STR_format("TODAY", "HH") + "_LETTERS.TXT"
            Else
                Label_SummaryFile.Text = dir_TEST + rkutils.STR_format("TODAY", "ccyymmdd") + "_" + rkutils.STR_format("TODAY", "HH") + "_LETTERS.TXT"
            End If
            '******************************************************************
            '* 2013-01-07 RFK: 
            MsgStatus("RunALL", True)
            Label_ClientRunning.Text = "Running"
            Do While iClientRow < DataGridView_Clients.Rows.Count - 1 And Label_ClientRunning.Text = "Running"
                Label_ClientRow.Text = iClientRow.ToString
                DataGridView_Clients.CurrentCell = DataGridView_Clients.Rows(Val(Label_ClientRow.Text)).Cells(0)
                '**************************************************************
                AccountInitSettings()
                Application.DoEvents()
                '**************************************************************
                '* 2021-07-12 RFK: (Check for Match Complete TODAY)
                If DataGridView_Clients.Item(rkutils.DataGridView_ColumnByName(DataGridView_Clients, "ClientName"), iRow).Style.BackColor = Color.Red Then
                    '**********************************************************
                    '* 2021-07-12 RFK:
                    rkutils.EMAILIT(msSQLConnectionString, msSQLuser, "DoNotReply@AnnuityHealth.com", "eLetters", "RKiechle@AnnuityHealth.com", "Letters", Me.Text, "eLetters", "eLetters unable to run for:" + rkutils.ReadField(DataGridView_Clients, "ClientName", iClientRow), "", "")
                    '**********************************************************
                Else
                    '**********************************************************
                    If Panel_RegF.Visible Then
                        '******************************************************
                        Select Case sSITE
                            Case "AnnuityOne"
                                AnnuityOne_AccountsLoad(rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "TOB", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "FACILITY", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTMATCHTYPE", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "MATCHBY", Val(Label_ClientRow.Text)), gtLetterType, "*", 1)
                            Case Else
                                '**********************************************
                        End Select
                        '******************************************************
                        Application.DoEvents()
                        Run()
                        If Val(Label_Printed.Text) > 0 Then FTP_put()
                        '******************************************************
                        Select Case sSITE
                            Case "AnnuityOne"
                                AnnuityOne_AccountsLoad(rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "TOB", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "FACILITY", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTMATCHTYPE", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "MATCHBY", Val(Label_ClientRow.Text)), gtLetterType, "*", 2)
                            Case Else
                                '**********************************************
                        End Select
                        '******************************************************
                        Application.DoEvents()
                        Run()
                        If Val(Label_Printed.Text) > 0 Then FTP_put()
                        '******************************************************
                    Else
                        AccountsLoadForClient()
                        Application.DoEvents()
                        Run()
                        If Val(Label_Printed.Text) > 0 Then FTP_put()
                    End If
                    '**********************************************************
                End If
                '**************************************************************
                iClientRow += 1
            Loop
            If IS_File(Label_SummaryFile.Text) Then
                Dim tMSG As String = File.ReadAllText(Label_SummaryFile.Text)
                Select Case sSITE
                    Case "AnnuityOne"
                        'rkutils.EMAILIT(msSQLConnectionString, msSQLuser, "DoNotReply@AnnuityHealth.com", "eLetters", TextBox_SummaryEMail.Text, "Letter Results", Me.Text, "eLetter Summary [" + rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)) + "]", tMSG, "", "")
                    Case Else
                        rkutils.EMAILIT(msSQLConnectionString, msSQLuser, "DoNotReply@FosterTech.net", "eLetters", TextBox_SummaryEMail.Text, "Letter Results", Me.Text, "eLetter Summary [" + rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)) + "]", tMSG, "", "")
                End Select
            End If
            RunReady()
            Label_ClientRunning.Text = "Ready"
        Catch ex As Exception
            MsgError("RunALL", ex.ToString)
        End Try
    End Sub

    Private Sub LettersPreCalc()
        Try
            '******************************************************************
            '* 2015-12-15 RFK: PreCalc number to be sent
            Dim sLetter As String = ""
            ListBox_Letters.Items.Clear()
            For i1 = 0 To DataGridView_Select.RowCount - 1
                sLetter = rkutils.ReadField(DataGridView_Select, "RALNAC", i1)
                If sLetter.Length > 0 Then
                    '**********************************************************
                    '* 2015-12-15 RFK: Only Count Master Accounts
                    If rkutils.ReadField(DataGridView_Select, "RALOCX", i1) = rkutils.ReadField(DataGridView_Select, "RAMLOCX", i1) Then
                        iLetterRow = Listbox_Contains(ListBox_Letters, sLetter, False)
                        If iLetterRow >= 0 Then
                            ListBox_Letters.Items(iLetterRow) = STR_BREAK(ListBox_Letters.Items(iLetterRow).ToString, 1) + " " + Str(Val(STR_BREAK(ListBox_Letters.Items(iLetterRow).ToString, 2)) + 1).Trim
                        Else
                            ListBox_Letters.Items.Add(sLetter + " 1")
                        End If
                    End If
                    '**********************************************************
                End If
            Next
            '******************************************************************
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button_LoadPP_Click(sender As Object, e As EventArgs) Handles Button_LoadPP.Click
        Try
            If Panel_RegF.Visible Then
                Button_Run.Enabled = False
                Button_Run.Text = "Please wait"
                '******************************************************************
                Select Case sSITE
                    Case "AnnuityOne"
                        AnnuityOne_AccountsLoad(rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "TOB", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "FACILITY", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTMATCHTYPE", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "MATCHBY", Val(Label_ClientRow.Text)), gtLetterType, "*", 3)
                    Case Else
                        '**********************************************************
                        MsgStatus(sSITE, True)
                End Select
                Button_Run.Text = "Run"
                Button_Run.Enabled = True
                '******************************************************************
                LettersPreCalc()
                '******************************************************************
            End If
        Catch ex As Exception
            MsgError("Button_LoadRegFnot_Click", ex.ToString)
        End Try
    End Sub

    Private Sub PrintLine(ByVal sModule As String, ByVal iLine As Integer, ByVal sLetterVendor As String, ByVal sLetterType As String, ByVal sLetterNumber As String, ByVal sFileName As String, ByVal bGhost As Boolean)
        Try
            '******************************************************************
            '* 2012-06-21 RFK:
            '* 2012-11-12 RFK:
            '******************************************************************
            Dim tClient As String = ""
            If ReadField(DataGridView_Select, "ERRORCODE", iLine) <> "READY" Then
                MsgStatus("PrintLine/ERROR/NOT_READY/iLine:" + iLine.ToString + " sLetterVendor:" + sLetterVendor + " sLetterType:" + sLetterType + " sLetterNumber:" + sLetterNumber + " RALOCX:" + ReadField(DataGridView_Select, "RALOCX", iLine) + " RAMLOCX:" + ReadField(DataGridView_Select, "RAMLOCX", iLine) + " NAME:" + ReadField(DataGridView_Select, "RAGLNM", iLine) + " " + ReadField(DataGridView_Select, "RAGFNM", iLine), bGhost)
                Exit Sub
            End If
            If CheckBox_DEBUG.Checked Then MsgStatus("PrintLine iLine:" + iLine.ToString + " sLetterVendor:" + sLetterVendor + " sLetterType:" + sLetterType + " sLetterNumber:" + sLetterNumber + " RALOCX:" + ReadField(DataGridView_Select, "RALOCX", iLine) + " RAMLOCX:" + ReadField(DataGridView_Select, "RAMLOCX", iLine) + " NAME:" + ReadField(DataGridView_Select, "RAGLNM", iLine) + " " + ReadField(DataGridView_Select, "RAGFNM", iLine), CheckBox_DEBUG.Checked)
            '******************************************************************
            PrintLine_tLINE = ""
            '******************************************************************
            '* 2015-07-28 RFK:
            Select Case sSITE
                Case "AnnuityOne"
                    If bGhost Then
                        tSysAccount = rkutils.DataGridView_ValueByColumnName(DataGridView_Ghosts, "LOCX", iLine).Trim
                    Else
                        tSysAccount = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RALOCX", iLine).Trim
                    End If
                Case "iTeleCollect"
                    'tSysAccount = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "SYSACCOUNT", iLine).Trim
            End Select
            '****************************************************************
            '* 2013-01-07 RFK:
            If bGhost = False Then
                If Listbox_Contains(ListBox_Printed, tSysAccount, False) >= 0 Then
                    MsgStatus("PrintLine/failed/" + sModule + "/SysAccount:" + tSysAccount + "]", True)
                    File.AppendAllText(FileNameRPT, "PrintLine/failed/" + sModule + "/SysAccount:" + tSysAccount + "]" + vbCrLf)
                    Exit Sub
                End If
            End If
            '******************************************************************
            Select Case sLetterVendor
                Case "ACCUDOC", "DANTOM", "REVSPRING"
                    '**********************************************************
                    '* 2013-07-16 RFK: Calculate only the ones being sent over in this file, selection 
                    '* 2015-07-28 RFK:
                    Select Case sSITE
                        Case "AnnuityOne"
                            sSentBalance = CalculateSentBalance(ReadField(DataGridView_Select, "RAMLOCX", iLine), sLetterVendor, sLetterType, sLetterNumber, iLine)
                        Case "iTeleCollect"
                            'sSentBalance = CalculateSentBalance(ReadField(DataGridView_Select, "SYSACCOUNT", iLine), sLetterVendor, sLetterType, sLetterNumber)
                    End Select
                    '**********************************************************
                    '* 2012-10-03 RFK:
                    If bGhost = False Then letter_sent("PrintLine.1", False, iLine, sLetterNumber, sSentBalance, False)
                    tAddress = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGADD", iLine).Trim
                    tAddress2 = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGAD2", iLine).Trim
                    tCity = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGCST", iLine).Trim + " "
                    tZip = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGZIP", iLine).Trim.PadLeft(5, "0")
                    tZip4 = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGZP2", iLine).Trim
                    If tAddress.Length < 2 Or tCity.Length < 5 Or Val(tZip) < 1 Then
                        rkutils.DataGridView_SetValueByColumnName(DataGridView_Select, "ERRORCODE", iLine, "BAD_ADDRESS")
                        Label_BadAddress.Text = Str(Val(Label_BadAddress.Text) + 1).Trim
                        '******************************************************
                        If CheckBox_Update.Checked = True Then
                            sSQL = "UPDATE ROIDATA.RACCTP"
                            sSQL += " SET RALNAC=0, RAGADI='B'"
                            sSQL += " WHERE RALOCX=" + tRALOCX
                            DB_COMMAND("DB2", DB2SQLConnectionString, DB2SQLuser, sSQL)
                            '**************************************************
                            TRACKS_update("MSSQL", msSQLConnectionString, msSQLuser, "", tRALOCX, "", "L", "LETTER UNABLE [BAD ADDRESS]")
                        End If
                        '******************************************************
                        Exit Sub
                        '******************************************************
                    End If
                    '**********************************************************
                    tDelimiter = ";"
                    PrintLine_tLINE += "#K#02"   'NCOA               "#K#03^"   'NO NCOA
                    PrintLine_tLINE += tSysAccount + tDelimiter                            'LOCX
                    '**********************************************************
                    If sLetterNumber.Trim.Length = 0 Or Val(sLetterNumber) <= 0 Then
                        Printed(iLine, "", "BAD_LETTERCODE")    '* Printed Error
                        Exit Sub
                    End If
                    '**********************************************
                    '* 2013-03-21 RFK:
                    tClient = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RACL#", iLine).Trim
                    '*************************************************************************
                    '* 2013-07-19 RFK: Shutdown ClientGroup until DANTOM can fix client number 
                    sClientGroup = ""   'Letter_Types_Value(sLetterNumber, "CLIENTGROUP")
                    If sClientGroup.Length > 0 And sClientGroup <> "*" Then
                        PrintLine_tLINE += sClientGroup + sLetterNumber + tDelimiter
                    Else
                        Select Case tClient
                            Case "COP"
                                Select Case rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAPAYR", iLine).Trim
                                    Case "MANAGED"
                                        Select Case sLetterNumber
                                            Case "310"  'Override to MANAGED LETTER
                                                PrintLine_tLINE += tClient + "311" + tDelimiter
                                            Case Else
                                                PrintLine_tLINE += tClient + sLetterNumber + tDelimiter
                                        End Select
                                    Case Else
                                        PrintLine_tLINE += tClient + sLetterNumber + tDelimiter
                                End Select
                            Case Else
                                PrintLine_tLINE += tClient + sLetterNumber + tDelimiter
                        End Select
                    End If
                    '**********************************************************
                    '* 2017-03-01 RFK: Guarantor is a Minor
                    Dim sTemp As String = rkutils.ReadField(DataGridView_Select, "RAGBMO", iLine) + "/" + rkutils.ReadField(DataGridView_Select, "RAGBDY", iLine) + "/" + rkutils.ReadField(DataGridView_Select, "RAGBYR", iLine)
                    If IsDate(sTemp) Then
                        Dim dtemp As Date = sTemp
                        Dim dAge As New Date(Now.Subtract(dtemp).Ticks)
                        If (dAge.Year - 1) < 18 Then
                            MsgStatus("Guarantor is a Minor [" + Str(dAge.Year - 1) + "years " + Str(dAge.Month - 1) + " days ", True)
                            If PrintLine_tLINE.ToUpper.Contains("GUARDIAN") Or PrintLine_tLINE.ToUpper.Contains("PARENT") Then
                                MsgStatus("Already contains Guardian or Parent", True)
                            End If
                            PrintLine_tLINE += "Guardian of "
                            'rkutils.EMAILIT(msSQLConnectionString, msSQLuser, "DoNotReply@AnnuityHealth.com", "eLetters", "LETTERS@AnnuityHealth.com", "Letters", Me.Text, "eLetters Guarantor Minor", "LOCX:" + tLOCX, "", "")
                        End If
                    End If
                    '**********************************************************
                    '* Guarantor Name   
                    '* 2013-10-09 RFK: Changed to First, Middle, Last (Jim from RevSpring)
                    '* 2015-07-21 RFK: 
                    If ReadField(DataGridView_Select, "RAGFNM", iLine).Trim.Length = 0 Then
                        Select Case rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "ELETTERS", Val(Label_ClientRow.Text))
                            Case "t"
                                '******************************************
                                '* 2015-07-21 RFK: 
                                PrintLine_tLINE += "TEST"
                            Case Else
                                rkutils.DataGridView_SetValueByColumnName(DataGridView_Select, "ERRORCODE", iLine, "INVALID_GUARANTOR")
                                LettersPrinted("GUAR", tRALOCX, True)
                                If CheckBox_Update.Checked Then rkutils.COMMAND_STATUS(msSQLConnectionString, msSQLuser, tRALOCX, sStatusCodeBadGuarantorName, "", "", "")
                        End Select
                    Else
                        If bGhost Then
                            PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Ghosts, "FNAME", iGhosts_CTR).Trim
                        Else
                            PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGFNM", iLine).Trim
                        End If
                    End If
                    '******************************************************
                    If bGhost Then
                        If rkutils.DataGridView_ValueByColumnName(DataGridView_Ghosts, "MNAME", iGhosts_CTR).Trim.Length > 0 Then PrintLine_tLINE += " " + rkutils.DataGridView_ValueByColumnName(DataGridView_Ghosts, "MNAME", iGhosts_CTR).Trim
                        PrintLine_tLINE += " " + rkutils.DataGridView_ValueByColumnName(DataGridView_Ghosts, "LNAME", iGhosts_CTR).Trim
                        PrintLine_tLINE += " " + rkutils.DataGridView_ValueByColumnName(DataGridView_Ghosts, "", iGhosts_CTR).Trim + tDelimiter
                        '**********************************************************
                        '* Guarantor Address
                        PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Ghosts, "ADDRESS", iGhosts_CTR).Trim + tDelimiter
                        PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Ghosts, "ADDRESS2", iGhosts_CTR).Trim + tDelimiter
                        PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Ghosts, "CITY", iGhosts_CTR).Trim + " "
                        PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Ghosts, "STATE", iGhosts_CTR).Trim + " "
                        PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Ghosts, "ZIP", iGhosts_CTR).Trim.PadLeft(5, "0")
                        tSTR = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "ZIP4", iGhosts_CTR).Trim
                        If tSTR.Length = 4 Then
                            PrintLine_tLINE += " " + tSTR
                        End If
                    Else
                        If rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGMI", iLine).Trim.Length > 0 Then PrintLine_tLINE += " " + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGMI", iLine).Trim
                        PrintLine_tLINE += " " + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGLNM", iLine).Trim
                        PrintLine_tLINE += " " + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGSFX", iLine).Trim + tDelimiter
                        '**************************************************************************
                        '* Guarantor Address
                        PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGADD", iLine).Trim + tDelimiter
                        PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGAD2", iLine).Trim + tDelimiter
                        PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGCST", iLine).Trim + " "
                        PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGZIP", iLine).Trim.PadLeft(5, "0")
                        tSTR = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGZP2", iLine).Trim
                        If tSTR.Length = 4 Then
                            PrintLine_tLINE += " " + tSTR
                        End If
                    End If
                    PrintLine_tLINE += ":" + tDelimiter
                    '******************************************************************************
                    '* 2012-10-12 RFK: Specific Facility
                    tFacility = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAFACL", iLine).Trim
                    iFacilityRow = rkutils.DataGridView_Contains(DataGridView_Facilities, "FARFID", tFacility)
                    '******************************************************************************
                    '*                                                                   [Insert 1]
                    PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RACL#", iLine).Trim + tDelimiter                                            'Client             (INSERT 1)
                    '******************************************************************************
                    '* 2013-02-08 RFK: C.eLettersNameToUse                               [Insert 2]
                    Select Case rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "ELETTERSNAMETOUSE", Val(Label_ClientRow.Text))
                        Case "B"    'Both Client.FriendlyName + ROIDATA.FANAME (Facility)
                            PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "FRIENDLYNAME", Val(Label_ClientRow.Text)).Trim                     'Friendly Name      
                            PrintLine_tLINE += " " + rkutils.DataGridView_ValueByColumnName(DataGridView_Facilities, "FANAME", iFacilityRow).Trim + tDelimiter                  'Facility Name
                        Case "C"    'Client.FriendlyName
                            PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "FRIENDLYNAME", Val(Label_ClientRow.Text)).Trim + tDelimiter        'Friendly Name
                        Case "L"    'Client.LetterName
                            PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "LETTERNAME", Val(Label_ClientRow.Text)).Trim + tDelimiter          '* 2019-07-01 RFK: Letter Name
                        Case "F"    'ROIDATA.FACILP/HFNAME
                            PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Facilities, "HFNAME", iFacilityRow).Trim + tDelimiter                        'FACILP/Facility Name
                        Case Else
                            '* 2019-07-01 RFK: PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "FRIENDLYNAME", Val(Label_ClientRow.Text)).Trim + tDelimiter        'Friendly Name
                            PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "LETTERNAME", Val(Label_ClientRow.Text)).Trim + tDelimiter          '* 2019-07-01 RFK: Letter Name
                    End Select
                    '******************************************************************************
                    '*                                                                   [Insert 3]
                    PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RATOB", iLine).Trim + tDelimiter
                    '******************************************************************************
                    '*                                                                   [Insert 4]
                    Select Case tClient
                        Case "CNS"  '2016-04-19 RFK:
                            PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Facilities, "FACFID", iFacilityRow).Trim + tDelimiter
                        Case Else
                            PrintLine_tLINE += tFacility + tDelimiter
                    End Select
                    '******************************************************************************
                    '* Facility Name                                                     [Insert 5]
                    '* 2022-06-09 RFK: changed to ELETTERSFACILITYTOUSE
                    Select Case rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "ELETTERSFACILITYTOUSE", Val(Label_ClientRow.Text))
                        Case "R"
                            PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Facilities, "FASRCK", iFacilityRow).Trim + tDelimiter
                        Case Else
                            PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Facilities, "FANAME", iFacilityRow).Trim + tDelimiter
                    End Select
                    '******************************************************************************
                    '* Match Type                                                        [Insert 6]
                    PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAMTTP", iLine).Trim + tDelimiter                                        'Match Type         (INSERT 6)
                    '******************************************************************************
                    '* LOCX                                                              [Insert 7]
                    PrintLine_tLINE += tSysAccount + tDelimiter                                                                                                                 'LOCX               (INSERT 7)
                    '***********************************
                    '* 2014-08-26 RFK: Trim/Strip RAACCT
                    Select Case tClient
                        Case "RMC"
                            If STR_LEFT(rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAACCT", iLine).Trim, 1) = "0" Then
                                PrintLine_tLINE += Val(rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAACCT", iLine)).ToString.Trim + tDelimiter              'Account Number     (INSERT 8)
                            Else
                                PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAACCT", iLine).Trim + tDelimiter                            'Account Number     (INSERT 8)
                            End If
                        Case Else
                            PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAACCT", iLine).Trim + tDelimiter                                'Account Number     (INSERT 8)
                    End Select
                    '***********************************************
                    '* 2013-03-21 RFK:
                    tSuffix = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RASUFX", iLine).Trim                                                              'SUFIX              (INSERT 9)
                    PrintLine_tLINE += tSuffix + tDelimiter
                    '***********************************************
                    '* 2013-03-21 RFK:
                    tMedRec = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAMR#", iLine).Trim                                                               'MedRecord          (INSERT 10)
                    Select Case tClient
                        Case "STV"
                            PrintLine_tLINE += "1" + tMedRec.PadLeft(9, "0") + "A2892" + tDelimiter
                        Case Else
                            PrintLine_tLINE += tMedRec + tDelimiter
                    End Select
                    '***********************************************
                    '* Guarantor Name
                    If bGhost Then
                        PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Ghosts, "LNAME", iGhosts_CTR).Trim + tDelimiter                                  '(INSERT 11)
                    Else
                        PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGLNM", iLine).Trim + tDelimiter                                    '(INSERT 11)
                    End If
                    '**********************************************************
                    '* Guarantor Name   
                    '* 2013-10-09 RFK: Changed to First, Middle, Last
                    '* 2015-07-21 RFK: 
                    If ReadField(DataGridView_Select, "RAGFNM", iLine).Trim.Length = 0 Then
                        Select Case rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "ELETTERS", Val(Label_ClientRow.Text))
                            Case "t"
                                '**********************************************
                                '* 2015-07-21 RFK: 
                                PrintLine_tLINE += "TEST" + tDelimiter
                            Case Else
                                rkutils.DataGridView_SetValueByColumnName(DataGridView_Select, "ERRORCODE", iLine, "INVALID_GUARANTOR")
                                LettersPrinted("GUAR", tRALOCX, True)
                                If CheckBox_Update.Checked Then rkutils.COMMAND_STATUS(msSQLConnectionString, msSQLuser, tRALOCX, sStatusCodeBadGuarantorName, "", "", "")
                        End Select
                    Else
                        If bGhost Then
                            PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Ghosts, "FNAME", iGhosts_CTR).Trim + tDelimiter                                  '(INSERT 12)
                        Else
                            PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGFNM", iLine).Trim + tDelimiter                                    '(INSERT 12)
                        End If
                    End If
                    '**********************************************************
                    PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGMI", iLine).Trim.PadRight(2) + tDelimiter                                 ' (INSERT 13)
                    '*******************
                    PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGSFX", iLine).Trim + tDelimiter                                              ' (INSERT 14)
                    '* Guarantor Info       
                    PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGSS#", iLine).Trim + tDelimiter                                              ' (INSERT 15)
                    PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGBMO", iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGBDY", iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGBYR", iLine).Trim + tDelimiter
                    'PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView1, "RAGSEX", iLine).Trim + tDelimiter
                    'PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView1, "RAGMAR", iLine).Trim + tDelimiter
                    '******************************************************************
                    '* 2013-10-14 RFK: Changed to FORMAT the AC+PHONE to 12 digit phone
                    '* 2013-10-14 RFK: PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGHAC", iLine).Trim + tDelimiter                           ' (INSERT 17)
                    '* 2013-10-14 RFK: PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGHPH", iLine).Trim + tDelimiter                           ' (INSERT 18)
                    PrintLine_tLINE += "" + tDelimiter                                                                                                                                ' (INSERT 17)
                    '******************************************************************
                    tSTR = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGHAC", iLine).Trim + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAGHPH", iLine).Trim
                    If tSTR.Length = 10 And STR_LEFT(tSTR, 3) <> "999" Then
                        PrintLine_tLINE += rkutils.STR_format(tSTR, "PHONE") + tDelimiter                                                                                             ' (INSERT 18)
                    Else
                        PrintLine_tLINE += "" + tDelimiter                                                                                                                            ' (INSERT 18)
                    End If
                    '******************************************************************
                    '* Patient
                    If bGhost Then
                        PrintLine_tLINE += rkutils.ReadField(DataGridView_Ghosts, "FNAME", iGhosts_CTR).Trim + " " + rkutils.ReadField(DataGridView_Ghosts, "LNAME", iGhosts_CTR).Trim + tDelimiter ' (INSERT 12)
                    Else
                        PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAPNAM", iLine).Trim + tDelimiter                                           ' (INSERT 19)
                    End If
                    PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAPSS#", iLine).Trim + tDelimiter                                               ' (INSERT 20)
                    PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RABMON", iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RABDAY", iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RABYR", iLine).Trim + tDelimiter
                    PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAPHAC", iLine).Trim + tDelimiter                                                ' (INSERT 22)
                    PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAPHPH", iLine).Trim + tDelimiter                                                ' (INSERT 23)
                    'Financial
                    sCurrentBalance = rkutils.STR_format(rkutils.ReadFieldNoSpecialCharacters(DataGridView_Select, "RABALD", iLine, False), "#,##0.00")
                    PrintLine_tLINE += rkutils.STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RABALD", iLine).Trim, "#,##0.00") + tDelimiter                'Balance Due
                    PrintLine_tLINE += rkutils.STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAOBAL", iLine).Trim, "#,##0.00") + tDelimiter                'OriginalBalance
                    PrintLine_tLINE += rkutils.STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAOPD", iLine).Trim, "#,##0.00") + tDelimiter                 'OriginalPaid
                    PrintLine_tLINE += rkutils.STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAOADJ", iLine).Trim, "#,##0.00") + tDelimiter                'OriginalAdjusted
                    PrintLine_tLINE += rkutils.STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAAAMT", iLine).Trim, "#,##0.00") + tDelimiter                'AssignedAmount
                    PrintLine_tLINE += rkutils.STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RATOTP", iLine).Trim, "#,##0.00") + tDelimiter                'TotalPaid
                    PrintLine_tLINE += rkutils.STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RATOTA", iLine).Trim, "#,##0.00") + tDelimiter                'TotalAdjusted
                    '*****************************************
                    '* 2014-07-16 RFK:
                    PrintLine_tLINE += rkutils.STR_format(sSentBalance, "#,##0.00") + tDelimiter                                                                                        'SentBalance
                    If Val(sSentBalance) <> Val(rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAMBAL", iLine).Trim) Then
                        File.AppendAllText(dir_REPORTS + tClient + "_" + rkutils.STR_format("TODAY", "ccyymmdd") + "_MATCHEDBALANCE_ERROR.TXT", "MATCHBAL DOES NOT MATCH:" + tLOCX + vbTab + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAMBAL", iLine).Trim + vbTab + sSentBalance + vbTab)
                    End If
                    '*****************************************
                    '* Last Pay Date
                    PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RALPMO", iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RALPDY", iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RALPYR", iLine).Trim + tDelimiter
                    '* Last Pay Amount
                    PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "", iLine).Trim + tDelimiter                                                '                   (INSERT )
                    '* Payment Plan_Amount
                    PrintLine_tLINE += rkutils.STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RATERM", iLine).Trim, "#,##0.00") + tDelimiter              '                   (INSERT )
                    '* Payment Play Date
                    '* 2014-04-01 RFK: Fix For Invalid Date
                    sDate = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RATRMM", iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RATRMD", iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RATRMY", iLine).Trim
                    If IsDate(sDate) Then
                        PrintLine_tLINE += sDate + tDelimiter
                    Else
                        sDate = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RATRMM", iLine).Trim + "/28/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RATRMY", iLine).Trim
                        If IsDate(sDate) Then
                            PrintLine_tLINE += sDate + tDelimiter
                        Else
                            sDate = Str(Val(rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RATRMM", iLine)) + 1) + "/1/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RATRMY", iLine).Trim
                            If IsDate(sDate) Then
                                PrintLine_tLINE += sDate + tDelimiter
                            Else
                                MsgStatus("PrintLine failed/Invalid Payment Date/" + tSysAccount, True)
                                File.AppendAllText(FileNameRPT, "PrintLine failed/" + tSysAccount + vbCrLf)
                                PrintLine_tLINE += "" + tDelimiter
                            End If
                        End If
                    End If
                    '******************************************************************************
                    '* Fin Class Client
                    PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RACLCL", iLine).Trim + tDelimiter                            'FinClassClient
                    '******************************************************************************
                    '* Fin Class
                    PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RACFCL", iLine).Trim + tDelimiter                            'FinClass
                    '******************************************************************************
                    '* Admit Date
                    PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAADMM", iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAADMD", iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAADMY", iLine).Trim + tDelimiter
                    '******************************************************************************
                    '* DOS                                                              (Insert 39)
                    PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RADISM", iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RADISD", iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RADISY", iLine).Trim + tDelimiter
                    '******************************************************************************
                    '* DatePlaced                                                       (Insert 40)
                    PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAAMON", iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAADAY", iLine).Trim + "/" + rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAAYR", iLine).Trim + tDelimiter
                    '******************************************************************************
                    '* 2014-05-14 RFK: DrReferring                                      (INSERT 41)
                    PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAREDR", iLine).Trim + tDelimiter
                    '******************************************************************************
                    '* 2014-05-14 RFK: DrRendering                                      (INSERT 42)
                    PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAPRNM", iLine).Trim + tDelimiter
                    '******************************************************************************
                    '* 2022-05-09 RFK: Company / Agency                                 (INSERT 43)
                    PrintLine_tLINE += Letter_Types_Value(sLetterNumber, "COMPANY") + tDelimiter
                    '******************************************************************************
                    '* Multi Account Number / Associated Balance                        (INSERT 44)
                    PrintLine_tLINE += MultiAccounts(iLine, sLetterType) + tDelimiter
                    '******************************************************************************
                    '* CLient Phone NUmber                                              (INSERT 45)
                    PrintLine_tLINE += rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CALLBACKNUMBER", Val(Label_ClientRow.Text)).Trim + tDelimiter
                    '******************************************************************************
                    '* BadDebtNotification                                              (INSERT 46)
                    If Letter_Types_Value(sLetterNumber, "BadDebtNotification") = "Y" Then
                        PrintLine_tLINE += "Y" + tDelimiter
                        rkutils.DataGridView_SetValueByColumnName(DataGridView_Select, "SentBadDebtNotification", iLine, "Y")
                    Else
                        PrintLine_tLINE += tDelimiter
                    End If
                    '**********************************************************
                    '* 2013-03-26 RFK: (INSERT 47)
                    '* 2021-11-24 RFK:
                    Select Case Letter_Types_Value(sLetterNumber, "ChargeOffDate")
                        Case "Y"
                            '**************************************************
                            '* 2021-11-26 RFK: RATS8                (INSERT 47)
                            sDate = rkutils.STR_format(rkutils.ReadField(DataGridView_Select, "RATS8", iLine).Trim, "mm/dd/ccyy")
                            If IsDate(sDate) Then
                                PrintLine_tLINE += STR_format(sDate, "mm/dd/ccyy") + tDelimiter
                            Else
                                PrintLine_tLINE += "" + tDelimiter
                            End If
                        Case "y"
                            '**************************************************
                            '* 2021-11-26 RFK: Placement Date - 1   (INSERT 47)
                            sDate = rkutils.STR_format(rkutils.ReadField(DataGridView_Select, "PlacementDate1", iLine).Trim, "mm/dd/ccyy")
                            If IsDate(sDate) = False Then
                                'sDate = rkutils.ReadField(DataGridView_Select, "RAAMON", iLine).Trim + "/" + rkutils.ReadField(DataGridView_Select, "RAADAY", iLine).Trim + "/" + rkutils.ReadField(DataGridView_Select, "RAAYR", iLine).Trim
                            End If
                            If IsDate(sDate) Then
                                PrintLine_tLINE += STR_format(sDate, "mm/dd/ccyy") + tDelimiter
                            Else
                                PrintLine_tLINE += "" + tDelimiter
                            End If
                        Case Else
                            '**************************************************
                            '*                          (INSERT 47)
                            PrintLine_tLINE += "" + tDelimiter
                    End Select
                    '***********************************************
                    '* 2013-03-21 RFK: Charge Off Amount (INSERT 48)
                    PrintLine_tLINE += rkutils.STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAAAMT", iLine).Trim, "#,##0.00") + tDelimiter 'AssignedAmount
                    '**********************************************************
                    '* 2013-10-15 RFK: SCANLINE (INSERT 49)
                    Select Case tClient
                        Case "CNS"
                            '**************************************************
                            '* 2016-04-13 RFK: MOD 10 371
                            '* 2016-05-18 RFK: [MIGHT] change to MEDREC instead of Account
                            sScanLine = STR_TRIM(rkutils.ReadField(DataGridView_Select, "FACFID", iLine).Trim.PadRight(4, " "), 4)
                            sScanLine += " "
                            sScanLine += STR_TRIM(rkutils.ReadField(DataGridView_Select, "RAACCT", iLine).Trim.PadLeft(11, "0"), 11)
                            'sScanLine += STR_TRIM(rkutils.ReadField(DataGridView_Select, "RAMR#", iLine).Trim.PadLeft(11, "0"), 11)
                            sScanLine += " "
                            sScanLine += STR_TRIM(STR_format(sSentBalance, "000"), 12).PadLeft(12, "0")
                            '**********************************************************************************************
                            sScanLine += " " + Mod10_CheckDigit(sScanLine, "371", True, False)
                            PrintLine_tLINE += sScanLine + tDelimiter
                        Case "COP"
                            '* 2014-05-15 RFK: MOD 10 21
                            sScanLine = STR_TRIM(tMedRec.PadLeft(10, "0"), 10)           'Medical Record Number
                            '**********************************************************************************************
                            sScanLine += STR_TRIM(STR_format(sSentBalance, "000"), 12).PadLeft(12, "0")
                            '**********************************************************************************************
                            sScanLine += Mod10_CheckDigit(sScanLine, "21", True, False)
                            PrintLine_tLINE += sScanLine + tDelimiter
                        Case "GPY"
                            '* 2013-10-15 RFK: Medical Record #   1-20 [20]
                            '* 2013-10-15 RFK: SPACE
                            '* 2013-10-15 RFK: Group # 22-26 [5]
                            '* 2013-10-15 RFK: SPACE
                            '* 2013-10-15 RFK: Amount Due 28-37 [10]
                            '* 2013-10-15 RFK: SPACE
                            '* 2013-10-15 RFK: Check Digit: Position 39  [spaces not calculated in check digit]
                            sScanLine = STR_TRIM(tMedRec.PadLeft(20, "0"), 20)           'Medical Record Number
                            sScanLine += " "
                            '**********************************************************************************************\
                            '* 2013-10-15 RFK: GROUP # is stored in FAFREE in FACILP
                            sSQL = "SELECT FAFREE FROM ROIDATA.FACILP WHERE FARCL#='" + tClient + "' AND FARFID='" + tFacility + "'"
                            sSQLout = rkutils.SQL_READ_FIELD(DataGridView3, "DB2", "FAFREE", DB2SQLConnectionString, DB2SQLuser, sSQL)
                            sScanLine += STR_TRIM(sSQLout.PadLeft(5, "0"), 5)
                            sScanLine += " "
                            '**********************************************************************************************
                            sScanLine += STR_TRIM(STR_format(sSentBalance, "000"), 10).PadLeft(10, "0")
                            sScanLine += " "
                            '**********************************************************************************************
                            sScanLine += Mod10_CheckDigit(sScanLine, "7532", True, False)
                            PrintLine_tLINE += sScanLine + tDelimiter
                        Case "MAH"
                            '**************************************************************
                            '* 2014-02-01 RFK: F00012345678000019999000017999022713031913
                            '* 2014-02-01 RFK: Account: F00012345678  (12 digits in length)
                            '* 2014-02-01 RFK: Amt Due: $199.99 (9 digits in length)
                            '* 2014-02-01 RFK: Dscnt Amt: $179.99  (9 digits in length)
                            '* 2014-02-01 RFK: Stmt Date: 02/27/13 (6 digits in length)
                            '* 2014-02-01 RFK: Due Date: 03/19/13 (6 digits in length)
                            '* 2014-02-01 RFK: Check digit -Check Digit is calculated using the Mod 10 sum of digits method using a weight calculation of 3,5,7,9
                            '* 2014-02-01 RFK: Alphanumeric(Conversion)
                            '* 2014-02-01 RFK: A, J, S = 1
                            '* 2014-02-01 RFK: B, K, T = 2
                            '* 2014-02-01 RFK: C, L, U = 3
                            '* 2014-02-01 RFK: D, M, V = 4
                            '* 2014-02-01 RFK: E, N, W = 5
                            '* 2014-02-01 RFK: F, O, X = 6
                            '* 2014-02-01 RFK: G, P, Y = 7
                            '* 2014-02-01 RFK: H, Q, Z = 8
                            '* 2014-02-01 RFK: I, R = 9
                            tAccountNumber = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAACCT", iLine).Trim
                            sScanLine = STR_TRIM(tAccountNumber, 12)
                            sScanLine += STR_TRIM(STR_format(sCurrentBalance, "#").Replace(".", ""), 9).PadLeft(9, "0")
                            sScanLine += "000000000"
                            sScanLine += STR_format("TODAY", "mmddyy")
                            sScanLine += STR_format("TODAY", "mmddyy")
                            '**************************************************
                            sScanLine += Mod10_CheckDigit(sScanLine, "3579", True, False)
                            'MsgStatus(tAccountNumber + "=" + sScanLine, True)
                            PrintLine_tLINE += sScanLine + tDelimiter
                        Case Else
                            PrintLine_tLINE += tDelimiter
                    End Select
                    '**********************************************************
                    Select Case sLetterVendor
                        Case "ACCUDOC"
                            '**************************************************
                            '* 2015-08-20 RFK: BUCKET 0 to 30
                            PrintLine_tLINE += SentBalanceDollars30.ToString + tDelimiter
                            '**************************************************
                            '* 2015-08-20 RFK: BUCKET 31 to 60
                            PrintLine_tLINE += SentBalanceDollars60.ToString + tDelimiter
                            '**************************************************
                            '* 2015-08-20 RFK: BUCKET 31 to 60
                            PrintLine_tLINE += SentBalanceDollars90.ToString + tDelimiter
                            '**************************************************
                            '* 2015-08-20 RFK: BUCKET 91 to 120
                            PrintLine_tLINE += SentBalanceDollars120.ToString + tDelimiter
                            '**************************************************
                            '* 2015-08-20 RFK: BUCKET Greater than 120
                            PrintLine_tLINE += SentBalanceDollars121.ToString + tDelimiter
                            '**********************************************************************************************
                            '* 2014-07-14 RFK:
                            Select Case Letter_Types_Value(sLetterNumber, "BCAP")
                                Case "Y"
                                    '**********************************************************************************************
                                    '* Hospital Balance         (INSERT 50)
                                    If Val(CalculateCAPB("B", "=", "1", tRamLOCX)) = 0 And Val(CalculateCAPB("B", "<>", "1", tRamLOCX)) = 0 Then
                                        MsgBox("ERROR:Val(CalculateCAPB=" + tRamLOCX)
                                    End If
                                    PrintLine_tLINE += STR_format(CalculateCAPB("B", "=", "1", tRamLOCX), "#,##0.00") + tDelimiter
                                    '* Hospital Charges         (INSERT 51)
                                    PrintLine_tLINE += STR_format(CalculateCAPB("C", "=", "1", tRamLOCX), "#,##0.00") + tDelimiter
                                    '* Hospital Adjustments     (INSERT 52)
                                    PrintLine_tLINE += STR_format(CalculateCAPB("A", "=", "1", tRamLOCX), "#,##0.00") + tDelimiter
                                    '* Hospital Payments        (INSERT 53)
                                    PrintLine_tLINE += STR_format(CalculateCAPB("P", "=", "1", tRamLOCX), "#,##0.00") + tDelimiter
                                    '* Phys Balance             (INSERT 54)
                                    PrintLine_tLINE += STR_format(CalculateCAPB("B", "<>", "1", tRamLOCX), "#,##0.00") + tDelimiter
                                    '* Phys Charges             (INSERT 55)
                                    PrintLine_tLINE += STR_format(CalculateCAPB("C", "<>", "1", tRamLOCX), "#,##0.00") + tDelimiter
                                    '* Phys Adjustments         (INSERT 56)
                                    PrintLine_tLINE += STR_format(CalculateCAPB("A", "<>", "1", tRamLOCX), "#,##0.00") + tDelimiter
                                    '* Phys Payments            (INSERT 57)
                                    PrintLine_tLINE += STR_format(CalculateCAPB("P", "<>", "1", tRamLOCX), "#,##0.00") + tDelimiter
                                Case Else
                                    '*                          (INSERT 50)
                                    PrintLine_tLINE += tDelimiter
                                    '*                          (INSERT 51)
                                    PrintLine_tLINE += tDelimiter
                                    '*                          (INSERT 52)
                                    PrintLine_tLINE += tDelimiter
                                    '*                          (INSERT 53)
                                    PrintLine_tLINE += tDelimiter
                                    '*                          (INSERT 54)
                                    PrintLine_tLINE += tDelimiter
                                    '*                          (INSERT 55)
                                    PrintLine_tLINE += tDelimiter
                                    '*                          (INSERT 56)
                                    PrintLine_tLINE += tDelimiter
                                    '*                          (INSERT 57)
                                    PrintLine_tLINE += tDelimiter
                            End Select
                            '**************************************************
                            '*                                  (INSERT 58)
                            Select Case Letter_Types_Value(tLetterPrinted, "INSURANCEJOIN")
                                Case "Y"
                                    Dim iRowInsure As Integer = 0
                                    If DataGridView_Insure.RowCount > 0 Then
                                        iRowInsure = rkutils.DataGridView_Contains(DataGridView_Insure, "IRCIN#", "1")
                                        If iRowInsure >= 0 Then
                                            PrintLine_tLINE += ReadField(DataGridView_Insure, "IRCARR", iRowInsure) + tDelimiter
                                        Else
                                            PrintLine_tLINE += "" + tDelimiter
                                        End If
                                    Else
                                        PrintLine_tLINE += "" + tDelimiter
                                    End If
                                    '******************************************
                                    '*                                  (INSERT 59)
                                    If DataGridView_Insure.RowCount > 0 Then
                                        iRowInsure = rkutils.DataGridView_Contains(DataGridView_Insure, "IRCIN#", "2")
                                        If iRowInsure >= 0 Then
                                            PrintLine_tLINE += ReadField(DataGridView_Insure, "IRCARR", iRowInsure) + tDelimiter
                                        Else
                                            PrintLine_tLINE += "" + tDelimiter
                                        End If
                                    Else
                                        PrintLine_tLINE += "" + tDelimiter
                                    End If
                                    '******************************************
                                    '*                                  (INSERT 60)
                                    If DataGridView_Insure.RowCount > 0 Then
                                        iRowInsure = rkutils.DataGridView_Contains(DataGridView_Insure, "IRCIN#", "3")
                                        If iRowInsure >= 0 Then
                                            PrintLine_tLINE += ReadField(DataGridView_Insure, "IRCARR", iRowInsure) + tDelimiter
                                        Else
                                            PrintLine_tLINE += "" + tDelimiter
                                        End If
                                    Else
                                        PrintLine_tLINE += "" + tDelimiter
                                    End If
                                    '******************************************
                                    '*                                  (INSERT 61)
                                    PrintLine_tLINE += "" + tDelimiter
                                    '**************************************************
                                    '*                                  (INSERT 62)
                                    PrintLine_tLINE += "" + tDelimiter
                                    '**************************************************
                                    '*                                  (INSERT 63)
                                    PrintLine_tLINE += "" + tDelimiter
                                    '**************************************************
                                    '*                                  (INSERT 64)
                                    PrintLine_tLINE += "" + tDelimiter
                                    '**************************************************
                                    '*                                  (INSERT 65)
                                    PrintLine_tLINE += "" + tDelimiter
                                    '**************************************************
                                    '*                                  (INSERT 66)
                                    PrintLine_tLINE += "" + tDelimiter
                                    '**************************************************
                                    '*                                  (INSERT 67)
                                    PrintLine_tLINE += "" + tDelimiter
                                    '**************************************************
                                    '*                                  (INSERT 68)
                                    PrintLine_tLINE += "" + tDelimiter
                                    '******************************************
                                    '*                                  (INSERT 69)
                                    PrintLine_tLINE += "" + tDelimiter
                                    '******************************************
                                    '*                                  (INSERT 70)
                                    PrintLine_tLINE += "" + tDelimiter
                            End Select
                            '**************************************************
                        Case "DANTOM", "REVSPRING"
                            '**********************************************************************************************
                            '* (INSERT 50)
                            PrintLine_tLINE += tDelimiter
                            '**********************************************************************************************
                            '* 2014-07-14 RFK:
                            Select Case Letter_Types_Value(sLetterNumber, "BCAP")
                                Case "Y"
                                    '**********************************************************************************************
                                    '* Hospital Balance         (INSERT 50)
                                    'PrintLine_tLINE += "HospBal:"
                                    If Val(CalculateCAPB("B", "=", "1", tRamLOCX)) = 0 And Val(CalculateCAPB("B", "<>", "1", tRamLOCX)) = 0 Then
                                        MsgStatus("WARNING: CalculateCAPB:" + tRamLOCX, True)
                                    End If
                                    PrintLine_tLINE += STR_format(CalculateCAPB("B", "=", "1", tRamLOCX), "#,##0.00") + tDelimiter
                                    '* Hospital Charges         (INSERT 51)
                                    'PrintLine_tLINE += "HospChg:"
                                    PrintLine_tLINE += STR_format(CalculateCAPB("C", "=", "1", tRamLOCX), "#,##0.00") + tDelimiter
                                    '* Hospital Adjustments     (INSERT 52)
                                    'PrintLine_tLINE += "HospAdj:"
                                    PrintLine_tLINE += STR_format(CalculateCAPB("A", "=", "1", tRamLOCX), "#,##0.00") + tDelimiter
                                    '* Hospital Payments        (INSERT 53)
                                    'PrintLine_tLINE += "HospPay:"
                                    PrintLine_tLINE += STR_format(CalculateCAPB("P", "=", "1", tRamLOCX), "#,##0.00") + tDelimiter
                                    '* Phys Balance             (INSERT 54)
                                    'PrintLine_tLINE += "PhyBal:"
                                    PrintLine_tLINE += STR_format(CalculateCAPB("B", "<>", "1", tRamLOCX), "#,##0.00") + tDelimiter
                                    '* Phys Charges             (INSERT 55)
                                    'PrintLine_tLINE += "PhyChg:"
                                    PrintLine_tLINE += STR_format(CalculateCAPB("C", "<>", "1", tRamLOCX), "#,##0.00") + tDelimiter
                                    '* Phys Adjustments         (INSERT 56)
                                    'PrintLine_tLINE += "PhyAdj:"
                                    PrintLine_tLINE += STR_format(CalculateCAPB("A", "<>", "1", tRamLOCX), "#,##0.00") + tDelimiter
                                    '* Phys Payments            (INSERT 57)
                                    'PrintLine_tLINE += "PhyPay:"
                                    PrintLine_tLINE += STR_format(CalculateCAPB("P", "<>", "1", tRamLOCX), "#,##0.00") + tDelimiter
                                Case Else
                                    '*                          (INSERT 50)
                                    PrintLine_tLINE += tDelimiter
                                    '*                          (INSERT 51)
                                    PrintLine_tLINE += tDelimiter
                                    '*                          (INSERT 52)
                                    PrintLine_tLINE += tDelimiter
                                    '*                          (INSERT 53)
                                    PrintLine_tLINE += tDelimiter
                                    '*                          (INSERT 54)
                                    PrintLine_tLINE += tDelimiter
                                    '*                          (INSERT 55)
                                    PrintLine_tLINE += tDelimiter
                                    '*                          (INSERT 56)
                                    PrintLine_tLINE += tDelimiter
                                    '*                          (INSERT 57)
                                    PrintLine_tLINE += tDelimiter
                            End Select
                            '***********************************************
                            '* 2014-07-14 RFK: RETURN ADDRESS BY FACILITY (INSERT 58)
                            Select Case tClient
                                Case "RMC"
                                    tFacility = rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RAFACL", iLine).Trim
                                    iFacilityRow = rkutils.DataGridView_Contains(DataGridView_Facilities, "FARFID", tFacility)
                                    If iFacilityRow >= 0 Then
                                        PrintLine_tLINE += ReadField(DataGridView_Facilities, "FAMANM", iFacilityRow)
                                    End If
                                    '*************************************
                            End Select
                            PrintLine_tLINE += tDelimiter   '(INSERT 58) 
                            '***********************************************
                            '* 2014-07-14 RFK: RETURN ADDRESS BY FACILITY (INSERT 59)
                            PrintLine_tLINE += tDelimiter   '(INSERT 59) 
                            '***********************************************
                            '* 2014-07-14 RFK: RETURN ADDRESS BY FACILITY (INSERT 60)
                            PrintLine_tLINE += tDelimiter   '(INSERT 60) 
                            '***********************************************
                            '* 2014-07-14 RFK: RETURN ADDRESS BY FACILITY (INSERT 61)
                            '* 2015-08-20 RFK: BUCKET 0 to 30
                            PrintLine_tLINE += SentBalanceDollars30.ToString + tDelimiter   '(INSERT 61) 
                            '**************************************************
                            '* 2015-08-20 RFK: BUCKET 31 to 60
                            PrintLine_tLINE += SentBalanceDollars60.ToString + tDelimiter   '(INSERT 62) 
                            '**************************************************
                            '* 2015-08-20 RFK: BUCKET 31 to 60
                            PrintLine_tLINE += SentBalanceDollars90.ToString + tDelimiter   '(INSERT 63) 
                            '**************************************************
                            '* 2015-08-20 RFK: BUCKET 91 to 120
                            PrintLine_tLINE += SentBalanceDollars120.ToString + tDelimiter   '(INSERT 64) 
                            '**************************************************
                            '* 2015-08-20 RFK: BUCKET Greater than 120
                            PrintLine_tLINE += SentBalanceDollars121.ToString + tDelimiter   '(INSERT 65) 
                            '*                          (INSERT 66)
                            PrintLine_tLINE += tDelimiter
                            '*                          (INSERT 67)
                            PrintLine_tLINE += tDelimiter
                            '*                          (INSERT 68)
                            PrintLine_tLINE += tDelimiter
                            '*                          (INSERT 69)
                            PrintLine_tLINE += tDelimiter
                    End Select
                    '**********************************************************
                    '* 2021-07-12 RFK:
                    Select Case sLetterType
                        Case "24", "25"
                            '**************************************************
                            '* 2021-07-12 RFK: Insurance 
                            PrintLine_tLINE += rkutils.ReadField(DataGridView_Select, "IRCARR", iLine) + tDelimiter
                            '**************************************************
                            '* 2021-07-12 RFK: Charges FREVCP
                            sSQL = "SELECT SUM(FRCHRG) AS TSUM"
                            Select Case rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "ELETTERS", Val(Label_ClientRow.Text))
                                Case "t"
                                    sSQL += " FROM ROITEST.FREVCP"
                                Case Else
                                    sSQL += " FROM ROIDATA.FREVCP"
                            End Select
                            sSQL += " WHERE FRLOCX IN('" + MultiAccountsLIST(iLine).Replace(" ", "','") + "')"
                            MsgStatus(sSQL, False)
                            sCharges = SQL_READ_FIELD(DataGridView3, "DB2", "TSUM", DB2SQLConnectionString, DB2SQLuser, sSQL)
                            '**************************************************
                            PrintLine_tLINE += rkutils.STR_format(sCharges, "0.00") + tDelimiter
                            '**************************************************
                            '* 2021-07-12 RFK: Total Adjustments
                            sSQL = "SELECT SUM(RATOTA)+SUM(RAOADJ) AS TSUM"
                            Select Case rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "ELETTERS", Val(Label_ClientRow.Text))
                                Case "t"
                                    sSQL += " FROM ROITEST.RACCTP"
                                Case Else
                                    sSQL += " FROM ROIDATA.RACCTP"
                            End Select
                            sSQL += " WHERE RALOCX IN('" + MultiAccountsLIST(iLine).Replace(" ", "','") + "')"
                            sCharges = SQL_READ_FIELD(DataGridView3, "DB2", "TSUM", DB2SQLConnectionString, DB2SQLuser, sSQL)
                            '**************************************************
                            PrintLine_tLINE += rkutils.STR_format(sCharges, "0.00") + tDelimiter
                            '**************************************************
                            '* 2021-07-12 RFK: Total Payments
                            sSQL = "SELECT SUM(RATOTP)+SUM(RAOPD) AS TSUM"
                            Select Case rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "ELETTERS", Val(Label_ClientRow.Text))
                                Case "t"
                                    sSQL += " FROM ROITEST.RACCTP"
                                Case Else
                                    sSQL += " FROM ROIDATA.RACCTP"
                            End Select
                            sSQL += " WHERE RALOCX IN('" + MultiAccountsLIST(iLine).Replace(" ", "','") + "')"
                            sCharges = SQL_READ_FIELD(DataGridView3, "DB2", "TSUM", DB2SQLConnectionString, DB2SQLuser, sSQL)
                            '**************************************************
                            PrintLine_tLINE += rkutils.STR_format(sCharges, "0.00") + tDelimiter
                            '**************************************************
                    End Select
                    '**********************************************************
                    '* 2012-08-03 RFK: Matched Accounts
                    PrintLine_tLINE += "{" + vbCrLf
                    PrintLine_tLINE += MatchedPrint(iLine, sLetterVendor, sLetterNumber, sLetterType, tSysAccount, bGhost)
                    PrintLine_tLINE += "}" + vbCrLf
                    '**********************************************************
                    PrintLine_tLINE += vbCrLf 'Blank Line 
                    File.AppendAllText(sFileName, PrintLine_tLINE)
                    '**********************************************************
                    If bGhost = False Then Printed(iLine, tLetterPrinted, "PRINTED")    '* Set initial RECORD as PRINTED
                    '**********************************************************
                Case "APEX"
            End Select
        Catch ex As Exception
            MsgError("PrintLine", ex.ToString)
        End Try
    End Sub

    Protected Function AccountsWhere(ByVal sSQL As String, ByVal sLetterType As String, ByVal sClient As String, ByVal sRAMlocx As String, ByVal sRegF As String, ByVal sChargeOffDate As String) As String
        Try
            '******************************************************************
            '* 2021-04-02 RFK:
            Dim sSQLwhere As String = ""
            sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RACL#='" + sClient + "'")
            sSQLwhere += " AND A.RACLOS<>'C'"
            '******************************************************************
            Select Case sLetterType
                Case "MATCHED"
                    '**********************************************************
                    '* 2014-07-16 RFK: NO CREDIT BALANCES  (COP)
                    '* 2017-04-24 RFK: CHANGED TO CALL CLIENTS
                    sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RABALD > 0")
                    '**********************************************************
                    '* 2021-11-26 RFK: Reg F (sort by Charge Off Date)
                    Select Case sRegF
                        Case "Y"
                            sSQLwhere += " ORDER BY RAMLOCX"
                            sSQLwhere += ",RATS8"
                        Case "y"
                            sSQLwhere += " ORDER BY RAMLOCX"
                            '* 2021-01-11 RFK: sSQLwhere += ",DAYS(DIGITS(A.RAAMON)||'/'||DIGITS(A.RAADAY)||'/'||DIGITS(A.RAAYR))"    '* 2021-11-24 RFK: ChargeOffDate [Placement Date]
                        Case Else
                            sSQLwhere += " ORDER BY RAMLOCX"
                            sSQLwhere += ",A.RAADMY,A.RAADMM,A.RAADMD"   'Admit Date
                            sSQLwhere += ",A.RADISY,A.RADISM,A.RADISD"   'Discharge Date
                            sSQLwhere += ",A.RAAYR,A.RAAMON,A.RAADAY"    'Assignment Date
                            sSQLwhere += ",A.RAFACL"                     'Facility
                    End Select
                    '**********************************************************
                Case "RAMLOCX"
                    '**********************************************************
                    '* 2021-11-26 RFK: Reg F (sort by Charge Off Date)
                    sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RABALD>0")    '* 2017-04-26 RFK: no 0 balance accounts in the matched data
                    sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RAMLOCX='" + sRAMlocx + "'")
                    '**********************************************************
                    '* 2021-11-26 RFK: Reg F (sort by Charge Off Date)
                    Select Case sRegF
                        Case "Y"
                            sSQLwhere += " ORDER BY RAMLOCX"
                            sSQLwhere += ",RATS8"
                        Case "y"
                            '******************************************
                            '* 2021-11-26 RFK: Placement Days - 1
                            'MsgStatus(rkutils.STR_format(sChargeOffDate, "mm/dd/ccyy"), True)
                            sSQLwhere += rkutils.WhereAnd(sSQLwhere, "(days(DIGITS(A.RAAMON)||'/'||DIGITS(A.RAADAY)||'/'||DIGITS(A.RAAYR))-1) = days('" + rkutils.STR_format(sChargeOffDate, "mm/dd/ccyy") + "')")
                            sSQLwhere += " ORDER BY RAMLOCX"
                            sSQLwhere += ",DAYS(DIGITS(A.RAAMON)||'/'||DIGITS(A.RAADAY)||'/'||DIGITS(A.RAAYR))"    '* 2021-11-24 RFK: ChargeOffDate [Placement Date]
                        Case Else
                            sSQLwhere += " ORDER BY RAMLOCX"
                            sSQLwhere += ",A.RAADMY,A.RAADMM,A.RAADMD"   'Admit Date
                            sSQLwhere += ",A.RADISY,A.RADISM,A.RADISD"   'Discharge Date
                            sSQLwhere += ",A.RAAYR,A.RAAMON,A.RAADAY"    'Assignment Date
                            sSQLwhere += ",A.RAFACL"                     'Facility
                    End Select
                    '**********************************************************
                Case Else
                    Select Case ComboBox_MatchType.Text
                        Case "Active Clients", "Inactive Clients", "Testing Clients"
                            sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RAGADI<>'B'")
                            sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RABALD>0")
                            Select Case ComboBox_ClientActive.Text
                                Case "Testing"
                                    '******************************
                                    '* 2015-08-06 RFK:
                                    If rkutils.STR_TRIM(ComboBox_IgnoreDate.Text, 1) = "Y" Then
                                        '**********************************************************
                                        '* 2021-12-01 RFK: 
                                    Else
                                        sSQLwhere += rkutils.WhereAnd(sSQLwhere, "DIGITS(A.RANLYR)||DIGITS(A.RANLMO)||DIGITS(A.RANLDY)<=" + STR_format("TODAY", "ccyymmdd"))
                                    End If
                                    '******************************************
                                    '* 2021-11-26 RFK:
                                    If rkutils.STR_TRIM(ComboBox_TestAll.Text, 1) = "Y" Then
                                        '**************************************
                                        '* 2021-11-26 RFK: READ ALL accounts
                                    Else
                                        '**************************************
                                        '* 2021-11-26 RFK: RALNAC accounts
                                        sSQLwhere += " AND S.STLTRI='Y'"
                                        'sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RALNAC>0")
                                    End If
                                    '**************************************************************
                                    If TextBox_TCodes.Visible And TextBox_TCodes.Text.Trim.Length > 0 Then
                                        If sSQLwhere.Contains("A.RARSTA IN ('" + TextBox_TCodes.Text.Replace(" ", "','") + "')") = False Then
                                            sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RARSTA IN ('" + TextBox_TCodes.Text.Replace(" ", "','") + "')")
                                        End If
                                    End If
                                    '**************************************************************
                                Case Else
                                    '**************************************************************
                                    sSQLwhere += " AND S.STLTRI='Y'"
                                    'sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RALNAC>0")
                                    'sSQLwhere += rkutils.WhereAnd(sSQLwhere, "DIGITS(A.RANLYR)||DIGITS(A.RANLMO)||DIGITS(A.RANLDY)<=" + STR_format("TODAY", "ccyymmdd"))
                                    '**************************************************************
                            End Select
                            '**********************************************************************
                            '* 2021-11-26 RFK:
                            Select Case sRegF
                                Case "1"
                                    sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RALNAC=910")
                                    '* 2021-01-11 RFK: sSQLwhere += rkutils.WhereAnd(sSQLwhere, "(DAYS(DIGITS(A.RAAMON)||'/'||DIGITS(A.RAADAY)||'/'||DIGITS(A.RAAYR))-1) >= days(date('11/30/2021'))")    '* 2021-11-30 RFK: Placement Date must be >= 11/30/2021
                                    sSQLwhere += rkutils.WhereAnd(sSQLwhere, "DIGITS(A.RANLYR)||DIGITS(A.RANLMO)||DIGITS(A.RANLDY)<=" + STR_format("TODAY", "ccyymmdd"))
                                    sSQLwhere += rkutils.WhereAnd(sSQLwhere, "LEFT(A.RARSTA, 2)<>'PP'")
                                Case "2"
                                    sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RALNAC>0")
                                    sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RALNAC<>910")
                                    sSQLwhere += rkutils.WhereAnd(sSQLwhere, "LEFT(A.RARSTA, 2)<>'PP'")
                                    sSQLwhere += rkutils.WhereAnd(sSQLwhere, "DIGITS(A.RANLYR)||DIGITS(A.RANLMO)||DIGITS(A.RANLDY)<=" + STR_format("TODAY", "ccyymmdd"))
                                Case "3"
                                    '**************************************************************
                                    sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RARSTA IN('PPC','PPL')")
                                    sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RALNAC IN (950, 955)")
                                    '**************************************************************
                                    '* 10 days before payment is taken (TERM DATE) and (not within 5 days)
                                    sSQLwhere += rkutils.WhereAnd(sSQLwhere, "DIGITS(A.RATRMY)||DIGITS(A.RATRMM)||DIGITS(A.RATRMD)>=" + STR_format(rkutils.STR_DATE_PLUS(rkutils.STR_format("TODAY", "mm/dd/ccyy"), "+", 5), "ccyymmdd"))
                                    sSQLwhere += rkutils.WhereAnd(sSQLwhere, "DIGITS(A.RATRMY)||DIGITS(A.RATRMM)||DIGITS(A.RATRMD)<=" + STR_format(rkutils.STR_DATE_PLUS(rkutils.STR_format("TODAY", "mm/dd/ccyy"), "+", 10), "ccyymmdd"))
                                    '**************************************************************
                                    '* TERM AMOUNT
                                    sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RATERM>0")
                                    '**************************************************************
                                    '* STATEMENT SENT AT LEAST 21 DAYS AGO
                                    sSQLwhere += rkutils.WhereAnd(sSQLwhere, "RASDATE<=" + STR_format(rkutils.STR_DATE_PLUS(rkutils.STR_format("TODAY", "mm/dd/ccyy"), "-", 21), "ccyymmdd"))
                                    '**************************************************************
                                Case Else
                                    sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RALNAC>0")
                                    sSQLwhere += rkutils.WhereAnd(sSQLwhere, "DIGITS(A.RANLYR)||DIGITS(A.RANLMO)||DIGITS(A.RANLDY)<=" + STR_format("TODAY", "ccyymmdd"))
                            End Select
                            '**********************************************************************
                        Case "By Facility"
                            'sSQLwhere += " AND CF.LetterMatchType<>'*'"
                            sSQLwhere += " AND DIGITS(A.RANLYR)||DIGITS(A.RANLMO)||DIGITS(A.RANLDY)<=" + STR_format("TODAY", "ccyymmdd")
                        Case "Sent Today"
                            sSQLwhere += rkutils.WhereAnd(sSQLwhere, "((RAVERI=" + STR_format("TODAY", "mmddccyy") + ") OR (RAVERI=" + STR_format("YESTERDAY", "mmddccyy") + "))")
                            sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RABALD > 0")
                        Case Else
                            sSQLwhere += " AND DIGITS(A.RANLYR)||DIGITS(A.RANLMO)||DIGITS(A.RANLDY)<=" + STR_format("TODAY", "ccyymmdd")
                            '**********************************************************************
                    End Select
                    '******************************************************************************
                    '* 2013-12-27 RFK:
                    If TextBox_FinClass.Visible And TextBox_FinClass.Text.Trim.Length > 0 Then
                        sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RAPAYR IN ('" + TextBox_FinClass.Text.Replace(" ", "','") + "')")
                    End If
                    '******************************************************************************
                    '* 2016-08-19 RFK:
                    If TextBox_Facility.Visible And TextBox_Facility.Text.Trim.Length > 0 Then
                        Select Case ComboBox_Facility.Text
                            Case "="
                                sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RAFACL IN ('" + TextBox_Facility.Text.Replace(" ", "','") + "')")
                            Case "<>"
                                sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RAFACL NOT IN ('" + TextBox_Facility.Text.Replace(" ", "','") + "')")
                            Case "<="
                                sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RAFACL <= '" + TextBox_Facility.Text.Replace(" ", "','") + "'")
                        End Select
                    End If
                    '******************************************************************************
                    '* 2022-02-10 RFK:
                    If TextBox_DOP_DATE.Visible And TextBox_DOP_DATE.Text.Trim.Length > 0 Then
                        Select Case ComboBox_DOP.Text
                            Case "=", "<=", ">="
                                sSQLwhere += " AND DIGITS(A.RAAYR)||DIGITS(A.RAAMON)||DIGITS(A.RAADAY) " + ComboBox_DOP.Text + " " + STR_format(TextBox_DOP_DATE.Text, "ccyymmdd")
                        End Select
                    End If
                    '******************************************************************************
                    '* 2012-08-31 RFK:
                    '* 2012-10-15 RFK:
                    Select Case ComboBox_MatchType.Text
                        Case "Inactive Clients"
                            sSQLwhere += rkutils.WhereAnd(sSQLwhere, "DIGITS(A.RANLYR)||DIGITS(A.RANLMO)||DIGITS(A.RANLDY)=" + STR_format("TODAY", "ccyymmdd"))
                    End Select
                    '******************************************************************************
                    '* 2014-07-10 RFK:
                    If TextBox_LettersOnly.Visible = True And TextBox_LettersOnly.Text.Length > 0 Then
                        sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RALNAC IN ('" + TextBox_LettersOnly.Text.Replace(" ", "','") + "')")
                    End If
                    '******************************************************************************
                    '* 2016-08-19 RFK:
                    If TextBox_DateMM.Visible = True And TextBox_DateMM.Text.Length >= 1 Then
                        sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RANLMO IN('" + TextBox_DateMM.Text.Replace(" ", "','") + "')")
                    End If
                    If TextBox_DateDD.Visible = True And TextBox_DateDD.Text.Length >= 1 Then
                        sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RANLDY IN('" + TextBox_DateDD.Text.Replace(" ", "','") + "')")
                    End If
                    If TextBox_DateCCYY.Visible = True And TextBox_DateCCYY.Text.Length >= 4 Then
                        sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RANLYR IN('" + TextBox_DateCCYY.Text.Replace(" ", "','") + "')")
                    End If
                    '**************************************************************************
                    '* 2022-03-25 RFK:
                    If ComboBox_RATZ.Visible And ComboBox_RATZ.Text.Trim.Length > 0 And TextBox_RATZ.Text.Trim.Length > 0 Then
                        sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A." + ComboBox_RATZ.Text + " IN ('" + TextBox_RATZ.Text.Replace(" ", "','") + "')")
                    End If
                    '**********************************************************
                    '* 2013-03-06 RFK: Corrected for BAD NAMES
                    '* 2015-07-21 RFK: Only for not t clients
                    Select Case rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "ELETTERS", Val(Label_ClientRow.Text))
                        Case "t"
                            'No Matched Balance
                        Case Else
                            Select Case ComboBox_ClientActive.Text
                                Case "Testing"
                                    'NOTHING
                                Case Else
                                    sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RAMBAL >= D.MINLETTERBAL")
                                    'If TextBox_MinBalance.Text.Length > 0 And Val(TextBox_MinBalance.Text) > 0 Then
                                    '    'sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RAMBAL >=" + TextBox_MinBalance.Text + "")
                                    '    '* 2021-04-05 RFK: moved to ROIDATA.DILIGENCE
                                    'Else
                                    '    sSQLwhere += rkutils.WhereAnd(sSQLwhere, "A.RAMBAL>0")
                                    'End If
                            End Select
                    End Select
                    '**********************************************************
                    '* 2013-01-21 RFK:
                    If TextBox_MaxBalance.Text.Length > 0 And Val(TextBox_MaxBalance.Text) > 0 Then sSQLwhere += rkutils.WhereAnd(sSQLwhere, "RAMBAL <=" + TextBox_MaxBalance.Text + "")
                    '**********************************************************
                    '* 2018-01-25 RFK:
                    If TextBox_Field.Text.Trim.Length > 0 And TextBox_FieldValue.Text.Trim.Length > 0 Then sSQLwhere += rkutils.WhereAnd(sSQLwhere, TextBox_Field.Text + "='" + TextBox_FieldValue.Text + "'")
                    '**********************************************************
                    '* 2013-01-07 RFK: sSQLwhere += " ORDER BY CAST(A.RALOCX AS INTEGER)"
                    '* 2016-12-28 RFK: sSQLwhere += " ORDER BY RAMLOCX"
                    '* 2021-11-26 RFK: Reg F (sort by Charge Off Date)
                    Select Case sRegF
                        Case "Y"
                            sSQLwhere += " ORDER BY A.RAGLNM, RAMLOCX, RATS8"
                        Case "y"
                            sSQLwhere += " ORDER BY A.RAGLNM,RAMLOCX"
                            sSQLwhere += ",DAYS(DIGITS(A.RAAMON)||'/'||DIGITS(A.RAADAY)||'/'||DIGITS(A.RAAYR))"    '* 2021-11-24 RFK: ChargeOffDate [Placement Date]
                        Case Else
                            sSQLwhere += " ORDER BY A.RAGLNM, RAMLOCX"
                    End Select
                    '**********************************************************
                    If Val(TextBox_MaxAccounts.Text) > 0 Then
                        'Only Select MAX RECORDS
                        sSQLwhere += " FETCH FIRST " + TextBox_MaxAccounts.Text + " ROWS ONLY"
                    End If
                    '**********************************************************
            End Select
            'MsgStatus(sLetterType + " " + sSQLwhere, CheckBox_DEBUG.Checked)
            'MsgStatus(sLetterType + " " + sSQLwhere, True)
            Return sSQLwhere
            '******************************************************************
        Catch ex As Exception
            MsgError("AccountsWhere", ex.ToString)
        End Try
        Return ""
    End Function

    Private Sub DataGridView_Clients_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles DataGridView_Clients.MouseClick
        '******************************************************************************************
        '* RFK:
        Try
            Label_ClientRow.Text = DataGridView_Clients.CurrentCellAddress.Y.ToString
            AccountInitSettings()
        Catch ex As Exception
            MsgError("DataGridView_Clients_MouseClick", ex.ToString)
        End Try
    End Sub

    Private Sub DataGridView_Select_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView_Select.CellClick
        Try
            '**************************************************************************************
            Label_AccountRow.Text = Trim(Str(DataGridView_Select.CurrentCellAddress.Y))
            'MsgStatus(Label_AccountRow.Text, True)
            ReadLine(Val(Label_AccountRow.Text))
            '**************************************************************************************
            If swReadAllMatched = False Then
                AnnuityOne_AccountsLoad(rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTNAME", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "TOB", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "FACILITY", Val(Label_ClientRow.Text)), rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "CLIENTMATCHTYPE", Val(Label_ClientRow.Text)), "RAMLOCX", gtLetterType, tRamLOCX, Letter_Types_Value(rkutils.DataGridView_ValueByColumnName(DataGridView_Select, "RALNAC", Val(Label_AccountRow.Text)), "ChargeOffDate"))
            End If
            '**************************************************************************************
            '* 2022-06-20 RFK:
            If Val(rkutils.DataGridView_ColumnByName(DataGridView_Select, "RALOCX")) = DataGridView_Select.CurrentCellAddress.X Then
                'MsgStatus(DataGridView_Select.CurrentCellAddress.X.ToString + "][" + rkutils.ReadField(DataGridView_Select, "RALOCX", Val(Label_AccountRow.Text)) + "]", True)
                TextBox_TEST.Text = rkutils.ReadField(DataGridView_Select, "RALOCX", Val(Label_AccountRow.Text))
            End If
        Catch ex As Exception
            MsgError("DataGridView_Select_CellClick", ex.ToString)
        End Try
    End Sub

    Private Sub AnnuityOne_AccountsLoad(ByVal tClient As String, ByVal tTOB As String, ByVal tFacility As String, ByVal tClientType As String, ByVal sMatchType As String, ByVal sLetterType As String, ByVal sRamLocx As String, ByVal sRegF As String)
        Try
            '**************************************************************************************
            sSQL = "SELECT"
            '**************************************************************************************
            '* 2021-07-15 RFK:
            Select Case sLetterType
                Case "10"
                    Select Case sMatchType
                        Case "RAMLOCX"
                            Exit Sub
                    End Select
                Case "24", "25"
                    sSQL += " DISTINCT"
                Case Else
            End Select
            '**************************************************************************************
            sSQL += " A.RAMTTP,A.RACL#"
            sSQL += ",A.RAFACL"
            sSQL += ",A.RAACCT,A.RASUFX"
            sSQL += ",A.RALOCX,A.RAMLOCX"
            sSQL += ",A.RALCAC,A.RALNAC"
            '**************************************************************************************
            '* 2021-11-26 RFK:
            Select Case sRegF
                Case "3"
                    '******************************************************************************
                    '* TERM AMOUNT / DATES
                    sSQL += ",A.RATERM,A.RATRMM,A.RATRMD,A.RATRMY"
                Case Else
                    '******************************************************************************
            End Select
            '**************************************************************************************
            sSQL += ",A.RAGLNM,A.RAGFNM,A.RAGMI,A.RAGSFX"
            sSQL += ",A.RANLMO,A.RANLDY,A.RANLYR"
            sSQL += ",A.RARSTA,S.STLTRI"
            sSQL += ",A.RAVERI,A.RASDATE,A.RALETC"
            sSQL += ",A.RAMR#"
            sSQL += ",A.RAGADD,A.RAGAD2,A.RAGCST,RIGHT(TRIM(A.RAGCST),2) AS RAGSTATE,A.RAGZIP,A.RAGADI"
            sSQL += ",A.RAGSS#,A.RAGBMO,A.RAGBDY,A.RAGBYR"
            sSQL += ",A.RAGHAC,A.RAGHPH"
            sSQL += ",A.RAPNAM,A.RAPADD,A.RAPAD2,A.RAPCST,A.RAPZIP"
            sSQL += ",A.RAPSS#,A.RABMON,A.RABDAY,A.RABYR"
            sSQL += ",A.RASEX,A.RAPMAR,A.RAPHAC,A.RAPHPH"
            sSQL += ",A.RABALD,A.RAMBAL,A.RAOBAL,A.RAOPD,A.RAOADJ,A.RAAAMT,A.RATOTP,A.RATOTA"
            '**************************************************************************************
            '* 2021-11-26 RFK:
            Select Case sRegF
                Case "3"
                    '******************************************************************************
                Case Else
                    '******************************************************************************
                    '* TERM AMOUNT / DATES
                    sSQL += ",A.RATERM,A.RATRMM,A.RATRMD,A.RATRMY"
            End Select
            '**************************************************************************************
            sSQL += ",A.RATOB"
            sSQL += ",A.RACLCL,A.RACFCL"
            sSQL += ",A.RAADMM,A.RAADMD,A.RAADMY"
            sSQL += ",A.RADISM,A.RADISD,A.RADISY"
            sSQL += ",A.RAAMON,A.RAADAY,A.RAAYR"
            sSQL += ",DATE(DIGITS(A.RAAMON)||'/'||DIGITS(A.RAADAY)||'/'||DIGITS(A.RAAYR)) AS PlacementDate"             '* 2021-11-26 RFK: [Placement Date]
            sSQL += ",A.RALPMO,A.RALPDY,A.RALPYR"                                                                       '* 
            sSQL += ",A.RAPAYR"                                                                                         '* 
            sSQL += ",A.RARDRN, A.RAPRNM"                                                                               '* 
            sSQL += ",A.RAFR30"                                                                                         '* 2016-01-27 RFK: SelfPayDate
            sSQL += ",A.RACLOS"                                                                                         '* 
            sSQL += ",'' AS SentBalance"                                                                                '* Calculated Balance of selected / sent accounts
            sSQL += ",'' AS SentBadDebtNotification"                                                                    '* 1st Bad Debt Validation
            sSQL += ",A.RALCAD"                                                                                         '* 1st Bad Debt Validation Date
            sSQL += ",A.RATS8"                                                                                          '* 2021-11-24 RFK: ChargeOffDate [RATS8]
            sSQL += ",(DAYS(DIGITS(A.RAAMON)||'/'||DIGITS(A.RAADAY)||'/'||DIGITS(A.RAAYR))-1) AS PlacementDays1"        '* 2021-11-24 RFK: ChargeOffDate [DAYS Placement Date - 1]
            sSQL += ",DATE(DAYS(DIGITS(A.RAAMON)||'/'||DIGITS(A.RAADAY)||'/'||DIGITS(A.RAAYR))-1) AS PlacementDate1"    '* 2021-11-24 RFK: ChargeOffDate [Placement Date - 1]
            sSQL += ",A.RAMLOCX AS RAML"
            sSQL += ",D.MINLETTERBAL"                                                                                   '* 2021-04-05 RFK:
            '******************************************************************
            '* 2014-06-27 RFK:
            Select Case rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "ELETTERSLOCKBOX", Val(Label_ClientRow.Text))
                Case "F"
                    sSQL += ",F.FACFID,F.FANAME,F.FAMANM,F.FAMADD,F.FAMAD2,F.FAMACS,F.FAMAZP,F.FAMAZ2"
            End Select
            '******************************************************************
            '* 2014-08-05 RFK:
            '* 2012-11-06 RFK: Matched(All) Accounts 
            Select Case sMatchType
                Case "MATCHED"
                    'Nothing
                Case "RAMLOCX"
                    'Nothing
                Case Else
                    'sSQL += ",I.IRCARR" '2013-04-04 RFK:
                    sSQL += ",'' AS LETTERPRINTED"
                    sSQL += ",'' AS LETTERPRINTEDDATE"
                    sSQL += ",'' AS LETTERACTUAL"
                    sSQL += ",'' AS LETTERNEXT"
                    sSQL += ",'' AS LETTERNEXTDATE"
                    sSQL += ",'READY' AS ERRORCODE"
            End Select
            '******************************************************************
            '* 2014-08-04 RFK:
            sSQL += ",P.LETTERS AS StateAllow"
            '******************************************************************
            '* 2021-07-12 RFK:
            Select Case sLetterType
                Case "24", "25"
                    sSQL += ",I.IRCARR"
                Case Else
                    'MsgStatus(sLetterType, True)
            End Select
            '******************************************************************
            Select Case tClient
                Case "WPL"
            End Select
            '******************************************************************
            '* 2018-01-25 RFK:
            If TextBox_Field.Text.Trim.Length > 0 And TextBox_FieldValue.Text.Trim.Length > 0 And sSQL.Contains(TextBox_Field.Text) = False Then sSQL += "," + TextBox_Field.Text
            '******************************************************************
            '* 2013-01-21 RFK:
            Select Case rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "ELETTERS", Val(Label_ClientRow.Text))
                Case "t"
                    sSQL += " FROM ROITEST.RACCTP A"
                    '**********************************************************
                    '* 2014-06-27 RFK:
                    Select Case rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "ELETTERSLOCKBOX", Val(Label_ClientRow.Text))
                        Case "F"
                            sSQL += " LEFT JOIN ROITEST.FACILP F ON (A.RACL#=F.FARCL# AND A.RAFACL=F.FARFID)"
                    End Select
                Case Else
                    sSQL += " FROM ROIDATA.RACCTP A"
                    '**********************************************************
                    '* 2014-06-27 RFK:
                    Select Case rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "ELETTERSLOCKBOX", Val(Label_ClientRow.Text))
                        Case "F"
                            sSQL += " LEFT JOIN ROIDATA.FACILP F ON (A.RACL#=F.FARCL# AND A.RAFACL=F.FARFID)"
                    End Select
            End Select
            '******************************************************************
            '* 2021-08-20 RFK:
            sSQL += " LEFT JOIN ROIDATA.HCLNTP C ON (A.RACL#=C.HCCL#)"
            '******************************************************************
            '* 2013-01-21 RFK:
            sSQL += " LEFT JOIN ROIDATA.STATP S ON (A.RARSTA=S.STSTAT AND A.RAMTTP=S.STMTTP)"
            '******************************************************************
            '* 2014-08-04 RFK:
            '* 2021-08-20 RFK: P.SGROUP=C.HCFR30
            sSQL += " LEFT JOIN ROIDATA.STATEBLOCKING P ON ((RIGHT(TRIM(A.RAGCST),2)=P.POSTALCODE) AND (P.SGROUP=C.HCFR30))"
            '******************************************************************
            '* 2021-04-05 RFK:
            sSQL += " LEFT JOIN ROIDATA.DILIGENCE D ON D.DICL# = A.RACL# AND A.RAMBAL >= D.BAL_GTE AND A.RAMBAL <= D.BAL_LTE"
            '******************************************************************
            '* 2019-03-19 RFK: INSURANCE JOIN IRCPAY
            '* 2021-07-12 RFK:
            Select Case sLetterType
                Case "24", "25"
                    sSQL += JoinIRSURP()
            End Select
            '******************************************************************
            '* 2016-11-04 RFK:
            Select Case Letter_Types_Value(tLetterPrinted, "INSURANCEJOIN")
                Case "Y"
                    Select Case sMatchType
                        Case "RAMLOCX"
                            Dim sSQLinsure As String = String.Format(
                            "SELECT IRCL#, IRTOB, IRCPAY, IRMTTP, IRGLN5, IRGFN5, IRGAD5, IRGZIP, IRGSS#, IRIN#, IRCIN#, IRCARR, IRATTN, " +
                            "IRAREA, IRPHON, IREXTN, IRPOL#, IRGRP#, IRGRPN, IRRELC, IRPAYR, IRADR1, IRADR2, IRCITY, IRSTAT, " +
                            "IRZIP, IRZIPE, IRPOLH, IRHADD, IRHAD2, IRHCTY, IRHST, IRHZIP, IRHBMO, IRHBDY, IRHBYR, IRHPOE, IRHAC, " +
                            "IRHPH, IREFMO, IREFDY, IREFYR, IRETMO, IRETDY, IRETYR, IRHSS#, IRFR30, ISLOCX, ISREC, ISACTV, ISPAYT, ISCLM# " +
                            "FROM ROIDATA.IRSURP LEFT JOIN ROIDATA.ISTATP ON " +
                            "ISLOCX = {0} AND ISREC = {1} WHERE IRCL# = '{2}' AND IRTOB = {3} AND IRMTTP = '{4}' AND " +
                            "IRGLN5 = '{5}' AND IRGFN5 = '{6}' AND IRGAD5 = '{7}' AND IRGZIP = {8} AND IRGSS# = {9} AND IRCIN# = {1}",
                            ReadField(DataGridView_Select, "RALOCX", Val(Label_AccountRow.Text)),
                            "1",
                            ReadField(DataGridView_Select, "RACL#", Val(Label_AccountRow.Text)),
                            ReadField(DataGridView_Select, "RATOB", Val(Label_AccountRow.Text)),
                            ReadField(DataGridView_Select, "RAMTTP", Val(Label_AccountRow.Text)),
                            ReadField(DataGridView_Select, "RAGLNM", Val(Label_AccountRow.Text)).Trim.PadRight(5).Substring(0, 5).Replace("'", ""),
                            ReadField(DataGridView_Select, "RAGFNM", Val(Label_AccountRow.Text)).Trim.PadRight(5).Substring(0, 5).Replace("'", ""),
                            ReadField(DataGridView_Select, "RAGADD", Val(Label_AccountRow.Text)).PadRight(5).Substring(0, 5).Replace("'", ""),
                            ReadField(DataGridView_Select, "RAGZIP", Val(Label_AccountRow.Text)).PadLeft(5),
                            ReadField(DataGridView_Select, "RAGSS#", Val(Label_AccountRow.Text)))
                            rkutils.SQL_READ_DATAGRID(DataGridView_Insure, "DB2", "*", DB2SQLConnectionString, DB2SQLuser, sSQLinsure)
                            DataGridView_Insure.Visible = False
                    End Select
            End Select
            '******************************************************************
            '* 2014-08-05 RFK:
            MsgStatus("LOAD-" + sMatchType + "]RamLOCX:" + sRamLocx + "]" + Label_AccountRow.Text + "]", False)
            '**********************************************************
            '* 2021-04-02 RFK: WHERE subroutine
            '* 2021-11-26 RFK: Reg F (sort by Charge Off Date)
            Select Case sRegF
                Case "Y"
                    sSQL += AccountsWhere(sSQL, sMatchType, tClient, sRamLocx, sRegF, ReadField(DataGridView_Select, "RATS8", Val(Label_AccountRow.Text)))
                Case "y"
                    sSQL += AccountsWhere(sSQL, sMatchType, tClient, sRamLocx, sRegF, ReadField(DataGridView_Select, "PlacementDate1", Val(Label_AccountRow.Text)))
                Case Else
                    sSQL += AccountsWhere(sSQL, sMatchType, tClient, sRamLocx, sRegF, "")
            End Select
            Select Case sMatchType
                Case "MATCHED"
                    '**********************************************************
                    MsgStatus(sSQL, False)
                    If rkutils.SQL_READ_DATAGRID(DataGridView_Multi, "DB2", "*", DB2SQLConnectionString, DB2SQLuser, sSQL) Then
                        DataGridView_Multi.Visible = True
                        'Label_MultiRamLocx.Text = sRamLocx
                    Else
                        DataGridView_Multi.Visible = False
                    End If
                    Label_Total.Text = DataGridView_Multi.RowCount.ToString
                Case "RAMLOCX"
                    '**********************************************************
                    MsgStatus(sSQL, CheckBox_DEBUG.Checked)
                    '**********************************************************
                    If rkutils.SQL_READ_DATAGRID(DataGridView_Multi, "DB2", "*", DB2SQLConnectionString, DB2SQLuser, sSQL) Then
                        DataGridView_Multi.Visible = True
                        'Label_MultiRamLocx.Text = sRamLocx
                    Else
                        DataGridView_Multi.Visible = False
                    End If
                    Label_Total.Text = DataGridView_Multi.RowCount.ToString
                Case Else
                    '**********************************************************
                    Select Case sSITE
                        Case "AnnuityOne"
                            '**************************************************
                            '* 2019-03-14 RFK:
                            MsgStatus("Selecting Accounts", True)
                            MsgStatus(sSQL, False)
                            If swDTable Then
                                If SQL_READ_DATATABLE(dTable_Select, "DB2", "*", DB2SQLConnectionString, DB2SQLuser, sSQL) Then
                                    MsgStatus("Selected:" + dTable_Select.Rows.Count.ToString, True)
                                Else
                                    MsgStatus("DID NOT SELECT ANYTHING:" + sSQL, True)
                                End If
                            Else
                                If SQL_READ_DATAGRID(DataGridView_Select, "DB2", "*", DB2SQLConnectionString, DB2SQLuser, sSQL) Then
                                    DataGridView_Select.Visible = True
                                Else
                                    DataGridView_Select.Visible = False
                                End If
                            End If
                        Case Else
                            If SQL_READ_DATAGRID(DataGridView_Select, "MSSQL", "*", msSQLConnectionString, msSQLuser, sSQL) Then
                                DataGridView_Select.Visible = True
                            Else
                                DataGridView_Select.Visible = False
                            End If
                    End Select
                    '**********************************************************
                    '* 2019-03-14 RFK:
                    If swDTable Then
                        Label_NumberAccounts.Text = dTable_Select.Rows.Count - 1.ToString
                        Label_AccountsRemaining.Text = Label_NumberAccounts.Text
                    Else
                        Label_NumberAccounts.Text = DataGridView_Select.Rows.Count - 1.ToString
                        Label_AccountsRemaining.Text = Label_NumberAccounts.Text
                    End If
                    MsgStatus("AccountsLoaded [" + tClient + "](" + sMatchType + ")=" + Label_NumberAccounts.Text, True)
                    '**********************************************************
            End Select
        Catch ex As Exception
            MsgError("AccountsLoad", ex.ToString)
        End Try
    End Sub

End Class
