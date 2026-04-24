VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form reportfrm30 
   Caption         =   "”Ã· «·„—«Ã⁄…"
   ClientHeight    =   1725
   ClientLeft      =   1065
   ClientTop       =   1875
   ClientWidth     =   6630
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   1725
   ScaleWidth      =   6630
   Begin VB.CommandButton cmdExit 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   90
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   1035
      Width           =   1500
   End
   Begin VB.CommandButton cmdClear 
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1620
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "„”Õ «·ﬂ·"
      Top             =   1035
      Width           =   1500
   End
   Begin VB.CommandButton CmdApply 
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3150
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "⁄—÷ «·»Ì«‰« "
      Top             =   1035
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   1890
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   45
      Width           =   4605
      Begin VB.TextBox xCode2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Tag             =   "D"
         Top             =   540
         Width           =   1725
      End
      Begin VB.TextBox xcode1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Tag             =   "D"
         Top             =   180
         Width           =   1725
      End
      Begin VB.Label Label2 
         Caption         =   " ≈·Ì —ﬁ„"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   540
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "„‰ —ﬁ„"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   225
         Width           =   825
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   1485
      Top             =   -90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Threed.SSCommand cmdPdf 
      Cancel          =   -1  'True
      Height          =   555
      Left            =   4680
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1035
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   979
      _Version        =   196610
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "report30.frx":0000
      Caption         =   "Pdf ÿ»«⁄…"
      ButtonStyle     =   1
      PictureAlignment=   10
      BevelWidth      =   0
      PictureDisabledFrames=   1
      PictureDisabled =   "report30.frx":25CB
   End
End
Attribute VB_Name = "reportfrm30"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim oSearchYear As New Search_empty, oSearchJob As New Search, oSearchComp As New Search
Private Sub cmdApply_Click()
doprint
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Function doprint(Optional bPdf As Boolean = False)
Dim temptable As New ADODB.Recordset, sourcetable As New ADODB.Recordset
Dim aHeader(11)

Dim aPrm As Variant

If ValidNum(xcode1.text) Then
    aPrm = AddFlag(aPrm, "CODE1", xcode1.text)
    aHeader(1) = IIf(ValidNum(xCode2.text), BetweenString(xcode1.text, xCode2.text, "„‰ —ﬁ„ ⁄÷ÊÌ… : ", "Õ Ì —ﬁ„ ⁄÷ÊÌ… : "), "—ﬁ„ ⁄÷ÊÌ… :" & xcode1.text)
End If

If ValidNum(xCode2.text) Then
    aPrm = AddFlag(aPrm, "CODE2", xCode2.text)
    aHeader(1) = BetweenString(xcode1.text, xCode2.text, "„‰ —ﬁ„ ⁄÷ÊÌ… : ", "Õ Ì —ﬁ„ ⁄÷ÊÌ… : ")
End If


contemp.Execute "delete * from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adcmtext

Set sourcetable = cmd("[dbo].[sp_mem_check]", con, adStoredProc, aPrm).Execute

With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!val1 = sourcetable!code
    temptable!str1 = ArbString(sourcetable!code)
    temptable!str2 = sourcetable!Desca
    
    temptable!str3 = sourcetable!wife_name1
    temptable!str4 = sourcetable!wife_name2
    temptable!str5 = sourcetable!wife_name3
    temptable!str6 = sourcetable!wife_name4
    
    temptable!str7 = TurnValue(ArbString(myFormat_p(sourcetable!DATE_BIRTH)))
    temptable!str8 = TurnValue(ArbString(myFormat_p(sourcetable!wife_date_birth1)))
    temptable!str9 = TurnValue(ArbString(myFormat_p(sourcetable!wife_date_birth2)))
    temptable!str10 = TurnValue(ArbString(myFormat_p(sourcetable!wife_date_birth3)))
    temptable!str11 = TurnValue(ArbString(myFormat_p(sourcetable!wife_date_birth4)))
    
    temptable!str12 = TurnValue(ArbString(myFormat_p(sourcetable!date_begin)))
    temptable!str13 = TurnValue(ArbString(myFormat_p(sourcetable!wife_date_begin1)))
    temptable!str14 = TurnValue(ArbString(myFormat_p(sourcetable!wife_date_begin2)))
    temptable!str15 = TurnValue(ArbString(myFormat_p(sourcetable!wife_date_begin3)))
    temptable!str16 = TurnValue(ArbString(myFormat_p(sourcetable!wife_date_begin4)))
    
    temptable!str17 = TurnValue(ArbString(sourcetable!JOB_desca))
    temptable!str18 = TurnValue(ArbString(sourcetable!Phone))
    If IsNull(sourcetable!Phone) Then
        temptable!str18 = TurnValue(ArbString(sourcetable!Mobil))
    Else
        temptable!str46 = TurnValue(ArbString(sourcetable!Mobil))
    End If
    temptable!str19 = TurnValue(ArbString(sourcetable!Address))
    
    temptable!str20 = TurnValue(sourcetable!son_name1)
    temptable!str21 = TurnValue(sourcetable!son_name2)
    temptable!str22 = TurnValue(sourcetable!son_name3)
    temptable!str23 = TurnValue(sourcetable!son_name4)
    temptable!str24 = TurnValue(sourcetable!son_name5)
        
    temptable!str25 = TurnValue(ArbString(myFormat_p(sourcetable!SON_DATE_BIRTH1)))
    temptable!str26 = TurnValue(ArbString(myFormat_p(sourcetable!son_date_birth2)))
    temptable!str27 = TurnValue(ArbString(myFormat_p(sourcetable!son_date_birth3)))
    temptable!str28 = TurnValue(ArbString(myFormat_p(sourcetable!son_date_birth4)))
    temptable!str29 = TurnValue(ArbString(myFormat_p(sourcetable!son_date_birth5)))
    
    temptable!str30 = TurnValue(sourcetable!son_name6)
    temptable!str31 = TurnValue(sourcetable!son_name7)
    temptable!str32 = TurnValue(sourcetable!son_name8)
    temptable!str33 = TurnValue(sourcetable!son_name9)
    temptable!str34 = TurnValue(sourcetable!son_name10)
    
    temptable!str35 = TurnValue(ArbString(myFormat_p(sourcetable!son_date_birth6)))
    temptable!str36 = TurnValue(ArbString(myFormat_p(sourcetable!son_date_birth7)))
    temptable!str37 = TurnValue(ArbString(myFormat_p(sourcetable!son_date_birth8)))
    temptable!str38 = TurnValue(ArbString(myFormat_p(sourcetable!son_date_birth9)))
    temptable!str39 = TurnValue(ArbString(myFormat_p(sourcetable!son_date_birth10)))
         
    temptable!str40 = TurnValue(sourcetable!parent_name1)
    temptable!str41 = TurnValue(sourcetable!parent_name2)
    temptable!str42 = TurnValue(sourcetable!parent_name3)
    temptable!str43 = TurnValue(sourcetable!parent_name4)
    temptable!str44 = TurnValue(sourcetable!parent_name5)
    
    temptable.Update
    sourcetable.MoveNext
Loop

temptable.Requery
    
If temptable.BOF And temptable.EOF Then
    Me.MousePointer = 0
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    Report1.Reset
    Report1.ProgressDialog = False
    Report1.WindowState = crptMaximized
    Report1.DataFiles(0) = tempFile
    If bPdf Then
        FixPrinter Report1, 1
        Report1.ReportFileName = sPath_App & "\REPORTS\CHECK1.RPT"
        Report1.Destination = crptToPrinter
    Else
        Report1.ReportFileName = sPath_App & "\REPORTS\CHECK1.RPT"
        Report1.Destination = crptToWindow
    End If
    Report1.Action = 1
    Me.MousePointer = 0
End If

Set temptable = Nothing
Set sourcetable = Nothing
End With
End Function

Private Sub cmdPdf_Click()
doprint True
End Sub

Private Sub cmdYear_Click(Index As Integer)
'Years_LookupAll Me, oSearchYear, , cmdYear(Index).Tag <> ""
End Sub

Private Sub DataCombo1_Click(Area As Integer)

End Sub

Private Sub Form_Load()
openCon con

FixRpImage Me
End Sub

Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub
Private Sub xType_GotFocus()
myGotFocus xType
End Sub
Private Sub xType_LostFocus()
myLostFocus xType
If Not xType.MatchedWithList Then xType.BoundText = ""
End Sub
Sub myProc()
'If ActiveControl.Name = cmdYear(0).Name Then
'    ActiveControl.Tag = oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 0)
'    ActiveControl.Caption = IIf(oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 0) = "", "«Œ «— «·„Ê”„", oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 1))
'    oSearchYear.Hide
'ElseIf ActiveControl.Name = xJob.Name Then
'    xJob.BoundText = oSearchJob.grid1.TextMatrix(oSearchJob.grid1.Row, 0)
'ElseIf ActiveControl.Name = xCompany.Name Then
'    xCompany.BoundText = oSearchComp.grid1.TextMatrix(oSearchComp.grid1.Row, 0)
'End If
End Sub

Private Sub xJob_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Job_Lookup Me, oSearchJob
End If
End Sub
Private Sub xJob_desca_GotFocus()
myGotFocus xJob_desca
End Sub
Private Sub xJob_desca_LostFocus()
myLostFocus xJob_desca
End Sub
Private Sub xTitle_GotFocus()
myGotFocus xtitle
End Sub
Private Sub xTitle_LostFocus()
myLostFocus xtitle
End Sub
Private Sub xJob_GotFocus()
myGotFocus xJob
End Sub
Private Sub xJob_LostFocus()
myLostFocus xJob
If Not xJob.MatchedWithList Then xJob.BoundText = ""
End Sub
Private Sub xCompany_GotFocus()
myGotFocus xCompany
End Sub
Private Sub xCompany_LostFocus()
myLostFocus xCompany
If Not xCompany.MatchedWithList Then xCompany.BoundText = ""
End Sub
Private Sub xCode1_GotFocus()
myGotFocus xcode1
End Sub
Private Sub xCode1_LostFocus()
myLostFocus xcode1
End Sub
Private Sub xCode2_GotFocus()
myGotFocus xCode2
End Sub
Private Sub xCode2_LostFocus()
myLostFocus xCode2
End Sub

