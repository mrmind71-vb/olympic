VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form bankStatefrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ﬂ‘› Õ”«» «·»‰ﬂ"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2070
   ScaleWidth      =   5475
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
      Left            =   3105
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "⁄—÷ «·»Ì«‰« "
      Top             =   1440
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
      Left            =   1575
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "„”Õ «·ﬂ·"
      Top             =   1440
      Width           =   1500
   End
   Begin VB.CommandButton cmdExit 
      Height          =   555
      Left            =   45
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   1440
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Height          =   1410
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   0
      Width           =   5325
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1365
      End
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   945
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo xBank 
         Height          =   345
         Left            =   90
         TabIndex        =   0
         Top             =   180
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "«·»‰ﬂ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "„‰  «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   630
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«·Ï  «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   990
         Width           =   840
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   4275
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
End
Attribute VB_Name = "bankStatefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim ChargeTable As ADODB.Recordset
Private Sub cmdApply_Click()
doprint1
End Sub
Private Sub CmdClear_Click()
DefineText Me
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
FixRpImage Me
openCon con
Set ChargeTable = New ADODB.Recordset
ChargeTable.Open "File5_00", con, adOpenStatic, adLockReadOnly, adCmdTable

xdate1.Text = ""
xDate2.Text = ""

DATA1.ConnectionString = strCon
DATA1.RecordSource = "FILE5_10"

Set xBank.RowSource = DATA1
xBank.ListField = "Desca"
xBank.BoundColumn = "code"
End Sub
Function MYVALID() As Boolean
If Not IsDate(xdate1.Text) And (xdate1.Text <> "") Then
    MsgBox "Œÿ√ ›Ï «· «—ÌŒ"
    Exit Function
End If
If Not IsDate(xDate2.Text) And (xDate2.Text <> "") Then
    MsgBox "Œÿ√ ›Ï «· «—ÌŒ"
    Exit Function
End If
MYVALID = True
End Function
Private Sub doprint1()
If Not MYVALID Then Exit Sub
Dim i As Integer, nPrevious As Double
Dim aHeader(1)
Dim sourcetable As ADODB.Recordset
Dim temptable As ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
Set sourcetable = New ADODB.Recordset
aHeader(0) = " ﬂ‘› Õ”«» »‰ﬂ " & xBank.Text
If IsDate(xdate1.Text) Then
    Dim loctable As New ADODB.Recordset
    loctable.Open "select sum([value1]  - [value2] ) as Balance from bankmove where " & _
                  " bank = " & MyParn(xBank.BoundText) & _
                  " AND TYPE <= 4.5 " & _
                  " and [date] < " & DateSq(xdate1.Text), con, adOpenStatic, adLockReadOnly
    If Not loctable.EOF Then nPrevious = Val(loctable!BALANCE & "")
    If nPrevious <> 0 Then
        temptable.AddNew
        temptable!str1 = "—’Ìœ ”«»ﬁ"
        temptable!Val3 = nPrevious
        temptable!val1 = nPrevious
        temptable!str21 = TurnValue(retHeader(aHeader, 0, 1))
        temptable!str22 = TurnValue(retHeader(aHeader, 1, 1))
        temptable!Date1 = DateAdd("d", -1, xdate1.Text)
        temptable!val10 = 0
        temptable.Update
    End If
End If

cString = "Select * from BankMove Where TYPE <= 4.5 AND  BANK = " & MyParn(xBank.BoundText)
            
If IsDate(xdate1.Text) Then
    cString = cString & turnFound(cString) & "Date >= " & DateSq(xdate1.Text)
    aHeader(1) = BetweenString(Format(xdate1.Text, "d-m-yyyy"), Format(xDate2.Text, "d-m-yyyy"))
End If
          
If IsDate(xDate2.Text) Then
    cString = cString & turnFound(cString) & "Date <= " & DateSq(xDate2.Text)
    aHeader(1) = BetweenString(Format(xdate1.Text, "d-m-yyyy"), Format(xDate2.Text, "d-m-yyyy"))
End If
          
cString = cString & " Order by [Date],value1"
If sourcetable.State = adStateOpen Then sourcetable.Close
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

With sourcetable
Do Until sourcetable.EOF
    i = i + 1
    temptable.AddNew
    temptable!str1 = !DOC_NO
    temptable!str2 = !TypeDesca
    temptable!str3 = !desca
    temptable!Date1 = !Date
    temptable!val1 = !value1
    temptable!val2 = !Value2
    temptable!Val3 = nPrevious + Val(!value1 & "") - Val(!Value2 & "")
    nPrevious = nPrevious + Val(!value1 & "") - Val(!Value2 & "")
    temptable!str21 = TurnValue(retHeader(aHeader, 0, 1))
    temptable!str22 = TurnValue(retHeader(aHeader, 1, 1))
    temptable!val10 = i
    temptable.Update
    sourcetable.MoveNext
Loop
End With

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    main.REPORT1.ReportFileName = App.Path & "\Reports\BANK1.rpt"
    main.REPORT1.DataFiles(0) = "c:\tempmrshd\Temp.MDB"
    main.REPORT1.Action = 1
End If
If temptable.State = adStateOpen Then temptable.Close
If sourcetable.State = adStateOpen Then sourcetable.Close
Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Private Sub doprint2()
If Not MYVALID Then Exit Sub
Dim i As Integer, nPrevious As Double

Dim sourcetable As ADODB.Recordset
Dim temptable As ADODB.Recordset

Tempdb.Execute "DELETE * FROM TEMP"
Set temptable = New ADODB.Recordset
temptable.Open "temp", Tempdb, adOpenStatic, adLockOptimistic, adCmdTable
Set sourcetable = New ADODB.Recordset
If IsDate(xdate1.Text) Then
    cString = "Select Sum([IN])AS sumOfIN,Sum([OUT]) AS SUMOFOUT FROM FILE5_11 " & _
              " WHERE ( FILE5_11.Date < " & DateSq(xdate1.Text) & " or FILE5_11.TYPE = '1' ) " & _
              " AND BANK = " & MyParn(xBank.BoundText)
    
    sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
    If Not (sourcetable.EOF And sourcetable.BOF) Then
    nPrevious = TurnValue(sourcetable!SUMOFIN, Null, 0) - TurnValue(sourcetable!SUMOFOUT, Null, 0)
    temptable.AddNew
    temptable!str1 = "—’Ìœ ”«»ﬁ"
    'temptable!xdate1 = sourcetable!Date
    temptable!Val6 = nPrevious
    temptable!str7 = " ﬂ‘› Õ”«» «·»‰ﬂ " & xBank.Text & " „‰  «—ÌŒ " & xdate1.Text & " ≈·Ï  «—ÌŒ " & xDate2.Text
    temptable!xdate1 = Format(xdate1.Text, "YYYY-MM-DD")
    temptable!str19 = Firsttitle
    temptable!val10 = i
  '  temptable!str20 = SecondTitle
    temptable.Update
    End If
End If

cString = "Select * from File5_11 Where BANK = " & MyParn(xBank.BoundText)

If IsDate(xdate1.Text) Then
    cString = cString & turnFound(cString) & "Date >= " & DateSq(xdate1.Text)
End If
          
If IsDate(xDate2.Text) Then
    cString = cString & turnFound(cString) & "Date <= " & DateSq(xDate2.Text)
End If
          
cString = cString & " Order by [Date],[in]"
If sourcetable.State = adStateOpen Then sourcetable.Close
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
With sourcetable
    Do Until sourcetable.EOF
        If !Type <> "1" Then
            i = i + 1
            temptable.AddNew
    '        temptable!str1 = !DOC_NO
    '        temptable!str2 = TurnValue(RetFind(ChargeTable, "code", "Desca", !CODE), "", Null)
    '        temptable!str3 = TurnValue(!desca, "", Null)
            temptable!xdate1 = Format(!Date, "YYYY-MM-DD")
            If TurnValue(!In, Null, 0) > 0 And !Type = "2" Then
                temptable!val1 = !In
                temptable!Val6 = nPrevious + TurnValue(!In, Null, 0)
                nPrevious = TurnValue(temptable!Val6, Null, 0)
            End If
            
            If TurnValue(!In, Null, 0) > 0 And !Type = "3" Then
                temptable!val2 = !In
                temptable!Val6 = nPrevious + TurnValue(!In, Null, 0)
                nPrevious = TurnValue(temptable!Val6, Null, 0)
            End If
            
            If TurnValue(!out, Null, 0) > 0 And !Type = "2" And IsNull(!code) Then
                temptable!Val3 = !out
                temptable!Val6 = nPrevious - TurnValue(!out, Null, 0)
                    nPrevious = TurnValue(temptable!Val6, Null, 0)
            End If
            
            If TurnValue(!out, Null, 0) > 0 And !Type = "2" And Not IsNull(!code) Then
                temptable!val4 = !out
                temptable!Val6 = nPrevious - TurnValue(!out, Null, 0)
                nPrevious = TurnValue(temptable!Val6, Null, 0)
            End If
            
            If TurnValue(!out, Null, 0) > 0 And !Type = "4" Then
                temptable!Val5 = !out
                temptable!Val6 = nPrevious - TurnValue(!out, Null, 0)
                nPrevious = TurnValue(temptable!Val6, Null, 0)
            End If
                    
            temptable!str7 = " ﬂ‘› Õ”«» «·»‰ﬂ " & xBank.Text & " „‰  «—ÌŒ " & xdate1.Text & " ≈·Ï  «—ÌŒ " & xDate2.Text
            temptable!str19 = Firsttitle
       '     temptable!str20 = SecondTitle
            temptable!val10 = i
            temptable.Update
        End If
        sourcetable.MoveNext
    Loop
End With

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
    GoTo lastsub
End If
Tempdb.BeginTrans
Tempdb.CommitTrans
main.REPORT1.ReportFileName = App.Path & "\Report\BANK3.rpt"
main.REPORT1.DataFiles(0) = "c:\elmorshed\Temp.MDB"
main.REPORT1.Action = 1
lastsub:
If temptable.State = adStateOpen Then temptable.Close
If sourcetable.State = adStateOpen Then sourcetable.Close
Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xdate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xdate1
End Sub
Private Sub xDate1_Validate(Cancel As Boolean)
myValidDate xdate1
End Sub
Private Sub xDate2_Validate(Cancel As Boolean)
myValidDate xDate2
End Sub
Private Sub xdate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xdate2_LostFocus()
myLostFocus xDate2
End Sub
Private Sub xBank_GotFocus()
myGotFocus xBank
End Sub
Private Sub xBank_LostFocus()
myLostFocus xBank
End Sub
