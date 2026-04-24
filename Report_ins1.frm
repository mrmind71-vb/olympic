VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form report_insfrm1 
   Caption         =   "«⁄œ«œ «·«⁄÷«¡ «· ﬁ”Ìÿ"
   ClientHeight    =   1875
   ClientLeft      =   1065
   ClientTop       =   1875
   ClientWidth     =   4905
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
   ScaleHeight     =   1875
   ScaleWidth      =   4905
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
      Left            =   180
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   1215
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
      Left            =   1710
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "„”Õ «·ﬂ·"
      Top             =   1215
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
      Left            =   3240
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "⁄—÷ «·»Ì«‰« "
      Top             =   1215
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Height          =   1050
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   135
      Width           =   4515
      Begin MSDataListLib.DataCombo xType 
         Height          =   330
         Left            =   135
         TabIndex        =   1
         Top             =   540
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xStatus 
         Height          =   330
         Left            =   135
         TabIndex        =   0
         Top             =   180
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "›∆… «·⁄÷ÊÌ…"
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
         Index           =   1
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   585
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "Õ«·… «·⁄÷Ê"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   960
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   3510
      Top             =   225
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   2970
      Top             =   180
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
End
Attribute VB_Name = "report_insfrm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim oSearchYear As New Search_empty
Private Sub cmdApply_Click()
doprint
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Function doprint()
Dim temptable As New ADODB.Recordset, sourcetable As New ADODB.Recordset
Dim aHeader(2)
cString = "Select Count(*) as countofMember,FILE2_10.[type],TYPE_CODES.DescA " & _
          " From File2_10 left join TYPE_codes on FILE2_10.TYPE = TYPE_codes.Code"

If xStatus.MatchedWithList Then
    aHeader(0) = "Õ«·… «·⁄÷ÊÌ… : " & xStatus.Text
    cWhere = cWhere & turn(cWhere, " and ") & "FILE2_10.STATUS = " & addvalue(xStatus.BoundText)
End If

If xType.MatchedWithList Then
    aHeader(1) = "›∆… «·⁄÷ÊÌ… : " & xType.Text
    cWhere = cWhere & turn(cWhere, " and ") & "FILE2_10.TYPE = " & addvalue(xType.BoundText)
End If

If cWhere <> "" Then cString = cString & " WHERE " & cWhere

cString = cString & " Group By FILE2_10.[type],type_codes.Desca"

sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adcmtext

contemp.Execute "delete * from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adcmtext

With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!Val10 = sourcetable![Type]
    temptable!val1 = sourcetable!CountOfMember
    temptable!val2 = 0
    temptable!VAL3 = 0
    temptable!val4 = 0
    temptable!str2 = sourcetable!Desca
    temptable!str11 = TurnValue(Me.Caption)
    temptable!str12 = TurnValue(retHeader(aHeader, 0, 3))
    temptable!str13 = TurnValue(retHeader(aHeader, 3, 5))
    temptable.Update
    sourcetable.MoveNext
Loop

temptable.Requery
    
cField = myiif("Relation = 1", 1) & " as countofwife"
cField = cField & "," & myiif("Relation = 2", 1) & " as countofSon"
cField = cField & "," & myiif("Relation > 2", 1) & " as countofRel"

cString = "Select FILE2_10.[type],type_codes.Desca," & _
          cField & _
          " From FILE2_10 inner join FILE2_11 on FILE2_10.code = FILE2_11.Member left join type_codes on FILE2_10.[type] = type_codes.code "
          
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
cString = cString & " Group by FILE2_10.[TYPE] ,TYPE_codes.Desca"

sourcetable.Close
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adcmtext
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str2 = sourcetable!Desca
    temptable!Val10 = sourcetable![Type]
    temptable!val2 = sourcetable!countOfWife
    temptable!VAL3 = sourcetable!countOfSon
    temptable!val4 = sourcetable!countOfRel
    temptable!str11 = TurnValue(Me.Caption)
    temptable!str12 = TurnValue(retHeader(aHeader, 0, 3))
    temptable!str13 = TurnValue(retHeader(aHeader, 3, 5))
    temptable.Update
    sourcetable.MoveNext
Loop


If temptable.BOF And temptable.EOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    
    Report1.ReportFileName = sPath_App & "\REPORTS\REPORT_INS1.rpt"
    Report1.DataFiles(0) = tempFile
    Report1.Action = 1
End If

Set temptable = Nothing
Set sourcetable = Nothing
End With
End Function
Private Sub Form_Load()
openCon con

Set data1.Recordset = myRecordSet("SELECT CODE,DESCA FROM TYPE_CODES ORDER BY CODE", con)
Set xType.RowSource = data1
xType.ListField = "Desca"
xType.BoundColumn = "Code"

Set data2.Recordset = myRecordSet("select * from status_Codes", con)
Set xStatus.RowSource = data2
xStatus.ListField = "Desca"
xStatus.BoundColumn = "Code"

FixRpImage Me

LoadText Me
xStatus.BoundText = 1

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
