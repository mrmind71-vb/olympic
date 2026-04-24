VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpitem1 
   Caption         =   "ÿ»«⁄… "
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   3810
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
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
      Left            =   3195
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "⁄—÷ «·»Ì«‰« "
      Top             =   3195
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
      Left            =   1665
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "„”Õ «·ﬂ·"
      Top             =   3195
      Width           =   1500
   End
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
      Left            =   135
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   3195
      Width           =   1500
   End
   Begin VB.Frame Frame2 
      Height          =   690
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   2475
      Width           =   6225
      Begin VB.CheckBox xcost 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   " ﬁÌÌ„ »”⁄— «· ﬂ·›…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2025
         RightToLeft     =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   225
         Width           =   1950
      End
      Begin VB.CheckBox xPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   " ﬁÌÌ„ »”⁄— «·»Ì⁄"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4125
         RightToLeft     =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   225
         Width           =   1740
      End
      Begin VB.CheckBox xNoBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "ÿ»«⁄… ··Ã—œ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   225
         Width           =   1695
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   2445
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   45
      Width           =   6180
      Begin VB.TextBox xitem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1980
         Width           =   1680
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1620
         Width           =   1680
      End
      Begin MSDataListLib.DataCombo xstore 
         Height          =   315
         Left            =   855
         TabIndex        =   3
         Top             =   1260
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   855
         TabIndex        =   2
         Top             =   900
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xGroupMain 
         Height          =   315
         Left            =   855
         TabIndex        =   1
         Top             =   540
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   855
         TabIndex        =   0
         Top             =   180
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "«·ﬁ”„ :"
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
         Index           =   3
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   270
         Width           =   1230
      End
      Begin VB.Label Label1 
         Caption         =   "«·„Ã„Ê⁄… «·—∆Ì”Ì… :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   630
         Width           =   1680
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«·’‰› :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   2115
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«· «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1755
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "«·„Ã„Ê⁄… :"
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
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   990
         Width           =   1230
      End
      Begin VB.Label Label4 
         Caption         =   "„Œ“‰ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1350
         Width           =   1005
      End
   End
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   5175
      Top             =   3015
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   180
      Top             =   630
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
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   2655
      Top             =   2520
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
   Begin VB.Label Label6 
      Height          =   255
      Left            =   4275
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2175
      Width           =   1005
   End
End
Attribute VB_Name = "rpitem1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub cmdApply_Click()
Dim aHeader(5)
If Not MYVALID Then Exit Sub
Dim temptable As ADODB.Recordset
Dim sourcetable As ADODB.Recordset
contemp.Execute "delete * from temp"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

cString = "SELECT Sum((FILE1_11.[IN])-(FILE1_11.[out] )) AS Balance,FILE1_10.ITEM,FILE1_10.DESCA, FILE1_10.[GROUP],FILE1_50.[GROUP] AS FILE1_50GROUP,FILE1_10.COST, file1_10.price , FILE1_50.DESCA AS FILE1_50DESCA,FILE1_50G.DESCA AS FILE1_50GDESCA" & _
          " FROM ((FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM) LEFT JOIN FILE1_50 ON FILE1_10.[GROUP] = FILE1_50.CODE) LEFT JOIN FILE1_50G ON FILE1_50.[GROUP] = FILE1_50G.CODE"


If IsNumeric(xSection.BoundText) Then
    cString = cString & turn(cString) & "File1_10.[SECTION] = " & xSection.BoundText
    aHeader(0) = "«·ﬁ”„ : " & xSection.Text
End If

If IsDate(xDate1.Text) Then
    cString = cString & turn(cString) & " date <= " & DateSq(xDate1.Text)
    aHeader(1) = "[" & "Õ Ì : " & xDate1.Text & "]"
End If

If Trim(xGroup.BoundText) <> "" Then
    cString = cString & turn(cString) & "File1_10.[GROUP] = " & xGroup.BoundText
    aHeader(2) = "„Ã„Ê⁄… " & xGroup.Text & "]"
End If

If Trim(xGroupMain.BoundText) <> "" Then
    cString = cString & turn(cString) & "File1_50.[GROUP] = " & xGroupMain.BoundText
    aHeader(3) = "„Ã„Ê⁄… —∆Ì”Ì…" & xGroup.Text & "]"
End If


If Trim(xstore.BoundText) <> "" Then
    cString = cString & turnFound(cString) & "File1_11.store = " & MyParn(xstore.BoundText)
    aHeader(4) = "[" & "«·„Œ“‰ " & xstore.Text & "]"
End If

If Trim(xitem.Text) <> "" Then
    cString = cString & turnFound(cString) & "File1_10.ITEM LIKE " & MyParn(xitem.Text & "%")
    aHeader(5) = "[" & "«·ﬂÊœ " & xitem.Text & "]"
End If

cString = cString & " GROUP BY FILE1_10.REORDER , FILE1_10.ITEM,file1_10.price , FILE1_10.DESCA, FILE1_10.[GROUP],FILE1_50.[GROUP],FILE1_10.COST, FILE1_50.DESCA,FILE1_50G.DESCA"
          
Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim cCondition As Boolean
With sourcetable
    Do Until .EOF

        If xNoBal.Value = 1 Then
            cCondition = True
        ElseIf xNoBal.Value = 0 Then
            cCondition = Val(sourcetable!Balance & "") <> 0
        End If
            
        If cCondition Then
            temptable.AddNew
            
            temptable!str7 = TurnValue(turn(![file1_50GDESCA] & "", "„Ã„Ê⁄… —∆Ì”Ì… : ") & ![file1_50GDESCA])
            temptable!str8 = TurnValue(turn(![file1_50desca] & "", "„Ã„Ê⁄… : ") & ![file1_50desca])
            
            temptable!str1 = !Item
            temptable!str2 = ![desca]
            temptable!val1 = !Balance
            nCost = LastCostDate(!Item, xDate1.Text, con)
'            If !Balance > 0 Then
'                'nCost = LastCost(!Item, con)
'                nCost = Round(RetitemCost(!Item, !Balance, xDate1.Text, con) / !Balance, 2)
'            Else
'                nCost = itemCost(!Item, con)
'            End If
            If nCost = 0 Then temptable!val2 = !cost Else temptable!val2 = nCost
            temptable!val3 = !Balance * temptable!val2
            temptable!val4 = !price
            temptable!val5 = !price * !Balance
            
            temptable!str21 = TurnValue(retHeader(aHeader, 0, 6))
            temptable.Update
        End If
      .MoveNext
    Loop
End With

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»«⁄ Â«"
Else

    If xNoBal.Value = 1 Then
        main.Report1.ReportFileName = App.Path & "\Reports\Item1_NoBal.rpt"
    ElseIf (xPrice.Value = 1 And xcost.Value = 1) Or (xPrice.Value = 0 And xcost.Value = 0) Then
        main.Report1.ReportFileName = App.Path & "\Reports\Item1.rpt"
    ElseIf xcost.Value = 1 Then
        main.Report1.ReportFileName = App.Path & "\Reports\Item1_cost.rpt"
    ElseIf xPrice.Value = 1 Then
        main.Report1.ReportFileName = App.Path & "\Reports\Item1_price.rpt"
    End If
    contemp.BeginTrans
    contemp.CommitTrans
    
    main.Report1.DataFiles(0) = tempFile
    main.Report1.Action = 1
End If
temptable.Close
sourcetable.Close
Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Private Sub CmdClear_Click()
xGroup.BoundText = ""
xstore.BoundText = ""
End Sub
Private Sub cmdexit_Click()
Unload Me
End Sub
Private Sub Form_Load()
FixRpImage Me
openCon con

data1.ConnectionString = strCon
data1.RecordSource = "Select Code,DescA From File1_10SC ORDER BY DESCA"
Set xSection.RowSource = data1
xSection.ListField = "Desca"
xSection.BoundColumn = "Code"

DATA2.ConnectionString = strCon
DATA2.RecordSource = "Select Code,DescA From File1_50G order by Desca"
Set xGroupMain.RowSource = DATA2
xGroupMain.ListField = "Desca"
xGroupMain.BoundColumn = "Code"

data3.ConnectionString = strCon
data3.RecordSource = "Select Code,DescA From File1_50 ORDER BY DESCA"
Set xGroup.RowSource = data3
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

data4.ConnectionString = strCon
data4.RecordSource = "Select Code,DescA From File0_40"
Set xstore.RowSource = data4
xstore.ListField = "Desca"
xstore.BoundColumn = "Code"
xcost.Visible = bopt1
End Sub
Private Function MYVALID() As Boolean
If Not IsDate(xDate1.Text) And Trim(xDate1.Text) <> "" Then
    MsgBox "«· «—ÌŒ €Ì— ’ÕÌÕ"
    Exit Function
End If
MYVALID = True
End Function

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

Private Sub xdate1_Validate(Cancel As Boolean)
myValidDate xDate1
End Sub

Private Sub xGroupMain_Validate(Cancel As Boolean)
If Not xGroupMain.MatchedWithList Then xGroupMain.BoundText = ""
data3.RecordSource = "Select Code,DescA From File1_50 " & IIf(xGroupMain.BoundText <> "", " where file1_50.[GROUP] = " & xGroupMain.BoundText, "")
data3.Refresh
End Sub
Private Sub xitem_GotFocus()
myGotFocus xitem
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xstore_GotFocus()
myGotFocus xstore
End Sub
Private Sub xGroup_GotFocus()
myGotFocus xGroup
End Sub
Private Sub xGroupMain_GotFocus()
myGotFocus xGroupMain
End Sub
Private Sub xSection_GotFocus()
myGotFocus xSection
End Sub

Private Sub xDate1_LostFocus()
myLostFocus xDate1
End Sub
Private Sub xstore_LostFocus()
myLostFocus xstore
End Sub
Private Sub xGroup_LostFocus()
myLostFocus xGroup
End Sub
Private Sub xGroupMain_LostFocus()
myLostFocus xGroupMain
End Sub
Private Sub xSection_LostFocus()
myLostFocus xSection
End Sub
