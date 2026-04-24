VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form ClientReport5 
   Caption         =   " ﬁ«—Ì— «·⁄„·«¡"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4710
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
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2055
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -900
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   810
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox xdate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2250
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   675
      Width           =   1290
   End
   Begin VB.TextBox xDate2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   675
      Width           =   1290
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1350
      Width           =   3240
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   390
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   990
      End
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
         Height          =   390
         Left            =   1050
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   225
         Width           =   1065
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   390
         Left            =   2100
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   1065
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   4080
      Top             =   1530
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
   End
   Begin MSDBCtls.DBCombo xGroup 
      Bindings        =   "Clrep5.frx":0000
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   "DBCombo1"
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„‰  «—ÌŒ :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3630
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   705
      Width           =   765
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "«·Ï :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   630
      Width           =   390
   End
   Begin VB.Label iLabel 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4485
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·„Ã„Ê⁄…:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   225
      Width           =   795
   End
End
Attribute VB_Name = "ClientReport5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ClientTable As Recordset
Dim TempTable As Recordset
Function CountBalance(Client)
Dim ClientMove As Recordset
NRETURN = 0
Set ClientMove = mydb.CreateDynaset("select * from file3_11 where code = " & MyParn(Client))
If ClientMove.RecordCount = 0 Then
    CountBalance = 0
    Exit Function
End If
ClientMove.MoveFirst
Do
    NRETURN = NRETURN + TurnValue(ClientMove.sal, Null, 0) - TurnValue(ClientMove.PAY, Null, 0)
    ClientMove.MoveNext
Loop Until ClientMove.EOF
CountBalance = NRETURN
End Function
Sub Groupitems()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
If Not MYVALID Then Exit Sub
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")
cField1 = myiif( _
               "Type = '6' OR type = '0'", "[out]") & _
               " As SalesQuant ,sum([out2]) as  SalesQuant2"

cField2 = myiif( _
               "Type = '3'", "[IN]") & "as RetSalesQuant,sum(in2) As RetSalesQuant2"

cField3 = myiif( _
               "Type = '6' OR Type = '0'", "Total") & _
               " As SalesValue"

cField4 = myiif( _
               "Type = '3'", "Total") & _
               " As RetSalesValue"

cString = "Select Mid(File1_11.Item,1,6) as MidOfItem," & _
          "File1_10.DescA,File1_10.Pack,File1_10.[Group]," & _
          "File3_10.Group," & _
          cField1 & "," & cField2 & "," & cField3 & "," & cField4 & _
          " From (File1_11 Inner Join File1_10 on File1_11.Item = File1_10.Item)" & _
                " Inner Join File3_10 on File1_11.Code = File3_10.Code " & _
          " Where Date >= DateValue(" & MyParn(xdate1.Text) & ")" & _
          " and Date <= DateValue(" & MyParn(xDate2.Text) & ")" & _
          " And File3_10.Group = " & MyParn(xGroup.BoundText) & _
          "Group By " & _
          " Mid(File1_11.Item,1,6)," & _
          "File1_10.DescA,File1_10.Pack,File1_10.[Group]," & _
          "File3_10.Group"
Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
With SourceTable
If .RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
Do
    TargetTable.AddNew
    TargetTable.str1 = SourceTable![file3_10.Group]
    TargetTable.str2 = xGroup.Text
    TargetTable.str3 = SourceTable.MidofItem
    TargetTable.str4 = SourceTable.DESCA
    TargetTable.str9 = " : ’«›Ï „»Ì⁄«  «·«’‰«› ··„Ã„Ê⁄…"
    
    TargetTable.val1 = .SALESQUANT
    TargetTable.val2 = .RetSalesQuant
    TargetTable.VAL3 = .SALESQUANT - .RetSalesQuant
    TargetTable.VAL4 = .SalesQuant2
    TargetTable.VAL5 = .RetSalesQuant2
    TargetTable.VAL6 = .SalesQuant2 - .RetSalesQuant2
    
    TargetTable.VAL7 = .salesvalue
    TargetTable.VAL8 = .retsalesvalue
    TargetTable.VAL9 = .salesvalue - .retsalesvalue
    
    TargetTable.Date1 = xdate1.Text
    TargetTable.date2 = xDate2.Text
    TargetTable.str19 = firsttitle
    TargetTable.str20 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
Report1.ReportFileName = PublicPath & "\Reports\Client8.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.mdb"
Report1.Action = 1
End With
End Sub
Sub myProc()
ActiveControl.Text = GrdText(Search.Grid1, 0)
Unload Search
End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
cSelection = ""
Report1.DataFiles(0) = App.Path & "\Temp.MDB"
Select Case publicFlag
Case 1, 2
    GroupInvtotal
Case 3
    Groupitems
End Select
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
Data1.DatabaseName = MdbPath
Data1.RecordSource = "ClGroup"
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"
Data1.Refresh

End Sub
Function MYVALID()
If xGroup.BoundText = "" Then Exit Function
If Not IsDate(xdate1.Text) And xdate1.Text <> "" Then Exit Function
If Not IsDate(xDate2.Text) And xDate2.Text <> "" Then Exit Function
If DateValue(xdate1.Text) > DateValue(xDate2.Text) Then Exit Function
MYVALID = True
End Function
Private Sub xClient_LostFocus()
If xClient.Text = "" Then Exit Sub
ClientTable.FindFirst "Code = " & MyParn(xClient.Text)
If ClientTable.NoMatch Then Exit Sub
ClientName.Caption = ClientTable.DESCA
End Sub
Private Sub GroupInvtotal()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
If publicFlag = 1 Then
    cString = "Select File6_20.DOC_NO," & _
              "File6_20.DATE," & _
              "File3_10.Group," & _
              "Sum(File6_20.Total) as sumTotal " & _
              "From file6_20 Inner Join file3_10 on file6_20.CODE = file3_10.Code " & _
              " Where File3_10.Group = " & MyParn(xGroup.BoundText) & _
              " And Date Between DateValue(" & MyParn(xdate1.Text) & ") and DateValue(" & MyParn(xDate2.Text) & ")" & _
              " Group by File6_20.DOC_NO,File3_10.Group,Date" & _
              " Order by Date "
Else
    cString = "Select File6_10.DOC_NO," & _
              "File6_10.DATE," & _
              "File3_10.Group," & _
              "Sum(File6_10.Total) as sumTotal " & _
              "From file6_10 Inner Join file3_10 on file6_10.CODE = file3_10.Code " & _
              " Where File3_10.Group = " & MyParn(xGroup.BoundText) & _
              " And Date Between DateValue(" & MyParn(xdate1.Text) & ") and DateValue(" & MyParn(xDate2.Text) & ")" & _
              " Group by File6_10.DOC_NO,File3_10.Group,Date" & _
              " Order by Date "
End If
tempdb.Execute "DELETE * FROM TEMP"
Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
Set TargetTable = tempdb.OpenRecordset("TEMP")
If SourceTable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
Do
    TargetTable.AddNew
    TargetTable.str1 = SourceTable.DOC_NO
    TargetTable.str2 = xGroup.Text
    TargetTable.str4 = IIf(publicFlag = 1, "≈Ã„«·Ï „»Ì⁄«  „Ã„Ê⁄…", "≈Ã„«·Ï „—œÊœ«  „Ã„Ê⁄… ")
    
    TargetTable.val1 = SourceTable.SUMTOTAL
    TargetTable.Date1 = xdate1.Text
    TargetTable.date2 = xDate2.Text
    TargetTable.date3 = SourceTable.Date
    
    TargetTable.str19 = firsttitle
    TargetTable.str20 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
Report1.ReportFileName = PublicPath & "\Reports\Comp5.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.MDB"
Report1.Action = 1
End Sub
