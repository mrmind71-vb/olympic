VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form TGrItem 
   BackColor       =   &H00FFFFFF&
   Caption         =   "≈Ã„«·Ï „»Ì⁄«  ⁄„Ì· "
   ClientHeight    =   6315
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   10005
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   10005
   Begin VB.OptionButton xInd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "«· — Ì» ⁄·Ï «·ﬁÌ„…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   375
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   600
      Width           =   1740
   End
   Begin VB.OptionButton xInd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "«· — Ì» ⁄·Ï «·ﬂÊœ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   600
      Value           =   -1  'True
      Width           =   1740
   End
   Begin VB.CommandButton Cmd_Exit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   300
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   75
      Width           =   1290
   End
   Begin VB.CommandButton Cmd_Print 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ÿ»«⁄…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1725
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   75
      Width           =   1290
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   3150
      Top             =   75
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3630
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   75
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSFlexGridLib.MSFlexGrid invGrid 
      Height          =   4890
      Left            =   300
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1125
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   8625
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedCols       =   0
      BackColor       =   12648447
      BackColorFixed  =   32896
      ForeColorFixed  =   16777215
      BackColorSel    =   65535
      BackColorBkg    =   16777215
      Enabled         =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2
      BorderStyle     =   0
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      Height          =   5190
      Left            =   75
      Top             =   975
      Width           =   9765
   End
   Begin VB.Label xDate2 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      Height          =   165
      Left            =   4050
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   300
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label xDate1 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      Height          =   165
      Left            =   4050
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   75
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label xDesc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Mudir MT"
         Size            =   10.5
         Charset         =   178
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5595
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   225
      Width           =   3120
   End
   Begin VB.Label xCode 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Mudir MT"
         Size            =   10.5
         Charset         =   178
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8925
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   225
      Width           =   840
   End
End
Attribute VB_Name = "TGrItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim invtable As Recordset
Dim TempTable As Recordset
Dim cString As String
Dim nInd As Byte
Sub Fillgrd()
I = 1
invGrid.Rows = 1
If formMode <> EmptyMode Then invGrid.Rows = 1
invtable.MoveFirst
Do
   invGrid.AddItem ""
   invGrid.TextArray(faIndex(I, 0, invGrid)) = TurnValue(invtable.CodeGroup, Null, "")
   invGrid.TextArray(faIndex(I, 1, invGrid)) = TurnValue(invtable.Desca_Group, Null, "")
   invGrid.TextArray(faIndex(I, 2, invGrid)) = TurnValue(invtable.SUMQUANT, Null, "")
   invGrid.TextArray(faIndex(I, 3, invGrid)) = TurnValue(Format(invtable.SUMTOTAL, "#0.00"), Null, "")
   invGrid.TextArray(faIndex(I, 4, invGrid)) = TurnValue(Format(invtable.SUMTOTAL - invtable.SUMCOST, "#0.00"), Null, "")
   invtable.MoveNext
   I = I + 1
Loop Until invtable.EOF
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub Cmd_Exit_Click()
Unload Me
Set TGrItem = Nothing
End Sub
Private Sub Cmd_Print_Click()
tempdb.Execute "Delete * From Temp"
Set TempTable = tempdb.OpenRecordset("TEMP")
With invtable
If .RecordCount > 0 Then
.MoveFirst
Do While Not .EOF
    TempTable.AddNew
    TempTable.str1 = .CodeGroup
    TempTable.str2 = .Desca_Group
    TempTable.val1 = .SUMQUANT
    TempTable.val2 = .SUMTOTAL
    TempTable.VAL3 = .SUMTOTAL - .SUMCOST
    TempTable.STR6 = "»Ì«‰ »≈Ã„«·Ï ﬁÌ„… „»Ì⁄«  ··„Ã„Ê⁄«  «·›—⁄Ì… ·„Ã„Ê⁄… —∆Ì”Ì… " & xDesc.Caption & " „‰  «—ÌŒ " & xdate1.Caption & " ≈·Ï  «—ÌŒ " & xDate2.Caption
'    TempTable.str19 = firsttitle
'    TempTable.str20 = Secondtitle
    TempTable.Update
    .MoveNext
Loop
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\Tsalitem.rpt"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End Sub
Private Sub Form_Load()
Dim cString As String
tempdb.Execute "Delete * From Temp"
'Me.Picture = LoadPicture(App.Path & "\mainback.jpg")
Set mydb = OpenDatabase(MdbPath)
xCode.Caption = TSalItem.xCode.Caption
xDesc.Caption = TSalItem.xDesc.Caption
xdate1.Caption = TSalItem.xdate1.Text
xDate2.Caption = TSalItem.xDate2.Text
'Unload TSalItem

invGrid.FormatString = "«·ﬂÊœ|" & "«·»Ì«‰|" & "≈Ã„«·Ï ﬂ„Ì…|" & "≈Ã„«·Ï ﬁÌ„…|" & "≈Ã„«·Ï —»Õ"
invGrid.Cols = 5
invGrid.ColWidth(0) = 1000
invGrid.ColWidth(1) = 3500
invGrid.ColWidth(2) = 1500
invGrid.ColWidth(3) = 1500
invGrid.ColWidth(4) = 1500
nInd = 1
cString = "SELECT FILE1_50.CODE AS CodeGroup, Sum(FILE6_20.QUANT) AS SumQUANT,  " & _
          "Sum(FILE6_20.TOTAL) AS SumTOTAL , Sum(FILE6_20.QUANT * FILE6_20.COST ) AS SumCOST, First(FILE1_50.DESCA) AS DESCA_GROUP " & _
          " FROM (FILE1_50 RIGHT JOIN FILE1_10 ON FILE1_50.CODE = FILE1_10.GROUP) RIGHT JOIN FILE6_20 ON FILE1_10.ITEM = FILE6_20.ITEM " & _
          " where file1_50.M_GROUP = " & MyParn(TSalItem.xCode.Caption) & _
          " And FILE6_20.DATE Between DateValue(" & MyParn(xdate1.Caption) & ") and DateValue(" & MyParn(xDate2.Caption) & ")" & _
          " GROUP BY FILE1_50.CODE"
If nInd = 1 Then
    cString = cString + " ORDER BY FILE1_50.CODE "
Else
    cString = cString + " ORDER BY Sum(FILE6_20.TOTAL) "
End If
Set invtable = mydb.CreateSnapshot(cString)
If invtable.RecordCount > 0 Then
    Fillgrd
Else
    invGrid.Rows = 1
End If
End Sub
Private Sub InvGrid_Click()
If invGrid.row > 0 Then
    xCode.Caption = GrdText(invGrid, 0)
    xDesc.Caption = GrdText(invGrid, 1)
    Load TItem
    TItem.Show 1
End If
End Sub
Private Sub xInd_Click(Index As Integer)
If xInd(Index).Value = True Then nInd = Index
cString = "SELECT FILE1_50.CODE AS CodeGroup, Sum(FILE6_20.QUANT) AS SumQUANT,  " & _
          "Sum(FILE6_20.TOTAL) AS SumTOTAL , Sum(FILE6_20.QUANT * FILE6_20.COST ) AS SumCOST, First(FILE1_50.DESCA) AS DESCA_GROUP " & _
          " FROM (FILE1_50 RIGHT JOIN FILE1_10 ON FILE1_50.CODE = FILE1_10.GROUP) RIGHT JOIN FILE6_20 ON FILE1_10.ITEM = FILE6_20.ITEM " & _
          " where file1_50.M_GROUP = " & MyParn(TSalItem.xCode.Caption) & _
          " And FILE6_20.DATE Between DateValue(" & MyParn(xdate1.Caption) & ") and DateValue(" & MyParn(xDate2.Caption) & ")" & _
          " GROUP BY FILE1_50.CODE"
If nInd = 1 Then
    cString = cString + " ORDER BY FILE1_50.CODE "
Else
    cString = cString + " ORDER BY Sum(FILE6_20.TOTAL) DESC "
End If
Set invtable = mydb.CreateSnapshot(cString)
If invtable.RecordCount > 0 Then
    Fillgrd
Else
    invGrid.Rows = 1
End If
End Sub
