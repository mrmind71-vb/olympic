VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form inventory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ã—œ"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   ControlBox      =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Report1 
      Left            =   180
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "ÿ»«⁄… «·Ã—œ"
      Height          =   315
      Left            =   4185
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   60
      Width           =   1275
   End
   Begin Threed.SSCheck xSeq 
      Height          =   315
      Left            =   7725
      TabIndex        =   34
      Top             =   6375
      Width           =   1665
      _Version        =   65536
      _ExtentX        =   2937
      _ExtentY        =   556
      _StockProps     =   78
      Caption         =   "„  «»⁄"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin VB.CommandButton CmdAddAll 
      Caption         =   "≈÷«›… ﬂ· «·√’‰«›"
      Height          =   315
      Left            =   5460
      RightToLeft     =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   60
      Width           =   1590
   End
   Begin VB.CommandButton cmdDelInv 
      Caption         =   "«·€«¡ «·„” ‰œ"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2850
      RightToLeft     =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   60
      Width           =   1335
   End
   Begin VB.CommandButton cmdNewinv 
      Caption         =   "„” ‰œ ÃœÌœ"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1455
      RightToLeft     =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   60
      Width           =   1395
   End
   Begin VB.TextBox xItem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   7575
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2250
      Width           =   1665
   End
   Begin VB.TextBox xRealBal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   315
      Left            =   2175
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2250
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H80000018&
      Caption         =   "Õ›Ÿ"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1800
      Width           =   795
   End
   Begin VB.CommandButton CmdUndo 
      Caption         =   " —«Ã⁄"
      Enabled         =   0   'False
      Height          =   390
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   2220
      Width           =   795
   End
   Begin VB.CommandButton CmdChange 
      Caption         =   " €ÌÌ— «· «—ÌŒ Ê«·„Œ“‰"
      Enabled         =   0   'False
      Height          =   390
      Left            =   4815
      RightToLeft     =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6375
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   " —ÕÌ· «·„” ‰œ"
      Enabled         =   0   'False
      Height          =   390
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6375
      Width           =   1590
   End
   Begin VB.CommandButton cmdunPost 
      Caption         =   "«·€«¡  —ÕÌ· «·„” ‰œ"
      Enabled         =   0   'False
      Height          =   390
      Left            =   1650
      RightToLeft     =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6375
      Width           =   1590
   End
   Begin VB.CommandButton cmdFix 
      Caption         =   "«⁄«œ… ÷»ÿ «·Ã—œ"
      Enabled         =   0   'False
      Height          =   390
      Left            =   75
      RightToLeft     =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6375
      Width           =   1590
   End
   Begin VB.TextBox LastOne 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   -555
      MaxLength       =   2
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   1920
      Width           =   405
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "«÷«›… √’‰«›"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7800
      RightToLeft     =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   75
      Width           =   1515
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ "
      Height          =   315
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   60
      Width           =   1395
   End
   Begin VB.Frame Frame3 
      ForeColor       =   &H8000000D&
      Height          =   1290
      Left            =   75
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   360
      Width           =   9315
      Begin VB.CommandButton CmdApply 
         Caption         =   "«” Ã«»…"
         Height          =   345
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   855
         Width           =   1560
      End
      Begin VB.TextBox xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   3765
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   885
         Width           =   4410
      End
      Begin VB.TextBox xStore 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7770
         MaxLength       =   2
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   555
         Width           =   405
      End
      Begin VB.TextBox xDoc_Id 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   7275
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   900
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   285
         HideSelection   =   0   'False
         Left            =   180
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   300
         Width           =   1140
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "»Ì‹‹«‰ :"
         Height          =   195
         Left            =   8220
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   930
         Width           =   450
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·„Œ“‰ :"
         Height          =   195
         Left            =   8220
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   585
         Width           =   885
      End
      Begin VB.Label xclientName 
         Alignment       =   1  'Right Justify
         Height          =   240
         Left            =   3630
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   720
         Width           =   3465
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·„” ‰œ :"
         Height          =   195
         Left            =   8235
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   270
         Width           =   990
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   " «—ÌŒ «·„” ‰œ :"
         Height          =   195
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   330
         Width           =   1050
      End
   End
   Begin MSFlexGridLib.MSFlexGrid invGrid 
      Height          =   3480
      Left            =   75
      TabIndex        =   10
      Top             =   2775
      Visible         =   0   'False
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   6138
      _Version        =   393216
      Rows            =   1
      BackColorBkg    =   12632256
      Enabled         =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2
      BorderStyle     =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   2115
      X2              =   2115
      Y1              =   1845
      Y2              =   2625
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   975
      X2              =   9300
      Y1              =   2175
      Y2              =   2175
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000003&
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   810
      Left            =   975
      Shape           =   4  'Rounded Rectangle
      Top             =   1845
      Width           =   8340
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂÊœ «·’‰›"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   7875
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   1875
      Width           =   1260
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   7500
      X2              =   7500
      Y1              =   1875
      Y2              =   2625
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   4800
      X2              =   4800
      Y1              =   1875
      Y2              =   2625
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   3525
      X2              =   3525
      Y1              =   1860
      Y2              =   2640
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·’‰›"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   5475
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   1875
      Width           =   1050
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "—’Ìœ ›⁄·Ï"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   2250
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   1875
      Width           =   1005
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "—’Ìœ ﬂ„»ÌÊ —Ï"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   3525
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   1875
      Width           =   1245
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—ﬁ"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   1275
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   1875
      Width           =   675
   End
   Begin VB.Label xItemName 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4875
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2250
      Width           =   2565
   End
   Begin VB.Label xComputerBal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2250
      Width           =   1140
   End
   Begin VB.Label xDiffer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1050
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   2250
      Width           =   990
   End
End
Attribute VB_Name = "inventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim storeTable As Recordset, itemTable As Recordset, MoveTable As Recordset
Dim StockTable, inventTable As Recordset, TempTable As Recordset
Dim formMode, formFlag As Byte, lPosted As Boolean
Dim nBalance As Double, nDiffer As Double, nComputerBal
Const applyMode = 4, NewMode = 5, PostMode = 6, UnPostMode = 7
Sub AddProc()
myDefine
If Not formMode = addmode Then Handlecontrols addmode
formMode = addmode
InvGrid.AddItem ""
xItem.SetFocus
End Sub
Sub changeDatestoreProc()
fixinv
End Sub
Sub delinv()
cString = "Delete Distinctrow file0_10.DOC_ID From file0_10 WHERE file0_10.DOC_ID = " & MyParn(xDoc_Id.Text)
mydb.Execute cString
End Sub
Sub deleteItem()
pItem = GrdText(InvGrid, 0)
inventTable.Seek ">=", xDoc_Id.Text, pItem
If myFound(inventTable, Array(xDoc_Id.Text, pItem)) Then
    inventTable.Delete
    MyDelGrd InvGrid, InvGrid.Row
    inventTable.Seek ">=", xDoc_Id.Text
    If Not myFound(inventTable, Array(xDoc_Id.Text)) Then EmptyProc
End If
End Sub
Sub dispProc()
formMode = dispMode
myDefine
Handlecontrols dispMode
End Sub
Sub editProc()
formMode = Editmode
Handlecontrols Editmode
MyLoad
xRealBal.SetFocus
End Sub
Sub fixinv()
inventTable.Seek ">=", xDoc_Id.Text
If inventTable.NoMatch Then Exit Sub
Do While inventTable.DOC_ID = xDoc_Id.Text
    inventTable.Edit
    inventTable.ComputerBal = CountBalance(inventTable.Item, xStore.Text, xDate.Text)
    inventTable!Date = DateValue(xDate.Text)
    inventTable!Store = xStore.Text
    inventTable.differ = inventTable.RealBal - inventTable.ComputerBal
    inventTable.Update
    inventTable.MoveNext
    If inventTable.EOF Then Exit Do
Loop
Fillgrd
End Sub
Sub FillRow()
inventTable.Seek ">=", xDoc_Id.Text, xItem.Text
If myFound(inventTable, Array(xDoc_Id.Text, xItem.Text)) Then
    Myrow = InvGrid.Row
    InvGrid.TextArray(faIndex(Myrow, 0, InvGrid)) = xItem.Text
    InvGrid.TextArray(faIndex(Myrow, 1, InvGrid)) = RetSeek(itemTable, xItem.Text, "DescA")
    InvGrid.TextArray(faIndex(Myrow, 2, InvGrid)) = inventTable.ComputerBal
    InvGrid.TextArray(faIndex(Myrow, 3, InvGrid)) = inventTable.RealBal
    InvGrid.TextArray(faIndex(Myrow, 4, InvGrid)) = inventTable.differ
End If
End Sub
Sub Fillgrd()
i = 1
InvGrid.Rows = 1
inventTable.Seek ">=", xDoc_Id.Text
If inventTable.NoMatch Then Exit Sub
Do While inventTable.DOC_ID = xDoc_Id.Text
    InvGrid.AddItem ""
    InvGrid.TextArray(faIndex(i, 0, InvGrid)) = inventTable.Item
    InvGrid.TextArray(faIndex(i, 1, InvGrid)) = TurnValue(RetSeek(itemTable, inventTable.Item, TurnValue("DescA", Null, "")), Null, "")
    InvGrid.TextArray(faIndex(i, 2, InvGrid)) = inventTable.ComputerBal
    InvGrid.TextArray(faIndex(i, 3, InvGrid)) = inventTable.RealBal
    InvGrid.TextArray(faIndex(i, 4, InvGrid)) = inventTable.differ
    i = i + 1
    inventTable.MoveNext
    If inventTable.EOF Then Exit Do
Loop
End Sub
Sub MyDelGrd(grd As MSFlexGrid, Row)
If grd.Rows <= 2 Then grd.Rows = 1 Else grd.RemoveItem (Row)
End Sub
Sub MyReplace()
inventTable.DOC_ID = xDoc_Id.Text
inventTable.DESCA = TurnValue(xDesca.Text, "", Null)
inventTable.Store = xStore.Text
inventTable.Item = xItem.Text
inventTable.Date = xDate.Text
inventTable.ComputerBal = nBalance
inventTable.RealBal = Val(xRealBal.Text)
inventTable.differ = nDiffer
inventTable.CLOSED = False
End Sub
Sub ApplyProc()
Handlecontrols applyMode
inventTable.Seek ">=", xDoc_Id.Text
If Not myFound(inventTable, Array(xDoc_Id.Text)) Then
    Handlecontrols UnPostMode
    EmptyProc
Else
    xDesca.Text = TurnValue(inventTable.DESCA, Null, "")
    xDate.Text = Format(inventTable.Date, "dd/mm/yyyy")
    xStore.Text = inventTable.Store
    Handlecontrols IIf(inventTable.CLOSED, PostMode, UnPostMode)
    Fillgrd
    dispProc
End If
End Sub
Sub myDefine()
xItem.Text = ""
xItemName.Caption = ""
xDiffer.Caption = ""
xComputerBal.Caption = ""
xRealBal.Text = ""

End Sub
Sub MyLoad()
xRealBal.SetFocus
nRow = InvGrid.Row
pItem = GrdText(InvGrid, 0)
nPack = myRetPack(itemTable, pItem)
inventTable.Seek ">=", xDoc_Id.Text, pItem
If myFound(inventTable, Array(xDoc_Id.Text, pItem)) Then
    xItem.Text = inventTable.Item
    xItemName.Caption = RetSeek(itemTable, pItem, "Desca")
    xComputerBal.Caption = TurnValue(inventTable.ComputerBal, 0, "")
    xRealBal.Text = TurnValue(inventTable.RealBal, 0, "")
    nBalance = TurnValue(inventTable.ComputerBal, Null, 0)
    xDiffer.Caption = ""
    
    nDiffer = inventTable.differ
    nComputerBal = inventTable.ComputerBal
End If
End Sub
Sub myProc()
ActiveControl.Text = GrdText(Search.Grid1, 0)
Unload Search
End Sub
Function CountBalance(pItem, pStore, pDate)
MyQuant = 0
MoveTable.Seek ">=", pItem, pStore
If myFound(MoveTable, Array(pItem, pStore)) Then
    Do While MoveTable.Item = pItem
        If MoveTable.Store = pStore And DateValue(MoveTable.Date) <= DateValue(pDate) And Not MoveTable!Type = "" Then
            MyQuant = MyQuant + TurnValue(MoveTable.In, Null, 0) - TurnValue(MoveTable.OUT, Null, 0)
        End If
        MoveTable.MoveNext
    If MoveTable.EOF Then Exit Do
    Loop
End If
CountBalance = MyQuant
End Function
Function myValid2()
' **************
 itemTable.Seek "=", xItem.Text
If itemTable.NoMatch Then Exit Function
If formMode <> Editmode Then
    inventTable.Seek ">=", xDoc_Id.Text, xItem.Text
    If myFound(inventTable, Array(xDoc_Id.Text, xItem.Text)) Then
        MsgBox "Â–« «·’‰› „ÊÃÊœ ›Ï «·Ã—œ „‰ ﬁ»·"
        Exit Function
    End If
End If
myValid2 = True
End Function
Function MYVALID()
    If xDoc_Id.Text = "" Then Exit Function
    If Not IsDate(xDate.Text) Then Exit Function
    MYVALID = True
End Function
Sub postinv(invcode)
Dim pinvTable As Recordset
Set pinvTable = mydb.OpenRecordset("Select * from file0_10 where Doc_id = " & MyParn(invcode), dbOpenDynaset)
Do
    If pinvTable.differ <> 0 Then
    MoveTable.AddNew

    MoveTable!Type = "z"
    MoveTable.Item = pinvTable.Item
    MoveTable!Date = pinvTable.[Date]
    MoveTable!Store = pinvTable.Store
    MoveTable.DESCA = TurnValue(pinvTable.DESCA & " ›Ï ", "", "÷»ÿ Ã—œ ›Ï ") & pinvTable.Date
    MoveTable.DOC_ID = pinvTable.DOC_ID
'   movetable.cgroup = RetSeek(itemTable, pinvTable.Item, "Group")
    MoveTable.In = pinvTable.differ
    MoveTable.Update
    End If
    pinvTable.Edit
    pinvTable.CLOSED = True
    pinvTable.Update
    
    pinvTable.MoveNext
Loop Until pinvTable.EOF
End Sub
Sub unpostinv(invcode)
cString = "Delete Distinctrow file1_11.Code From file1_11 WHERE file1_11.doc_id= " & MyParn(invcode) & " and File1_11.type = 'z'"
mydb.Execute cString

cString = "Update DistinctRow file0_10 SET file0_10.closed =false WHERE file0_10.Doc_Id = " & MyParn(invcode)
mydb.Execute cString
Handlecontrols UnPostMode
End Sub
Sub Undoinv()
Select Case formMode
Case addmode
    InvGrid.Rows = InvGrid.Rows - 1
    dispProc
Case Editmode
    dispProc
Case EmptyMode
    myDefine
End Select
End Sub
Private Sub cmdAddAll_Click()
If lPosted Then Exit Sub
itemTable.MoveFirst
Do While Not itemTable.EOF
    inventTable.Seek ">=", xDoc_Id.Text, itemTable.Item
    If Not myFound(inventTable, Array(xDoc_Id.Text, itemTable.Item)) Then
        inventTable.AddNew
        inventTable.DOC_ID = xDoc_Id.Text
        inventTable.DESCA = TurnValue(xDesca.Text, "", Null)
        inventTable.Item = itemTable.Item
        inventTable.Date = xDate.Text
        inventTable.Store = xStore.Text
        inventTable.ComputerBal = CountBalance(itemTable.Item, xStore.Text, xDate.Text)
        inventTable.RealBal = 0
        inventTable.differ = inventTable.ComputerBal
        inventTable.CLOSED = "0"
        inventTable.Update
    End If
    itemTable.MoveNext
Loop
dispProc
Fillgrd
End Sub
Private Sub CmdApply_Click()
  If MYVALID Then ApplyProc
End Sub
Private Sub CmdChange_Click()
Load ChangeForm
ChangeForm.Show 1
End Sub
Private Sub cmdDelinv_Click()
If lPosted Then Exit Sub
If MsgBox("«·€«¡ «·„” ‰œ ! Â· «‰  „Ê«›ﬁ ø", 4 + 256) = vbNo Then Exit Sub
delinv
CmdNewInv_Click
End Sub

Private Sub cmdFix_Click()
'amramramramr
fixinv
'Fillgrd
'handleCommands
End Sub
Private Sub cmdNew_Click()
If lPosted Then Exit Sub
AddProc
End Sub
Private Sub CmdNewInv_Click()
InvGrid.Rows = 1
xDoc_Id.Text = ""
xDate.Text = ""
xDesca.Text = ""
xStore.Text = ""
Handlecontrols NewMode
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Command3_Click()
unpostinv ("1")
End Sub
Private Sub CmdPost_Click()
postinv (xDoc_Id.Text)
Handlecontrols PostMode
lPosted = True
End Sub
Private Sub cmdSave_Click()
If Not myValid2 Then Exit Sub
Select Case formMode
Case addmode
    InvGrid.Row = InvGrid.Rows - 1
    inventTable.AddNew
    MyReplace
    inventTable.Update
    FillRow
    AddProc
    
Case EmptyMode
    InvGrid.Row = 1
    InvGrid.Enabled = True
    inventTable.AddNew
    MyReplace
    inventTable.Update
    FillRow
    AddProc

Case Editmode
    inventTable.Seek ">=", xDoc_Id.Text, xItem.Text
    inventTable.Edit
    MyReplace
    inventTable.Update
    FillRow
    If xSeq.Value And InvGrid.Row < InvGrid.Rows - 1 Then
        InvGrid.Row = InvGrid.Row + 1
        editProc
        MyLoad
    Else
        dispProc
    End If
End Select
End Sub

Private Sub cmdSave_GotFocus()
 If myValid2 Then
   nDiffer = Val(xRealBal.Text) - Val(xComputerBal.Caption)
   xDiffer.Caption = nDiffer
End If
End Sub
Private Sub cmdunPost_Click()
unpostinv (xDoc_Id.Text)
Handlecontrols UnPostMode
lPosted = False
End Sub
Private Sub CmdUndo_Click()
Undoinv
End Sub

Private Sub CmdPrint_Click()
PrintAllItems
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If TypeOf ActiveControl Is TextBox And KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub Form_Load()
Set storeTable = mydb.OpenRecordset("Stores")
Set MoveTable = mydb.OpenRecordset("File1_11")
Set StockTable = mydb.OpenRecordset("File1_30")
Set itemTable = mydb.OpenRecordset("file1_10")
Set inventTable = mydb.OpenRecordset("file0_10")
Set TempTable = tempdb.OpenRecordset("temp")

MoveTable.Index = "nFile1_11_3"
StockTable.Index = "nFile1_30_1"
itemTable.Index = "nFile1_10_1"
inventTable.Index = "nFile0_10_1"
InvGrid.FormatString = "ﬂÊœ «·’‰›|" & "«”„ «·’‰›|" & "—’Ìœ ﬂ„»ÌÊ —Ï|" & "—’Ìœ ›⁄·Ï|" & "«·›—ﬁ"
InvGrid.Cols = 5
InvGrid.ColWidth(0) = 1000
InvGrid.ColWidth(1) = 4000
InvGrid.ColWidth(2) = 1500
InvGrid.ColWidth(3) = 1500
InvGrid.ColWidth(4) = 1000
xDoc_Id.Text = IncRec(myLastField(inventTable, "DOC_ID"))
End Sub
Sub EmptyProc()
myDefine
If Not formMode = EmptyMode Then Handlecontrols EmptyMode
formMode = EmptyMode
InvGrid.AddItem ""
If xItem.Enabled Then xItem.SetFocus
End Sub
Sub Handlecontrols(pMode)
Select Case pMode
Case dispMode
    xItem.Enabled = False
    xRealBal.Enabled = False

    CmdUndo.Enabled = False
    CmdSave.Enabled = False
    cmdNew.Enabled = True
    InvGrid.Enabled = True
Case Editmode
    xItem.Enabled = False
    xRealBal.Enabled = True

    CmdUndo.Enabled = True
    CmdSave.Enabled = True
    cmdNew.Enabled = False
'    cmdNewinv.Enabled = False
'    cmdDelInv.Enabled = False
'    cmdPost.Enabled = False
'    cmdunPost.Enabled = False
'    cmdFix.Enabled = False
Case addmode
    xItem.Enabled = True
    xRealBal.Enabled = True

    CmdUndo.Enabled = True
    CmdSave.Enabled = True
'    cmdNewInv.Enabled = False
'    cmdDelInv.Enabled = False
'    cmdPost.Enabled = False
'    cmdunPost.Enabled = False
'    invGrid.Enabled = False
Case EmptyMode
    xItem.Enabled = True
    xRealBal.Enabled = True

    xDiffer.Enabled = True
    cmdNew.Enabled = False
    CmdUndo.Enabled = True
    CmdSave.Enabled = True
Case applyMode
    xDate.Enabled = False
    xDoc_Id.Enabled = False
    xDesca.Enabled = False
    xStore.Enabled = False
    CmdApply.Enabled = False
    InvGrid.Enabled = True
    cmdNewinv.Enabled = True
    cmdNew.Enabled = True
    InvGrid.Visible = True
Case NewMode
    xDoc_Id.Enabled = True
    xDate.Enabled = True
    xDesca.Enabled = True
    xStore.Enabled = True
'    cmdNewInv.Enabled = False
    CmdFix.Enabled = False
    cmdPost.Enabled = False
    cmdunPost.Enabled = False
    CmdDelInv.Enabled = False
    CmdApply.Enabled = True
    InvGrid.Enabled = False
Case PostMode
    CmdFix.Enabled = False
    cmdPost.Enabled = False
    cmdunPost.Enabled = True
    CmdDelInv.Enabled = False
    cmdNew.Enabled = False
    CmdChange.Enabled = False
    CmdAddAll.Enabled = False
Case UnPostMode
    CmdFix.Enabled = True
    cmdPost.Enabled = True
    cmdunPost.Enabled = False
    CmdDelInv.Enabled = True
    cmdNew.Enabled = True
    CmdChange.Enabled = True
    CmdAddAll.Enabled = True
End Select
End Sub
Private Sub invGrid_DblClick()
If InvGrid.Rows = 1 Then Exit Sub
If lPosted Then Exit Sub
If formMode <> dispMode Then Exit Sub
editProc
End Sub
Private Sub invGrid_KeyUp(KeyCode As Integer, Shift As Integer)
If formMode <> dispMode Then Exit Sub
If lPosted Then Exit Sub
If KeyCode = 46 Then
    If MsgBox("«·€«¡ «·’‰› !! Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then deleteItem
    Exit Sub
End If
End Sub
Private Sub xClient_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{Tab}"
If KeyCode = 112 Then
    Dim Generalarray(2)
    Dim GrdArray(2, 3)
    
        Generalarray(1) = "Select * From Clients"
    Set Generalarray(2) = Form2
    
    GrdArray(0, 0) = "Code"
    GrdArray(0, 1) = 1000
    GrdArray(0, 2) = "—ﬁ„ «·’‰›"
    
    GrdArray(1, 0) = "Desca"
    GrdArray(1, 1) = 1500
    GrdArray(1, 2) = " «—ÌŒ «·’·«ÕÌ…"
    
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End If
End Sub
Private Sub xClient_LostFocus()
If xClient.Text = "" Then Exit Sub
If xClient.Enabled = False Then Exit Sub
If Not IsDate(xDate.Text) Or xDate.Text = "" Then Exit Sub
ClientTable.FindFirst "code = '" & xClient.Text & "'"
If ClientTable.NoMatch Then Exit Sub
xClient.Enabled = False
xclientName.Caption = ClientTable.DESCA
ApplyProc
End Sub

Private Sub xDoc_Id_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    xDoc_Id.Text = ""
    Dim Generalarray(4)
    Dim GrdArray(4)
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Doc_Id as [—ﬁ„ «·„” ‰œ],Date as [«· «—ÌŒ],STORE AS [„Œ“‰],iif(Closed,'„—Õ·…','€Ì— „—Õ·…') as [„—Õ·…] From File0_10 "
    Generalarray(3) = "Where Doc_id Like '*cFilter*' or [„—Õ·…] Like '*cFilter*' "
    Generalarray(4) = "Group by Doc_id,Date,Closed,STORE"
    GrdArray(1) = 1000
    GrdArray(2) = 3000
    GrdArray(3) = 500
    GrdArray(4) = 1000
         
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End If
End Sub
Private Sub xDoc_Id_LostFocus()
If xDoc_Id.Text = "" Then Exit Sub
inventTable.Seek ">=", xDoc_Id.Text
If myFound(inventTable, Array(xDoc_Id.Text)) Then ApplyProc
End Sub
Private Sub xItem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then ItemsLookup
End Sub
Sub ItemsLookup()
ActiveControl.Text = ""
Dim Generalarray(4)
Dim GrdArray(3)
    
Set Generalarray(1) = Me
Generalarray(2) = "Select Item as «·’‰›,DescA,pack as [«”„ «·’‰›] From file1_10 as [«·»⁄Ê…] "
Generalarray(3) = " Where DescA Like('*cFilter*')"
Generalarray(4) = "Order by Item"
    
GrdArray(1) = 1000
GrdArray(2) = 3500
GrdArray(3) = 1500

Lookupdata = Array(Generalarray, GrdArray)
Load Search
Search.Caption = "«” ⁄·«„ "
Search.Show 1
End Sub
Private Sub xItem_LostFocus()
If xItem.Text = "" Then Exit Sub

itemTable.Seek "=", xItem.Text
If itemTable.NoMatch Then Exit Sub

inventTable.Seek "=", xDoc_Id.Text, xItem.Text
If Not inventTable.NoMatch Then
    MsgBox "Â–« «·’‰› „ÊÃÊœ »«·›« Ê—… „‰ ﬁ»·"
    Exit Sub
End If
nBalance = CountBalance(xItem.Text, xStore.Text, xDate.Text)
xItemName.Caption = itemTable.DESCA
xComputerBal.Caption = nBalance
End Sub
Private Sub xItemName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Undoinv
    Exit Sub
End If
End Sub
Private Sub xRealBal_KeyPress(KeyAscii As Integer)
KeyAscii = RetNumber(KeyAscii, True)
End Sub
Private Sub xStore_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Undoinv
    Exit Sub
End If

If KeyCode = 112 Then
    ActiveControl.Text = ""

    Dim Generalarray(3)
    Dim GrdArray(2)
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Code as [«·„”·”·],DescA as [«·«”„] From Stores "
    Generalarray(3) = "Where DescA Like '*cFilter*'"
    GrdArray(1) = 1000
    GrdArray(2) = 4000
    
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End If
End Sub
Private Sub xStore_LostFocus()
If Not xStore.Enabled Then Exit Sub
If xStore.Text = "" Then Exit Sub
End Sub
Sub PrintAllItems()
tempdb.Execute "DELETE * FROM TEMP"
i = 1
InvGrid.Rows = 1
inventTable.Seek ">=", xDoc_Id.Text
If inventTable.NoMatch Then Exit Sub
Do While inventTable.DOC_ID = xDoc_Id.Text
    If inventTable.RealBal <> 0 Or inventTable.differ <> 0 Then
    TempTable.AddNew
    TempTable.str1 = "„” ‰œ Ã—œ —ﬁ„ :" & " " & xDoc_Id.Text & "  » «—ÌŒ :  " & xDate.Text
    TempTable.str2 = xDesca.Text
    TempTable.str3 = inventTable.Item
    TempTable.str4 = RetSeek(itemTable, inventTable.Item, "DescA")
    TempTable.STR5 = inventTable.ComputerBal
    TempTable.str6 = inventTable.RealBal
    TempTable.str7 = inventTable.differ
    TempTable.str19 = firsttitle
    TempTable.str20 = Secondtitle
    TempTable.Update
    End If
    inventTable.MoveNext
    If inventTable.EOF Then Exit Do
Loop
Report1.ReportFileName = PublicPath & "\Reports\Invent.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.MDB"
Report1.Action = 1
End Sub
