VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form itemMove 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Õ—ﬂ… «·√’‰«›"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   FillColor       =   &H80000000&
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
   MousePointer    =   1  'Arrow
   RightToLeft     =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin VB.TextBox xItem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8175
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   825
      Width           =   2115
   End
   Begin VB.CommandButton cmdgo 
      BackColor       =   &H00C0FFFF&
      Caption         =   "«” Ã«»…"
      Height          =   390
      Left            =   375
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   1365
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   540
      ScaleWidth      =   11700
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   11700
      Begin VB.CommandButton CmdExit 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Œ—ÊÃ "
         Height          =   390
         Left            =   150
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   75
         Width           =   1365
      End
      Begin VB.CommandButton CmdCreat 
         BackColor       =   &H00C0FFFF&
         Caption         =   "÷»ÿ √—’œ… «·√’‰«›"
         Height          =   390
         Left            =   9120
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   120
         Width           =   2265
      End
   End
   Begin VB.TextBox LastOne 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   -525
      MaxLength       =   2
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1950
      Width           =   465
   End
   Begin MSFlexGridLib.MSFlexGrid invGrid 
      Height          =   6375
      Left            =   360
      TabIndex        =   0
      Top             =   2025
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   11245
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedCols       =   0
      BackColor       =   12648447
      BackColorFixed  =   32896
      ForeColorFixed  =   16777215
      BackColorSel    =   65535
      BackColorBkg    =   14737632
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
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      Height          =   1065
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   675
      Width           =   11550
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂÊœ «·’‰› :"
      Height          =   195
      Left            =   10650
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   900
      Width           =   810
   End
   Begin VB.Label xDesca 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6300
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1275
      Width           =   3990
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·’‰› :"
      Height          =   195
      Left            =   10650
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1350
      Width           =   915
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      Height          =   6615
      Left            =   225
      Top             =   1875
      Width           =   11520
   End
End
Attribute VB_Name = "itemMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim itemTable As Recordset, DocTable As Recordset
Sub Fillgrd()
If DocTable.RecordCount = 0 Then Exit Sub
InvGrid.Rows = 1
nPrevious = 0
DocTable.MoveFirst
I = 1
Do
   InvGrid.AddItem ""
   InvGrid.TextArray(faIndex(I, 0, InvGrid)) = TurnValue(DocTable.DESCA, Null, "")
   InvGrid.TextArray(faIndex(I, 1, InvGrid)) = IIf(IsDate(DocTable!Date), Format(DocTable!Date, "dd-mm-yyyy"), "")
   InvGrid.TextArray(faIndex(I, 2, InvGrid)) = Format(DocTable!In, "##0.00")
   InvGrid.TextArray(faIndex(I, 3, InvGrid)) = Format(DocTable.OUT, "##0.00")
   InvGrid.TextArray(faIndex(I, 4, InvGrid)) = Format(nPrevious + TurnValue(DocTable.In, Null, 0) - TurnValue(DocTable.OUT, Null, 0), "##0.00")
   InvGrid.TextArray(faIndex(I, 5, InvGrid)) = TurnValue(DocTable.DOC_ID, Null, "")
   nPrevious = nPrevious + TurnValue(DocTable.In, Null, 0) - TurnValue(DocTable.OUT, Null, 0)
   DocTable.MoveNext
   I = I + 1
Loop Until DocTable.EOF
End Sub
Sub myProc()
ActiveControl.Text = GrdText(Search.Grid1, 0)
Unload Search
validStr = ""
End Sub
Function MYVALID()
If xItem.Text = "" Then Exit Function
itemTable.FindFirst "ITEM = " & MyParn(xItem.Text)
If itemTable.NoMatch Then Exit Function
MYVALID = True
End Function
Private Sub CmdGo_Click()
If Not MYVALID Then Exit Sub
Set DocTable = mydb.OpenRecordset( _
               "Select * From file1_11 " & _
               " Where item = " & MyParn(xItem.Text) & _
                " Order by [Date],[in] desc ", dbOpenSnapshot)
Cmdgo.Enabled = False
Fillgrd
End Sub
Private Sub CmdCreat_Click()
Me.MousePointer = 11
mydb.Execute "DELETE * FROM FILE1_11"

cString = "INSERT INTO FILE1_11( " & _
           "[Date],Doc_Id,Code,[Type],item,Out,DescA,Price,Total,store)" & _
           " Select [Date],Doc_No,Code,'6'," & _
           "item,Quant," & _
           " '„»Ì⁄«  ›Ï ' & Format([Date], 'dd-mm-yy'), " & _
           " Price,PRICE * Quant ,Store" & _
           " From File6_20 " & _
           " WHERE FILE6_20.STORE <> 'SS'  "
mydb.Execute cString

cString = "INSERT INTO FILE1_11( " & _
           "[Date],Doc_Id,Code,[Type],item,[in],DescA,Price,Total,store)" & _
           " Select [Date],Doc_No,Code,'3'," & _
           " item,Quant," & _
           " '„—œÊœ „»Ì⁄«  ›Ï ' & Format([Date], 'dd-mm-yy'), " & _
           " Price,PRICE * Quant ,Store" & _
           " From File6_10 " & _
           " WHERE FILE6_10.STORE <> 'SS'  "
mydb.Execute cString

cString = "INSERT INTO FILE1_11( " & _
           "[Date],Doc_Id,Code,[Type],item,[in],DescA,Price,Total,store)" & _
           " Select [Date],Doc_No,Code,'2'," & _
           "item,Quant," & _
           " ' „‘ —Ì«  ›Ï ' & Format([Date], 'dd-mm-yy'), " & _
           " Price,PRICE * Quant ,Store" & _
           " From File7_20 " & _
           " WHERE FILE7_20.STORE <> 'SS'  "
mydb.Execute cString

cString = "INSERT INTO FILE1_11( " & _
           "[Date],Doc_Id,Code,[Type],item,[out],DescA,Price,Total,store)" & _
           " Select [Date],Doc_No,Code,'7'," & _
           "item,Quant," & _
           " '„—œÊœ „‘ —Ì«  ›Ï ' & Format([Date], 'dd-mm-yy'), " & _
           " Price,PRICE * Quant ,Store" & _
           " From File6_11 " & _
           " WHERE FILE6_11.STORE <> 'SS'  "
mydb.Execute cString


cString = "INSERT INTO FILE1_11( " & _
           "[Date],Doc_Id,[Type],item,[Out],DescA,STORE)" & _
           " Select [Date],Doc_No,'8'," & _
           "item,QUANT," & _
           " '’«œ— ›Ï' & Format([Date], 'dd-mm-yy'), " & _
           "STORE" & _
           " From file1_81 "
mydb.Execute cString

cString = "INSERT INTO FILE1_11( " & _
           "[Date],Doc_Id,[Type],item,[IN],DescA,STORE)" & _
           " Select [Date],Doc_No,'4'," & _
           "item,QUANT," & _
           " 'Ê«—œ ›Ï' & Format([Date], 'dd-mm-yy'), " & _
           " STORE" & _
           " From file1_80 "
mydb.Execute cString

cString = "INSERT INTO FILE1_11( " & _
           "[Date],Doc_Id,[Type],item,[Out],DescA,STORE)" & _
           " Select [Date],Doc_No,'9'," & _
           "item,QUANT," & _
           " 'Â«·ﬂ ›Ï' & Format([Date], 'dd-mm-yy'), " & _
           "STORE" & _
           " From file1_82 "
mydb.Execute cString

cString = "INSERT INTO FILE1_11( " & _
           "[Date],Doc_Id,[Type],item,[Out],DescA,STORE)" & _
           " Select [DATE],Doc_No,'F'," & _
           "item,QUANT," & _
           " ' ÕÊÌ·«  ≈·Ï' & STORE2, " & _
           "STORE1" & _
           " From file1_60 "
mydb.Execute cString

cString = "INSERT INTO FILE1_11( " & _
           "[Date],Doc_Id,[Type],item,[IN],DescA,STORE)" & _
           " Select [Date],Doc_No, 'T'," & _
           "item,QUANT," & _
           " ' ÕÊÌ·«  „‰' & STORE1, " & _
           "STORE2" & _
           " From file1_60 "
mydb.Execute cString


cString = "insert into File1_11(type,item,[date],store,desca,Doc_Id,[In])" & _
        " Select 'z',item,[date],store,' Ã—œ ›Ï ' & Format(Date,'dd-mm-yyyy'),Doc_NO,differ From File0_10 where Closed"
mydb.Execute cString
Me.MousePointer = 1
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
Set itemTable = mydb.OpenRecordset("file1_10", dbOpenDynaset)
Set movetable = mydb.OpenRecordset("File1_11", dbOpenDynaset)
'Me.Picture = LoadPicture(App.Path & "\graph\02-02.jpg")
InvGrid.FormatString = "»Ì«‰ «·Õ—ﬂ… |" & " «—ÌŒ «·„” ‰œ|" & "Ê«—œ|" & "’«œ— |" & "—’Ìœ|" & "„” ‰œ|"
InvGrid.Cols = 7
InvGrid.ColWidth(0) = 2800
InvGrid.ColWidth(1) = 1100
InvGrid.ColWidth(2) = 1100
InvGrid.ColWidth(3) = 1100
InvGrid.ColWidth(4) = 1100
InvGrid.ColWidth(5) = 1300
For I = 0 To InvGrid.Cols - 1
    InvGrid.ColAlignment(I) = 1
Next
End Sub
Private Sub xItem_Change()
itemTable.FindFirst "ITEM = " & MyParn(xItem.Text)
xDesca.Caption = IIf(itemTable.NoMatch, "", itemTable.DESCA)
Cmdgo.Enabled = Not itemTable.NoMatch
End Sub
Private Sub xItem_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then ItemsLookup
End Sub
Sub ItemsLookup()
    ActiveControl.Text = ""
    Dim Generalarray(4)
    Dim GrdArray(2)
        
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Item as «·’‰›,DescA as [«”„ «·’‰›] From file1_10 "
    Generalarray(3) = " Where DescA Like('*cFilter*')"
    Generalarray(4) = "Order by Item"
    GrdArray(1) = 1000
    GrdArray(2) = 4500
    
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End Sub
Private Sub xItem_LostFocus()
If TypeOf ActiveControl Is CommandButton Then Exit Sub
itemTable.FindFirst "ITEM = " & MyParn(xItem.Text)
If itemTable.NoMatch Then
    xDesca.Caption = ""
    Exit Sub
End If
xDesca.Caption = itemTable.DESCA
End Sub
