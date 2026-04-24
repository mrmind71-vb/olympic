VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form ItemCode 
   BackColor       =   &H00E0E0E0&
   Caption         =   "«ﬂÊ«œ «·«’‰«›"
   ClientHeight    =   1635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
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
   LinkTopic       =   "Form3"
   ScaleHeight     =   1635
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin MSDBCtls.DBCombo xGroup 
      Bindings        =   "Itemcode.frx":0000
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1575
      TabIndex        =   4
      Top             =   150
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.CommandButton CmdUndo 
      BackColor       =   &H00C0FFFF&
      Caption         =   " —«Ã⁄"
      Height          =   390
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1125
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
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1125
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton CmdApply 
      BackColor       =   &H00C0FFFF&
      Caption         =   "„Ê«›ﬁ"
      Height          =   390
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   675
      Width           =   1140
   End
   Begin VB.Label xDesca 
      BackColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   1575
      TabIndex        =   3
      Top             =   675
      Width           =   3465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Group"
      Height          =   195
      Left            =   4500
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   225
      Width           =   435
   End
End
Attribute VB_Name = "ItemCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim uTable As Recordset, itemTable As Recordset
Sub myProc()
    cString = GrdText(Search.Grid1, 0)
    xItem.Text = Mid(cString, 4, 3)
    Unload Search
End Sub
Function MYVALID()
If xGroup.BoundText = "" Then Exit Function
MYVALID = True
End Function
Private Sub CmdApply_Click()
Dim cString As String
If MYVALID Then
    If Len(xGroup.BoundText) = 2 Then
        itemTable.FindLast "Mid(item,1,2) =  " & MyParn(xGroup.BoundText)
        If itemTable.NoMatch Then
            cString = xGroup.BoundText & "-001"
        Else
            cString = xGroup.BoundText & "-" & IncRec(Mid(itemTable.Item, 4, 3))
        End If
    Else
        itemTable.FindLast "Mid(item,1,3) =  " & MyParn(xGroup.BoundText)
        If itemTable.NoMatch Then
            cString = xGroup.BoundText & "-001"
        Else
            cString = xGroup.BoundText & "-" & IncRec(Mid(itemTable.Item, 5, 3))
        End If
    End If
    items.xItem.Text = cString
    items.xGroup.BoundText = xGroup.BoundText
    items.xDescA.Text = xGroup.Text
    'items.xGrDesc.Caption = ""
    Unload ItemCode
End If
End Sub
Private Sub CmdUndo_Click()
Unload ItemCode
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If TypeOf ActiveControl Is TextBox Or DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
Set itemTable = mydb.OpenRecordset("Select * From File1_10 Order by item", dbOpenSnapshot)
Data1.DatabaseName = MdbPath
Data1.RecordSource = "Select * From FILE1_50"
xGroup.BoundColumn = "Code"
xGroup.ListField = "Desca"
Data1.Refresh
End Sub
Sub RefreshData()
itemTable.Requery
End Sub

'Private Sub xGroup_Click(Area As Integer)
'If xGroup.BoundText <> "" Then Data1.RecordSource = "Select * From FILE1_50 where mainGroup = " & MyParn(xGroup.BoundText)
'End Sub
Private Sub xGroup_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
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
