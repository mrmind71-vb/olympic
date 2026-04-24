VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form CopyItem 
   Caption         =   "ÇáÈíÇäÇÊ"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   3300
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "äÞá ÇáãáÝ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   45
      TabIndex        =   4
      Top             =   1485
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ÖÈØ ÇáÊÓÚíÑ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   45
      TabIndex        =   3
      Top             =   765
      Width           =   5775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ÎÑæÌ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   45
      TabIndex        =   1
      Top             =   2205
      Width           =   5775
   End
   Begin VB.CommandButton cmdCompress 
      Caption         =   "ÊÌåíÒ ÈíÇäÇÊ ÇáÃÕäÇÝ ááãÍá"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   5775
   End
   Begin MSComctlLib.ProgressBar prog1 
      Height          =   375
      Left            =   45
      TabIndex        =   5
      Top             =   2835
      Visible         =   0   'False
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
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
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   900
      Width           =   2115
   End
End
Attribute VB_Name = "CopyItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aDir(2) As String, cFilezip As String, cDir As String, cFileData As String
Dim fs As New FileSystemObject, cPathTemp As String
Dim cPathShop2 As String, cCodeShop2 As String
Dim con As New adodb.Connection
Dim conmdb As adodb.Connection
Private Sub cmdCompress_Click()
TransData
End Sub
Private Sub Command1_Click()
Set conmdb = New adodb.Connection
Dim loctable As New adodb.Recordset
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & cPathTemp

loctable.Open "SELECT * FROM FILE1_10 ORDER BY ITEM ", conmdb, adOpenKeyset, adLockOptimistic, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double

With loctable
    conmdb.BeginTrans
    Do Until .EOF
        
        i = i + 1
        nValue = Round(i / (nRecordCount), 2) * 100
        prog1.Value = IIf(nValue > 100, 100, nValue)
    
        If Val(!price & "") <> Val(!PRICE2 & "") Then
            conmdb.Execute "update file1_10 set file1_10.price2 = " & myNear(Format(!price * 1.5, "#0.00"), 0.5) & _
            "  where file1_10.item = " & MyParn(loctable!Item)
        End If
        .MoveNext
    Loop
    prog1.Visible = False
    conmdb.CommitTrans
End With
closeCon conmdb
MsgBox "Êã ÖÈØ ÇáÊÓÚíÑ ÈäÌÇÍ"
Exit Sub
myerror:
prog1.Visible = False
MsgBox Err.Description
Err.Clear
conmdb.RollbackTrans
closeCon conmdb
End Sub
Private Sub Command2_Click()
    End
End Sub
Private Sub Command3_Click()
On Error GoTo myerror
fs.CopyFile cPathTemp, cPathShop2 & "\datashop.mdb"
MsgBox "Êã äÞá ÇáãáÝ ÈäÌÇÍ"
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Form_Load()
openCon con
cPathTemp = App.Path & "\MDB\datashop.mdb"
End Sub
Private Sub TransData()
Set conmdb = New adodb.Connection
If Not fs.FolderExists(App.Path & "\MDB") Then MyCreateFolder App.Path & "\mdb"
fs.CopyFile App.Path & "\mdb\databnk.mdb", cPathTemp
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & cPathTemp

cPathShop2 = GetDesca("select path from path")
cCodeShop2 = GetDesca("select code from path")


conmdb.BeginTrans

Dim cCaption As String
cCaption = Me.Caption
Me.Caption = "ÇÖÇÝÉ ãáÝ ÇáÇÕäÇÝ - ÎØæÉ 1 ãä 8"
CopyItems
Me.Caption = "ÇÖÇÝÉ ãáÝ ãÌãæÚÇÊ ÇáÇÕäÇÝ - ÎØæÉ 2 ãä 8"
CopyGroup
Me.Caption = "ÇÖÇÝÉ ãáÝ ãÌãæÚÇÊ ÇáÇÕäÇÝ ÇáÑÆíÓíÉ- ÎØæÉ 3 ãä 8"
CopyFlag "file1_50G"
Me.Caption = "ÇÖÇÝÉ ãáÝ  ÇáÇÞÓÇã - ÎØæÉ 4 ãä 8"
CopyFlag "file1_10SC"
Me.Caption = "ÇÖÇÝÉ ãáÝ  ÇáãÔÊÑíÇÊ ÇáÑÆíÓí - ÎØæÉ 5 ãä 8"
CopyPurchaseHeader 1
Me.Caption = "ÇÖÇÝÉ ãáÝ  ãÑÏæÏ ÇáãÔÊÑíÇÊ ÇáÑÆíÓí - ÎØæÉ 6 ãä 8"
CopyPurchaseHeader 2
Me.Caption = "ÇÖÇÝÉ ãáÝ   ÇáãÔÊÑíÇÊ ÇáÝÑÚí - ÎØæÉ 7 ãä 8"
CopyPurchaseSub 1
Me.Caption = "ÇÖÇÝÉ ãáÝ  ãÑÏæÏ ÇáãÔÊÑíÇÊ ÇáÝÑÚí - ÎØæÉ 8 ãä 8"
CopyPurchaseSub 2
conmdb.CommitTrans
conmdb.Close
Set conmdb = Nothing
MsgBox "Êã ÊÑÍíá ÇáÈíÇäÇÊ "
Exit Sub
myerror:
MsgBox Err.Description
MsgBox "áã ÊÊã ÚãáíÉ ÊÍÏíË ÇáÈíÇäÇÊ ÈÑÌÇÁ ÇáãÍÇæáÉ ãÑÉ ÃÎÑì"
conmdb.RollbackTrans
closeCon conmdb
Err.Clear
prog1.Visible = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
    End
End Sub
Private Sub CopyItems()
Dim loctable As New adodb.Recordset
cString = "SELECT FILE1_10.ITEM, FILE1_10.DESCA, FILE1_10.[GROUP], FILE1_10.PRICE, " & _
          " FILE1_10.PRICE2, FILE1_10.COST , FILE1_10.[SECTION], FILE1_10.PACKAGE, FILE1_10.UNIT , FILE1_10.DISCOUNT FROM (FILE1_10 INNER JOIN file6_20 ON FILE1_10.ITEM = file6_20.ITEM) INNER JOIN FILE6_20H ON file6_20.DOC_NO = FILE6_20H.DOC_NO GROUP BY FILE6_20H.code, FILE1_10.ITEM, FILE1_10.DESCA, file1_10.cost , FILE1_10.[GROUP], FILE1_10.PRICE,FILE1_10.PRICE2, FILE1_10.SECTION, FILE1_10.PACKAGE, FILE1_10.UNIT , FILE1_10.DISCOUNT HAVING file6_20H.code =  " & MyParn(cCodeShop2)
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double

Dim aInsert(10, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "ITEM"
    aInsert(0, 1) = addstring(loctable!Item)
    
    aInsert(1, 0) = "DESCA"
    aInsert(1, 1) = addstring(loctable!desca & "")
    
    aInsert(2, 0) = "[GROUP]"
    aInsert(2, 1) = addvalue(loctable!Group & "")
    
    aInsert(3, 0) = "PRICE"
    aInsert(3, 1) = Val(loctable!price & "")
    
    aInsert(4, 0) = "PRICE2"
    aInsert(4, 1) = Val(loctable!PRICE2 & "")
    
    aInsert(5, 0) = "COST0"
    aInsert(5, 1) = Val(loctable!cost & "")
    
    aInsert(6, 0) = "[SECTION]"
    aInsert(6, 1) = addvalue(loctable!Section & "")
    
    aInsert(7, 0) = "[UNIT]"
    aInsert(7, 1) = addvalue(loctable!UNIT)
            
    aInsert(8, 0) = "COST"
    aInsert(8, 1) = "0"
            
    aInsert(9, 0) = "SUPLER"
    aInsert(9, 1) = "NULL"
            
    aInsert(10, 0) = "MAXDISC"
    aInsert(10, 1) = "0"
            
    conmdb.Execute CreateInsert(aInsert, "FILE1_10")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub CopyPurchase()

End Sub
Private Sub CopyGroup()
Dim loctable As New adodb.Recordset
cString = "SELECT * FROM FILE1_50"
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double

Dim aInsert(2, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)

    aInsert(0, 0) = "CODE"
    aInsert(0, 1) = addvalue(loctable!CODE)
    
    aInsert(1, 0) = "DESCA"
    aInsert(1, 1) = addstring(loctable!desca & "")
    
    aInsert(2, 0) = "[GROUP]"
    aInsert(2, 1) = addvalue(loctable!Group & "")
    
    conmdb.Execute CreateInsert(aInsert, "FILE1_50")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub CopyPurchaseHeader(nFlag)
Dim loctable As New adodb.Recordset
If nFlag = 1 Then
    cString = "SELECT FILE6_20H.* FROM FILE6_20H WHERE FILE6_20H.CODE =  " & MyParn(cCodeShop2)
Else
    cString = "SELECT FILE6_10H.* FROM FILE6_10H WHERE FILE6_10H.CODE =  " & MyParn(cCodeShop2)
End If

loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double

Dim aInsert(6, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)

    aInsert(0, 0) = "Doc_No"
    aInsert(0, 1) = addstring(loctable!doc_no)
    
    aInsert(1, 0) = "code"
    aInsert(1, 1) = addstring(RetZero("1", 5))
    
    aInsert(2, 0) = "[Date]"
    aInsert(2, 1) = addDate(loctable!Date)
    
    aInsert(3, 0) = "store"
    aInsert(3, 1) = addstring(RetZero("1", 2))
    
    aInsert(4, 0) = "Discount"
    aInsert(4, 1) = Val(loctable!discount & "")
       
    aInsert(5, 0) = "Tax"
    aInsert(5, 1) = Val(loctable!tax & "")
    
    aInsert(6, 0) = "userName"
    aInsert(6, 1) = addstring(sUserName)
    
    If nFlag = 1 Then
        conmdb.Execute CreateInsert(aInsert, "FILE7_20H")
    Else
        conmdb.Execute CreateInsert(aInsert, "FILE7_30H")
    End If
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub CopyPurchaseSub(nFlag)
Dim loctable As New adodb.Recordset
If nFlag = 1 Then
    cString = "SELECT FILE6_20.*,FILE6_20H.CODE FROM FILE6_20 INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO WHERE FILE6_20H.CODE =  " & MyParn(cCodeShop2)
Else
    cString = "SELECT FILE6_10.*,FILE6_10H.CODE FROM FILE6_10 INNER JOIN FILE6_10H ON FILE6_10.DOC_NO = FILE6_10H.DOC_NO WHERE FILE6_10H.CODE =  " & MyParn(cCodeShop2)
End If
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True_
Dim i As Long, nValue As Double

Dim aInsert(8, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
        
        
    aInsert(0, 0) = "[Code]"
    aInsert(0, 1) = addstring(RetZero("1"))
    
    aInsert(1, 0) = "ITEM"
    aInsert(1, 1) = addstring(loctable!Item)
  
    aInsert(2, 0) = "store"
    aInsert(2, 1) = addstring(RetZero("1", 2))
           
    aInsert(3, 0) = "Price"
    aInsert(3, 1) = Val(loctable!price & "")
    
    aInsert(4, 0) = "Total"
    aInsert(4, 1) = Val(loctable!TOTAL & "")
            
    aInsert(5, 0) = "Quant"
    aInsert(5, 1) = Val(loctable!Quant & "")
        
    aInsert(6, 0) = "Discount"
    aInsert(6, 1) = Val(loctable!discount & "")
                     
    aInsert(7, 0) = "doc_no"
    aInsert(7, 1) = addstring(loctable!doc_no)
    
    aInsert(8, 0) = "Row"
    aInsert(8, 1) = Val(loctable!Row & "")
        
    If nFlag = 1 Then
        conmdb.Execute CreateInsert(aInsert, "FILE7_20")
    Else
        conmdb.Execute CreateInsert(aInsert, "FILE7_30")
    End If
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub CopyFlag(cFile)
Dim loctable As New adodb.Recordset
cString = "SELECT * FROM " & cFile
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

Dim aInsert(1, 1)

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double

Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)

    aInsert(0, 0) = "CODE"
    aInsert(0, 1) = addvalue(loctable!CODE)
    
    aInsert(1, 0) = "DESCA"
    aInsert(1, 1) = addstring(loctable!desca & "")
    
    conmdb.Execute CreateInsert(aInsert, cFile)
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub

