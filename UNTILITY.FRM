VERSION 5.00
Begin VB.Form Utility 
   Caption         =   "Œœ„«  «·„·›« "
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MouseIcon       =   "Untility.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   2160
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdCompact 
      Caption         =   "÷€ÿ ﬁ«⁄œ… «·»Ì«‰« "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   75
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1125
      Width           =   4440
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "ÿ»«⁄… ﬂ· „·›«  «·›Â—”…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   75
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   75
      Width           =   4440
   End
   Begin VB.CommandButton CmdReIndex 
      Caption         =   "«⁄«œ… «·›Â—”…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   75
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   4440
   End
   Begin VB.CommandButton CmdDeleteIndex 
      Caption         =   "Õ–› „·›«  «·›Â—”…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   75
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1650
      Width           =   4440
   End
End
Attribute VB_Name = "Utility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MdbPath As String
Dim PackUpPath As String
Dim FileName As String
Private Sub CmdDeleteIndex_Click()
Me.MousePointer = 11
DeleteAllIndexes
Me.MousePointer = 0
End Sub
Private Sub CmdReIndex_Click()
Me.MousePointer = 11
DeleteAllIndexes
mydb.Execute "CREATE INDEX NFILE0_10_1 ON  FILE0_10 (CODE,ITEM,CDATE)"

mydb.Execute "CREATE INDEX NFILE0_20_1 ON  FILE0_20 (DOC_NO,ITEM)"

mydb.Execute "CREATE INDEX NFILE0_30_1 ON  FILE0_30 (DOC_ID)"
mydb.Execute "CREATE INDEX NFILE0_30_2 ON  FILE0_30 (CODE,DOC_ID)"
mydb.Execute "CREATE INDEX NFILE0_30_3 ON  FILE0_30 (DOC_ID,CODE)"

mydb.Execute "CREATE INDEX NFILE0_40_1 ON  FILE0_40 (DOC_NO,ITEM)"

mydb.Execute "CREATE INDEX NFILE1_10_1 ON  FILE1_10 (ITEM)"

mydb.Execute "CREATE INDEX NFILE1_11_1 ON  FILE1_11 (ITEM,CTYPE,DOC_ID)"
mydb.Execute "CREATE INDEX NFILE1_11_2 ON  FILE1_11 (CODE)"
mydb.Execute "CREATE INDEX NFILE1_11_3 ON  FILE1_11 (ITEM,CNO,CDATE)"
mydb.Execute "CREATE INDEX NFILE1_11_4 ON  FILE1_11 (ITEM,CDATE)"

mydb.Execute "CREATE INDEX NFILE1_30_1 ON  FILE1_30 (ITEM,STORE)"

mydb.Execute "CREATE INDEX NFILE1_50_1 ON  FILE1_50 (CODE)"


mydb.Execute "CREATE INDEX NFILE1_60_1 ON  FILE1_60 (DOC_NO)"

mydb.Execute "CREATE INDEX NFILE1_70_1 ON  FILE1_70 (CODE)"
mydb.Execute "CREATE INDEX NFILE1_70_2 ON  FILE1_70 (FLAG, CODE)"

mydb.Execute "CREATE INDEX NFILE1_80_1 ON  FILE1_80 (DOC_NO)"
mydb.Execute "CREATE INDEX NFILE1_80_2 ON  FILE1_80 (DOC_NO,ITEM)"

mydb.Execute "CREATE INDEX NFILE1_81_1 ON  FILE1_81 (DOC_NO)"
mydb.Execute "CREATE INDEX NFILE1_80_2 ON  FILE1_81 (DOC_NO,ITEM)"

mydb.Execute "CREATE INDEX NFILE1_82_1 ON  FILE1_82 (DOC_NO)"
mydb.Execute "CREATE INDEX NFILE1_82_2 ON  FILE1_82 (DOC_NO,ITEM)"

mydb.Execute "CREATE INDEX NFILE3_10_1 ON  FILE3_10 (CODE)"

mydb.Execute "CREATE INDEX NFILE3_11_1 ON  FILE3_11 (DOC_ID,VTYPE,VNO)"
mydb.Execute "CREATE INDEX NFILE3_11_2 ON  FILE3_11 (VTYPE,CODE,DOC_ID,VNO)"
mydb.Execute "CREATE INDEX NFILE3_11_3 ON  FILE3_11 (SALMAN,VTYPE,VDATE)"
mydb.Execute "CREATE INDEX NFILE3_11_4 ON  FILE3_11 (CODE,VTYPE,VDATE)"
mydb.Execute "CREATE INDEX NFILE3_11_5 ON  FILE3_11 (CODE,COMP,VDATE)"

mydb.Execute "CREATE INDEX NFILE3_20_1 ON  FILE3_20 (COMP,CLIENT)"

mydb.Execute "CREATE INDEX NFILE4_10_1 ON  FILE4_10 (CODE)"

mydb.Execute "CREATE INDEX NFILE4_11_1 ON  FILE4_11 (VTYPE,CODE,DOC_ID,VNO)"

mydb.Execute "CREATE INDEX  BNK_ID ON  FILE5_10 (BNK_ID)"
mydb.Execute "CREATE INDEX  ID ON  FILE5_10 (ID)"

mydb.Execute "CREATE INDEX  NFILE5_11_1 ON  FILE5_11 (VTYPE,BANK,DOC_ID,CODE)"
mydb.Execute "CREATE INDEX  NFILE5_11_2 ON  FILE5_11 (BANK)"

mydb.Execute "CREATE INDEX  NFILE5_20_1 ON  FILE5_20 (SER_NO)"
mydb.Execute "CREATE INDEX  NFILE5_20_2 ON  FILE5_20 (CLOSED,SAL,DATE_3)"

mydb.Execute "CREATE INDEX  NFILE5_21_1 ON  FILE5_21 (CHQ_NO,INV_NO)"
mydb.Execute "CREATE INDEX  NFILE5_21_2 ON  FILE5_21 (INV_NO)"

mydb.Execute "CREATE INDEX  NFILE5_31_1 ON  FILE5_31 (DOC_NO,CODE,BANK)"

mydb.Execute "CREATE INDEX  NFILE5_32_1 ON  FILE5_32 (DOC_NO,CODE,BANK)"

mydb.Execute "CREATE INDEX  NFILE5_33_1 ON  FILE5_33 (DOC_NO,CODE)"

mydb.Execute "CREATE INDEX  NFILE6_10_1 ON  FILE6_10 (CODE)"
mydb.Execute "CREATE INDEX  NFILE6_10_2 ON  FILE6_10 (SAL,ITEMCODE,INVDATE)"

mydb.Execute "CREATE INDEX  NFILE6_20_1 ON  FILE6_20 (CODE)"
mydb.Execute "CREATE INDEX  NFILE6_20_2 ON  FILE6_20 (SAL,ITEMCODE,INVDATE)"
mydb.Execute "CREATE INDEX  NFILE6_20_3 ON  FILE6_20 (CODE,CLIENT)"

mydb.Execute "CREATE INDEX  NFILE7_20_1 ON  FILE7_20 (INV_NO,ITEM)"

mydb.Execute "CREATE INDEX  NFILE7_30_1 ON  FILE7_30 (INV_NO,ITEM)"

mydb.Execute "CREATE INDEX  NFILE8_10_1 ON  FILE8_10 (DOC_NO,VNO,CODE)"
mydb.Execute "CREATE INDEX  NFILE8_10_2 ON  FILE8_10 (CODE,VDATE)"
mydb.Execute "CREATE INDEX  NFILE8_10_3 ON  FILE8_10 (INV_NO)"

mydb.Execute "CREATE INDEX  NFILE8_20_1 ON  FILE8_20 (DOC_NO,VNO,CODE)"
mydb.Execute "CREATE INDEX  NFILE8_20_2 ON  FILE8_20 (CODE)"

mydb.Execute "CREATE INDEX  NFILE8_30_1 ON  FILE8_30 (DOC_NO,VNO,CODE)"
mydb.Execute "CREATE INDEX  NFILE8_30_2 ON  FILE8_30 (CODE)"

mydb.Execute "CREATE INDEX  NFILE8_40_1 ON  FILE8_40 (DOC_NO,VNO,CODE)"
mydb.Execute "CREATE INDEX  NFILE8_40_2 ON  FILE8_40 (CODE)"

mydb.Execute "CREATE INDEX  NFILE8_50_1 ON  FILE8_50 (DOC_NO,VNO)"
Me.MousePointer = 0
End Sub
Private Sub CmdPrint_Click()
Dim aTables()
Me.MousePointer = 11
nCount = mydb.TableDefs.Count
ReDim aTables(nCount)
Open PathName & "indexes.txt" For Output As #1
For i = 0 To nCount - 1
    If Not Mid(mydb.TableDefs(i).Name, 1, 4) = "MSys" Then
        myPrint mydb, mydb.TableDefs(i)
    End If
Next
Close #1
cFileName = PathName & "indexes.txt"
Me.MousePointer = 0
RetVal = Shell("C:\Windows\notePad.Exe " + cFileName, 1)
End Sub
Private Sub CmdCompact_Click()
Me.MousePointer = 11
mydb.Close
DBEngine.CompactDatabase MdbPath, TempPath, dbLangArabic
Kill MdbPath
FileCopy TempPath, MdbPath
Kill TempPath
Set mydb = OpenDatabase(MdbPath)
Me.MousePointer = 0
End Sub
Private Sub CmdUndo_Click()
Kill FileName
FileCopy PathName & "backup.mdb", FileName
End Sub
Private Sub DeleteAllIndexes()
Dim aTables()
nCount = mydb.TableDefs.Count
ReDim aTables(nCount)
For i = 0 To nCount - 1
    If Not Mid(mydb.TableDefs(i).Name, 1, 4) = "MSys" Then
        myDelete mydb, mydb.TableDefs(i)
    End If
Next
End Sub
Private Sub myDelete(pmydb As Database, ptabledef As TableDef)
On Error Resume Next
Dim aIndexes()
nCount = ptabledef.Indexes.Count
ReDim aIndexes(nCount)
For i = 0 To nCount - 1
    cString = "drop index " & ptabledef.Indexes(i).Name & " on " & ptabledef.Name
    pmydb.Execute cString
Next
End Sub
Private Sub myPrint(pmydb As Database, ptabledef As TableDef)
Dim myIndex As Index
On Error Resume Next
Print #1, "******" & " Table " & ptabledef.Name & " Indexes List "
Print #1, "==========================================================="
Dim aIndexes()
nCount = ptabledef.Indexes.Count
ReDim aIndexes(nCount)
For i = 0 To nCount - 1
    Set myIndex = ptabledef.Indexes(i)
    cString = "   " & ptabledef.Indexes(i).Name & " ( " & myIndex.Fields & " )"
    Print #1, cString
Next
Print #1, " "
Print #1, " "
End Sub
Private Sub Form_Load()
Set mydb = OpenDatabase(MdbPath)
End Sub
Private Sub Makebackup()
mydb.Close
If Not Dir(PathName & "backup.mdb") = "" Then Kill PathName & "backup.mdb"
FileCopy FileName, PathName & "packup.mdb"
Set mydb = OpenDatabase(FileName)
End Sub





