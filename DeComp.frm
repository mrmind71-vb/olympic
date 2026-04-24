VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{57B5A17C-6DC0-4776-8F51-5519BF1235A6}#32.0#0"; "rsp-zip-compress-s140.ocx"
Begin VB.Form DeComp 
   Caption         =   "«·»Ì«‰« "
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2745
   ScaleWidth      =   5370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "›ﬂ ÷€ÿ ﬁ«⁄œ… «·»Ì«‰« "
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
      Left            =   2475
      TabIndex        =   5
      Top             =   1575
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.Frame Frame1 
      Height          =   2190
      Left            =   2325
      TabIndex        =   2
      Top             =   0
      Width           =   2940
      Begin VB.ListBox List1 
         Height          =   1815
         Left            =   150
         TabIndex        =   3
         Top             =   225
         Width           =   2625
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Œ—ÊÃ"
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
      Left            =   150
      TabIndex        =   1
      Top             =   1725
      Width           =   2115
   End
   Begin VB.CommandButton cmdCompress 
      Caption         =   " ÕœÌÀ »Ì«‰«  «·√«’‰«›"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   150
      TabIndex        =   0
      Top             =   975
      Width           =   2115
   End
   Begin RSPZipCompress140a.RSPZip RSPZip1 
      Left            =   1575
      Top             =   375
      _ExtentX        =   979
      _ExtentY        =   979
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   150
      TabIndex        =   4
      Top             =   2250
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
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
      TabIndex        =   6
      Top             =   900
      Width           =   2115
   End
End
Attribute VB_Name = "DeComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aDir(2) As String, cFilezip As String, cDir As String, cFileData As String
Dim fs
Public MYDB2 As Database
Private Sub cmdCompress2_Click()
'cFileName = App.Path & "\backup\" & systemName & ".zip"
'If fs.FileExists(cFileName) Then fs.DeleteFile cFileName
'ProgressBar1.Visible = True
'doCompress App.Path & "\backup\", systemName
End Sub
Private Sub cmdCompress_Click()
myCompress
End Sub
Private Sub doCompress(pDir, pFileData, pDirzip, pFilezip)
    Dim Comando As String
    Dim Ret As Long
    Dim MyString As String
    
    Me.Caption = "Status : Working..."
    
   
    '////////////////////////////////////////////
     
       
        'this will delete the destination zip file if it exist
        'and create a new zip file
        'this command is dangerous since it can delete a zip file
        'use it with cautions , it is only required when you really want to
        'create a new zip files instead updating it
        'normally it will not be used , since updating a zip with a file
        'will have the same effect
        Comando = Comando & "<zip-compression-mode=create-new-zipfile>"
    
    '////////////////////////////////////////////
    
    'addition of the <compression-level> command
    'this command what will be the compression level used during the zip execution
    ' 0 to store only , 9 to the maximum compression
    Comando = Comando & "<compression-level=" & 6 & ">"
    
    '////////////////////////////////////////////
    
    MyString = pDir
    If Len(MyString) Then
        Comando = Comando & "<directory-with-the-files-to-compress=" & MyString & ">"
    Else
        Label3.Caption = "Status : Finished"
        MsgBox "The command directory-with-the-files-to-compress is an obligatory command "
        Exit Sub
    End If
    
    '////////////////////////////////////////////
    
    'addition of the <destination-directory> command
    'this command will define the destination directory to the zip file
    'this is where the zip file will be created or updated
    'this command is obligatory
    'if the destination directory don't exist , it will be created
    'if the path cannot be created , and error and error description will be returned
    MyString = pDirzip
    If Len(MyString) Then
        Comando = Comando & "<destination-directory=" & MyString & ">"
    Else
        Label3.Caption = "Status : Finished"
        MsgBox "The command destination-directory is an obligatory command "
        Exit Sub
    End If
    
    '////////////////////////////////////////////
    
    'addition of the <destination-zipfile> command
    'this is the destination zip file to be created or updated
    'this command is obligatory , since you always need a zip file
    'for any command available in the control
    'the command defines only a zip file name , and not path , since the
    'path is the destination-directory command
    'if the zip file dont exist , it will be created , if the file cannot be created an error will be returned
    MyString = pFilezip
    If Len(MyString) Then
        Comando = Comando & "<destination-zipfile=" & MyString & ">"
    Else
        Label3.Caption = "Status : Finished"
        MsgBox "The command destination-zipfile is an obligatory command "
        Exit Sub
    End If
    '////////////////////////////////////////////
    
    'addition of the <files-selection> command
    'this is the files that you want to compress
    'to compress all files , pass *.* as the argument
    'if you want to compress the file mydatabase.mdb , pass mydatabase.mdb as the argument
    'It only accept one selection , for many selections call the compression function for each required selection
    'Thus , if you want to add *.txt and *.mdb to the zip file , make two calls , one with *.txt and another with *.mdb using the same destination zip file
    MyString = pFileData
    If Len(MyString) Then
        Comando = Comando & "<files-selection=" & MyString & ">"
    Else
        Label3.Caption = "Status : Finished"
        MsgBox "The command files-selection is an obligatory command "
        Exit Sub
    End If
    
    '////////////////////////////////////////////
    'if everything is ok , just call the zip function
    
    'Text7.Text = Comando
    
    
     First = GetTickCount
    
    
    
     'with the version 1.2.0 and above , processor priority functions were added
    'the following code will select a processor level to execute the decompression
    
    'values to define the processor priority
    '1 = IDLE
    '2 = LOWEST
    '3 = BELOW_NORMAL
    '4 = NORMAL
    '5 = ABOVE_NORMAL
    '6 = HIGHEST
    '7 = TIME_CRITICAL
       

    
    
    'this is the function to compress , it has only one argument
    'the argument is created in such a way to explain exactly to the control what you want to the zip compression to do
    RSPZip1.RSPZipCompress (Comando)
End Sub
Private Sub Decompress(pDir, pFile)
    Dim Comando As String
    Dim Ret As Long
    Me.Caption = "Status: Working..."
        
   MyString = Trim(pDir & "\" & pFile)
   If Len(MyString) Then
        'The creation of the command is only a sequence of additions to the Comando string
        Comando = Comando & "<zipfile=" & MyString & ">"
    Else
        MsgBox "A zip file is required , with complete path"
        Label10.Caption = "Status: Finished"
        Me.Refresh
        Exit Sub
    End If
    
    'addition of the <files-selection> command
    'this command will define the files that you want to extract
    'if you want to extract all files , just select *.*
    'to extract a unique file , just select the filename like this mydatabase.mdb , it will extract only the file  mydatabase.mdb
    If Len(MyString) Then
        Comando = Comando & "<files-selection=" & "*.*" & ">"
    Else
        MsgBox "The selection of files to test or extract are required in any unzip command"
        Label10.Caption = "Status: Finished"
        Me.Refresh
        Exit Sub
    End If
    
  
    'addition of the <destination-path> command
    'this is the path where the selected files will be extracted
    'if the path don't exist , it will be created
    'MyString = pDir
    MyString = "E:"
    If Len(MyString) Then
        Comando = Comando & "<destination-path=" & MyString & ">"
    End If
    
    'addition of the <file-extraction-mode> command
    'the file extraction mode will define whether the files being extracted will overwrite the files in the destination
    'or don't overwrite or freshen the files
    Comando = Comando & "<file-extraction-mode=overwrite>"
      
   
'  Comando = Comando & "<test-zipfile>"

        
    'Now just pass the command string the the unzip function , the return information will be available after the Finish event is raised , the event returns error number and error description
    'The RSPZipUncompress don't returns nothing because the code execution occurs in a different thread , and there is no way to predict what will be the return value , this is why
    'we need to wait for the finish event to be raised in order to see what occurred with the information passed
    
        
    First = GetTickCount 'profile code
        
    'with the version 1.2.0 and above , processor priority functions were added
    'the following code will select a processor level to execute the decompression
    
    'values to define the processor priority
    '1 = IDLE
    '2 = LOWEST
    '3 = BELOW_NORMAL
    '4 = NORMAL
    '5 = ABOVE_NORMAL
    '6 = HIGHEST
    '7 = TIME_CRITICAL
    
    RSPZip1.RSPZipUncompress (Comando)
    
    Me.Refresh
    
    'As you can see , you don't more than a unique command in order to execute the code , you only need to learn about the small number of commands available to the control
    'in order to start using the control without possible problems
    'It is very easy to use , and possible ( we hope ) bug free
    
End Sub
Private Sub Command1_Click()
cFileName = MdbPath
Set fs = CreateObject("Scripting.FileSystemObject")
If fs.FileExists(cFileName) Then fs.DeleteFile cFileName
ProgressBar1.Visible = True
Decompress "E:\", "DATAPOST.zip"
End Sub
Private Sub Command2_Click()
Set compressFrm = Nothing
Unload Me
End Sub
Private Sub Form_Load()
'nLast = LastInStr(MdbPath, "\")
cDir = App.Path & "\DATA"
cFileData = App.Path & "\DATA\POSTDATA.MDB"
cFilezip = "POSTDATA.zip"
Set fs = CreateObject("Scripting.FileSystemObject")
aDir(0) = "C:"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'if the code is executing , just cancel when closing the form
RSPZip1.RSPZipCancel
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set mydb = OpenDatabase(MdbPath)
'MydbEdit.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=true;Data Source=" & MdbPath
End Sub

Private Sub RSPZip1_ErrorCode(ErrorNumber As Long, ErrorDescription As String)
    'this event will be raised in case of internal errors
    
    Static NumberOfErrors As Long
    NumberOfErrors = NumberOfErrors + 1
    List1.AddItem NumberOfErrors & " : " & ErrorNumber & " : " & ErrorDescription    'This event is raised in any exception
    Dim Ret As Long
    Ret = List1.ListCount
    List1.ListIndex = Ret - 1
End Sub
Private Sub RSPZip1_Finished(ReturnCode As Long, ReturnDescription As String)
    'This event is the way to know that the execution is finished , this occurs because the
    'code execution is running  in a different thread created internally by the decompressor
    'it will return also the return number of the compression
    'in case of errors the ReturnCode will be different of 0
    
    Me.Caption = "«·«‰ Â«¡ „‰ ÷€ÿ «·„·›« "
    Me.MousePointer = 0
    'MsgBox " „ ›ﬂ «·‰”Œ… »‰Ã«Õ"
    ProgressBar1.Visible = False
    'Me.Caption = "Ì „ ‰”Œ «·„·›« "
   ' MsgBox "Ì „ ⁄„· ‰”Œ… " & I & " „‰ " & UBound(aDir)
    fs.CopyFile aDir(0) & "/" & cFilezip, aDir(0) & "/" & cFilezip
    MsgBox " „ ›ﬂ «·‰”Œ… »‰Ã«Õ"
End Sub
Private Sub RSPZip1_Progress(Progress As Long)
    'this event will update a progress bar with the actual position of the compression execution
    ProgressBar1.Value = Progress
End Sub
Private Sub RSPZip1_Status(Value As Long)
    
    'this event will be raised in any change in the internal status of the compressor
    
    If (Value = 0) Then
        Me.Caption = "«·«‰ Â«¡ „‰ ÷€ÿ «·„·›« "
       End If
    
    If (Value = 1) Then
        Me.Caption = "«·»ÕÀ"
    End If
    
    If (Value = 2) Then
        Me.Caption = "÷€ÿ «·„·›«  ..."
    End If
End Sub
'RSP Software - Thu Jan 22 17:05:58 2004 - http://rspsoftware.clic3.net
Private Sub RSPZip1_Warning(WarningCode As Long, WarningDescription As String)
 Static NumberOfWarnings As Long
    NumberOfWarnings = NumberOfWarnings + 1
    List1.AddItem NumberOfWarnings & " : " & WarningCode & " : " & WarningDescription    'This event is raised in any exception
    Dim Ret As Long
    Ret = List1.ListCount
    List1.ListIndex = Ret - 1
End Sub
Function RetZero(cString, Optional nLen As Integer = 2)
If Len(cString) >= nLen Then
    RetZero = cString
    Exit Function
End If
nLen = nLen - Len(cString)
RetZero = String(nLen, "0") & cString
End Function
Private Sub myCompress()
    On Error Resume Next
    Dim cFileName As String
    Dim Mdb2Data As String
    Dim LF As Boolean
    Set fs = CreateObject("Scripting.FileSystemObject")
    If MsgBox(" ÕœÌÀ »Ì«‰«  «·√’‰«› ", vbOKCancel) = vbOK Then
        If fs.FileExists("D:\ITEMDATA.mdb") Then
            cFileName = "D:\ITEMDATA.mdb"
            Dim MdbData As String
            Set MYDB2 = OpenDatabase(cFileName)
            MdbData = App.Path & "\data\data.mdb"
        
                 
             
            mydb.Execute " DELETE * FROM FILE6_20 WHERE STORE2 = '1' "
        
            Me.Caption = "Ì „  ⁄œÌ· «·»Ì«‰«  «·√‰ " & "  ===> 0 "
            MYDB2.Execute " INSERT INTO FILE6_20 IN  " & MyParn(MdbData) & " SELECT FILE6_20.* FROM FILE6_20 "
        
            MsgBox " „  —ÕÌ· «·»Ì«‰«   1 "
            MYDB2.Close
        End If
        
        If fs.FileExists("G:\SHOP2\SALDATA.mdb") Then
            cFileName = "G:\SHOP2\SALDATA.mdb"
            
            Set MYDB2 = OpenDatabase(cFileName)
            MdbData = App.Path & "\data\data.mdb"
        
            mydb.Execute " DELETE * FROM FILE6_20 WHERE STORE2 = '2' "
        
            Me.Caption = "Ì „  ⁄œÌ· «·»Ì«‰«  «·√‰ " & "  ===> 0 "
            MYDB2.Execute " INSERT INTO FILE6_20 IN  " & MyParn(MdbData) & " SELECT FILE6_20.* FROM FILE6_20 "
        
            MsgBox " „  —ÕÌ· «·»Ì«‰«   2 "
            MYDB2.Close
        End If
        
        If fs.FileExists("G:\SHOP3\SALDATA.mdb") Then
            cFileName = "G:\SHOP3\SALDATA.mdb"
            
            Set MYDB2 = OpenDatabase(cFileName)
            MdbData = App.Path & "\data\data.mdb"
        
            mydb.Execute " DELETE * FROM FILE6_20 WHERE STORE2 = '3' "
        
            Me.Caption = "Ì „  ⁄œÌ· «·»Ì«‰«  «·√‰ " & "  ===> 0 "
            MYDB2.Execute " INSERT INTO FILE6_20 IN  " & MyParn(MdbData) & " SELECT FILE6_20.* FROM FILE6_20 "
        
            MsgBox " „  —ÕÌ· «·»Ì«‰«  3 "
            MYDB2.Close
        End If
        
        Me.MousePointer = 0
    End If
Me.MousePointer = 0
mydb.Close
Set mydb = Nothing
End
End Sub
