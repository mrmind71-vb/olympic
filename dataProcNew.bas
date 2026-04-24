Attribute VB_Name = "Module2"
Public Enum cmType
adStoredProc = adCmdStoredProc
adText = adCmdText
adTable = adCmdTable
End Enum

Public Enum tbMode
    tbFind = 1
    tbPrevious = 2
    tblast = 3
    tbNext = 4
    tbFirst = 5
End Enum

Public Enum tableMode
found = 1
NotFound = 2
NoRecords = 3
NotMatch = 3
End Enum
Public Function cmdNew(pString As String, Optional con As ADODB.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeOut As Integer = 1000) As ADODB.command
Set cmdNew = New ADODB.command
cmdNew.CommandTimeout = nTimeOut
cmdNew.ActiveConnection = con
cmdNew.CommandType = pType
cmdNew.CommandText = pString
If Not IsEmpty(aParam) Then
    Dim I As Long
    For I = 0 To UBound(aParam) Step 2
        cmdNew.Parameters("@" & aParam(I)).Value = aParam(I + 1)
    Next
End If
End Function

