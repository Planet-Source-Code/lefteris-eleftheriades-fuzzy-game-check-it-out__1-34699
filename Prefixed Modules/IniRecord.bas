Attribute VB_Name = "IniRecord"
'''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''
'''''' All functions you need to create  ''''''
'''''' database table Add, Load, Remove  ''''''
'''''' any record or field.              ''''''
'''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'FIELD PROCESSING
Public Function GetIniRecordData$(Filename$, RecordName$, Field$)
    'Loads the data of a field in a record in a file
    Dim Ret As String, NC As Long
    Ret = String(255, 0)
    NC = GetPrivateProfileString(RecordName$, Field$, "Error Retriving Data", Ret, 255, Filename$)
    If NC <> 0 Then Ret = Left$(Ret, NC)
    GetIniRecordData$ = Ret
End Function

Public Sub SetIniRecordData(Filename$, RecordName$, Field$, Value$)
   'Saves the data of a field in a record in a file
   WritePrivateProfileString RecordName$, Field$, Value$, Filename$
End Sub

Public Sub RemoveField(Filename$, RecordName$, Field$)
   'Removes the data of a field in a record in a file
   SetIniRecordData Filename$, RecordName$, Field$, ""
End Sub

'END FIELD PROCESSING
'________________________________________
'RECORD PROCESSING
Public Sub GetIniRecordBeginningAndEndPos(ByVal Filename$, ByVal RecordName$, ByRef StartChar&, ByRef StopChar&)
    'Returns the possission of the first caracter of
    'the record you specify as well as the last one.
    Dim First$, Last$, Temp$, AllText$
    Dim SStart&, SStop&
    Open Filename$ For Input As #1
         While Not EOF(1)
           Line Input #1, Temp$
           AllText$ = AllText$ & Temp$ & vbCrLf
         Wend
    Close #1
    SStart = InStr(1, AllText$, RecordName$) - 1
    If Not InStr(SStart + 1, AllText$, "[") = 0 Then
       SStop = InStr(SStart + 1, AllText$, "[") - 3
    Else
       SStop = Len(AllText$)
    End If
    
    StartChar = SStart
    StopChar = SStop
End Sub

Public Function LoadINIRecordDataAsText(ByVal Filename$, ByVal RecordName$)
'Returns all record's data as source code
Dim SStart&, SStop&
Dim Temp$, AllText$
    Open Filename$ For Input As #1
         While Not EOF(1)
           Line Input #1, Temp$
           AllText$ = AllText$ & Temp$ & vbCrLf
         Wend
    Close #1
GetIniRecordBeginningAndEndPos Filename$, RecordName$, SStart&, SStop&
LoadINIRecordDataAsText = Mid(AllText$, SStart, SStop - SStart)
End Function

Public Function RemoveRecord(ByVal Filename$, ByVal RecordName$)
'Deletes a whole Record. If the record deleted wasn't
'The last, and you have the records Indeed, you have to
'resort them. To help with sorting use the "ChangeRecordName"
'function in combination with "GetRecordNameText"
Dim SStart&, SStop&
Dim Temp$, AllText$
Dim StaT$, StoT$
    'Load all the text of the file
    Open Filename$ For Input As #1
         While Not EOF(1)
           Line Input #1, Temp$
           AllText$ = AllText$ & Temp$ & vbCrLf
         Wend
    Close #1
    'Get the beginning and end of the record
    GetIniRecordBeginningAndEndPos Filename$, RecordName$, SStart&, SStop&
    'Save the file bypassing the record
    StaT$ = Mid(AllText$, 1, SStart& - 3)
    StoT$ = Mid(AllText$, SStop& + 1)
    
    DoEvents
    'MsgBox SStart& & ", " & SStop& & vbCrLf & StaT$ & StoT$
    
    Open Filename$ For Output As #1
           Print #1, StaT$ & StoT$
    Close #1
End Function

Function GetRecordNameText(ByVal Filename$, ByVal RecordNo&)
Dim Count&
    Open Filename$ For Input As #1
         While Not EOF(1)
           Line Input #1, Temp$
           If InStr(Temp$, "[") <> 0 Then
             'Note the result should be 1-*
             Count& = Count& + 1
             If RecordNo& = Count& Then
                GetRecordText = Mid(Temp$, 2, Len(Temp$) - 2)
             End If
           End If
         Wend
    Close #1
End Function

Function RenameRecord(ByVal Filename$, ByVal OldName$, ByVal NewName$)
Dim Count&
Dim SStart&, SStop&
Dim Temp$, AllText$
Dim StaT$, StoT$

    Open Filename$ For Input As #1
         While Not EOF(1)
           Line Input #1, Temp$
           AllText$ = AllText$ & Temp$ & vbCrLf
         Wend
    Close #1
    
    SStart = InStr(1, AllText$, OldName) - 1
    SStop = InStr(SStart + 1, AllText$, "]")
    
    StaT$ = Mid(AllText$, 1, SStart&)
    StoT$ = Mid(AllText$, SStop&)
    'MsgBox StaT$ & NewName & StoT$
    DoEvents
    Open Filename$ For Output As #1
           Print #1, StaT$ & NewName$ & StoT$
    Close #1
End Function


'END RECORD PROCESSING
