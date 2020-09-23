<div align="center">

## Shell Print Any Document


</div>

### Description

Easy code allows you to print any document on the computer using its default print handler. This is the same as if you right-click in the windows explorer and select Print. No command switches are needed. So simple I added a handy ShellExecute Error Handler.
 
### More Info
 
To make it a public function for use in a bas, I use the form's Hwnd as a parameter. This hwnd is only used to retrieve errors. So if you only care about pass/fail you could leave that part out. Just pass the function the hwnd and the path to the file and call.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jay Kreusch](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jay-kreusch.md)
**Level**          |Beginner
**User Rating**    |4.7 (28 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jay-kreusch-shell-print-any-document__1-10021/archive/master.zip)

### API Declarations

```
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
  (ByVal hwnd As Long, _
  ByVal lpszOp As String, _
  ByVal lpszFile As String, _
  ByVal lpszParams As String, _
  ByVal lpszDir As String, _
  ByVal FsShowCmd As Long) As Long
Private Const SE_ERR_FNF = 2&
Private Const SE_ERR_PNF = 3&
Private Const SE_ERR_ACCESSDENIED = 5&
Private Const SE_ERR_OOM = 8&
Private Const SE_ERR_DLLNOTFOUND = 32&
Private Const SE_ERR_SHARE = 26&
Private Const SE_ERR_ASSOCINCOMPLETE = 27&
Private Const SE_ERR_DDETIMEOUT = 28&
Private Const SE_ERR_DDEFAIL = 29&
Private Const SE_ERR_DDEBUSY = 30&
Private Const SE_ERR_NOASSOC = 31&
Private Const ERROR_BAD_FORMAT = 11&
```


### Source Code

```
Public Function ShellPrint(jFormHwnd As Long, FilePath As String) As String
  Dim Answer As Integer
  Dim Msg As String
  Answer = ShellExecute(jFormHwnd, "Print", FilePath, vbNullString, vbNullString, vbNormalFocus)
  If Answer <= 32 Then
    'There was an error
    Select Case Answer
      Case SE_ERR_FNF
        Msg = "File not found"
      Case SE_ERR_PNF
        Msg = "Path not found"
      Case SE_ERR_ACCESSDENIED
        Msg = "Access denied"
      Case SE_ERR_OOM
        Msg = "Out of memory"
      Case SE_ERR_DLLNOTFOUND
        Msg = "DLL not found"
      Case SE_ERR_SHARE
        Msg = "A sharing violation occurred"
      Case SE_ERR_ASSOCINCOMPLETE
        Msg = "Incomplete or invalid file association"
      Case SE_ERR_DDETIMEOUT
        Msg = "DDE Time out"
      Case SE_ERR_DDEFAIL
        Msg = "DDE transaction failed"
      Case SE_ERR_DDEBUSY
        Msg = "DDE busy"
      Case SE_ERR_NOASSOC
        Msg = "No association for file extension"
      Case ERROR_BAD_FORMAT
        Msg = "Invalid EXE file or error in EXE image"
      Case Else
        Msg = "Unknown error"
    End Select
  End If
  ShellPrint = Msg
End Function
Private Sub Command1_Click()
  Dim x As String
  x = ShellPrint(Me.hwnd, "C:\Bad File")
  If x <> vbNullString Then
    MsgBox x
  End If
End Sub
```

