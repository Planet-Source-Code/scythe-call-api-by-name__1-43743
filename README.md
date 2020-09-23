<div align="center">

## Call Api by Name


</div>

### Description

Call an api by its name without Declare. Usefull for Sripts or if u dont know if the Machine is XP or 95...
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Scythe](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/scythe.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Libraries](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/libraries__1-49.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/scythe-call-api-by-name__1-43743/archive/master.zip)





### Source Code

```
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Sub Form_Load()
 Dim Libary As Long
 Dim PrcAdress As Long
 On Error GoTo NoApi
 'Load the Libary
 Libary = LoadLibrary("user32")
 'Find the procedure we want
 Procadress = GetProcAddress(Libary, "MessageBoxA")
 'Call the Api
 CallWindowProc Procadress, Me.hWnd, "My Message", "Api without Declare", &H0&
 'Unload the libary
 FreeLibrary Libary
NoApi:
End Sub
```

