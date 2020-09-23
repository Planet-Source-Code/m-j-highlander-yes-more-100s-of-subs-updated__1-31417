Option Explicit
Const SW_SHOWNORMAL = 1
Declare Function GetDesktopWindow Lib "USER" () As Integer
Declare Function ShellExecute Lib "SHELL" (ByVal hwnd%, ByVal lpszOp$, ByVal lpszFile$, ByVal lpszParams$, ByVal lpszDir$, ByVal fsShowCmd%) As Integer


'Parameter   Description
'---------   -----------
'hwnd%       Identifies the parent window. This window receives any message
'            boxes an application produces (for example, for error
'            reporting).
'
'lpszOp$     Points to a null-terminated string specifying the operation to
'            perform. This string can be "open" or "print." If this
'            parameter is NULL, "open" is the default value.
'
'lpszFile$   Points to a null-terminated string specifying the file to open.
'
'lpszParams$ Points to a null-terminated string specifying parameters
'            passed to the application when the lpszFile parameter
'            specifies an executable file. If lpszFile points to a
'            string specifying a document file, this parameter is NULL.
'
'lpszDir$    Points to a null-terminated string specifying the default
'            directory.
'
'fsShowCmd% Specifies whether the application window is to be shown when
'            the application is opened. This parameter can be one of the
'            values described in the API ShowWindow().

Function OpenDoc (DocName As String) As Integer
Dim Scr_hDC As Integer

      Scr_hDC = GetDesktopWindow()
      OpenDoc = ShellExecute(Scr_hDC, "open", DocName, "", "C:\", SW_SHOWNORMAL)

End Function

