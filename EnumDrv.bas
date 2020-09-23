Declare Function GetDriveType Lib "Kernel" (ByVal nDrive As Integer) As Integer
'for nDrive   A:=0, B:=1,......

Global Const DRIVE_FIXED = 3      'Hard Disk
Global Const DRIVE_REMOTE = 4     'Network Drive or CD-ROM
Global Const DRIVE_REMOVABLE = 2  'Floppy

Sub EnumDrives (DrvLst() As String)
ReDim DrvLst(1 To 26)
Dim iDrv As Integer
Dim iTmp As Integer
Dim index As Integer
index = 0

For iDrv = 0 To 25
    iTmp = GetDriveType(iDrv)
    If iTmp <> 0 Then
        index = index + 1
        DrvLst(index) = Chr$(iDrv + 97)
    End If
Next iDrv

ReDim Preserve DrvLst(1 To index)
End Sub

Function sGetDrvType (sDrv As String) As String
' Const DRIVE_FIXED = 3      'Hard Disk
' Const DRIVE_REMOTE = 4     'Network Drive or CD-ROM
' Const DRIVE_REMOVABLE = 2  'Floppy
Dim iResut As Integer, iDrv As Integer
iDrv = Asc(LCase(Left(sDrv, 1))) - 97
iResult = GetDriveType(iDrv)
Select Case iResult
    Case DRIVE_FIXED
        sGetDrvType = "hard"
    Case DRIVE_REMOTE
        sGetDrvType = "cdrom"
    Case DRIVE_REMOVABLE
        sGetDrvType = "floppy"
    Case Else
        sGetDrvType = "none"
End Select

End Function

