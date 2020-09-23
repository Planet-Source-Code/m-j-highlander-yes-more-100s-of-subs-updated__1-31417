Attribute VB_Name = "Module1"

'Arabic Right_To_Left Menus for 32 bit VB

Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Public Const MFT_RIGHTORDER = &H2000&
Public Const MIIM_TYPE = &H10&
Public Const MF_BYPOSITION = &H400&


Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Boolean
Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Boolean
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long


Public Sub ArabicMenu(frm As Form)
Dim hMenu As Long
Dim Result As Long

hMenu = GetMenu(frm.hwnd)
Result = MakeMenuRtoL(hMenu, 1)
DrawMenuBar frm.hwnd

End Sub

Public Sub USMenu(frm As Form)
Dim hMenu As Long
Dim Result As Long

hMenu = GetMenu(frm.hwnd)
Result = MakeMenuRtoL(hMenu, 0)
DrawMenuBar frm.hwnd

End Sub

Public Function MakeMenuRtoL(ByVal hMenu As Long, ByVal RtoL As Long) As Long
Dim b As Long
Dim MenuName As String
Dim MenuInfo As MENUITEMINFO

If hMenu = 0 Then MakeMenuRtoL = False: Exit Function
MenuName = String(50, 0)
b = GetMenuString(hMenu, 0, MenuName, 50, MF_BYPOSITION)
MenuInfo.cbSize = Len(MenuInfo)
MenuInfo.fMask = MIIM_TYPE
MenuInfo.dwTypeData = MenuName
cch = Len(MenuName)
b = GetMenuItemInfo(hMenu, 0, True, MenuInfo)
If b Then
    If RtoL Then
    MenuInfo.fType = MenuInfo.fType Or MFT_RIGHTORDER
    Else
    MenuInfo.fType = MenuInfo.fType And Not MFT_RIGHTORDER
    End If
b = SetMenuItemInfo(hMenu, 0, True, MenuInfo)
End If
MakeMenuRtoL = b
End Function


