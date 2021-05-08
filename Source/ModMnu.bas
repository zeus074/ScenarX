Attribute VB_Name = "ModMnu"
Public Declare Function CreatePopupMenu Lib "user32.dll" () As Long
Public Declare Function DestroyMenu Lib "user32.dll" (ByVal hMenu As Long) As Long
Public Type MENUITEMINFO
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
Public Const MIIM_STATE = &H1
Public Const MIIM_ID = &H2
Public Const MIIM_TYPE = &H10
Public Const MFT_SEPARATOR = &H800
Public Const MFT_STRING = &H0
Public Const MFS_DEFAULT = &H1000
Public Const MFS_ENABLED = &H0
Public Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal _
hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As _
MENUITEMINFO) As Long
Public Declare Function TrackPopupMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal _
uFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal _
hWnd As Long, ByVal prcRect As Long) As Long
Public Const TPM_LEFTALIGN = &H0
Public Const TPM_TOPALIGN = &H0
Public Const TPM_NONOTIFY = &H80
Public Const TPM_RETURNCMD = &H100
Public Const TPM_LEFTBUTTON = &H0
Public Type POINT_TYPE
X As Long
Y As Long
End Type
Public Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINT_TYPE) As Long

