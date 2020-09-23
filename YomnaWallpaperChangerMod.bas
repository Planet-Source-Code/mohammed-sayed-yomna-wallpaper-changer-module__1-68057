Attribute VB_Name = "YomnaWallpaperChangerMod"
'ÈÓã Çááå ÇáÑÍãä ÇáÑÍíã
' In the name of Allah, Most Gracious, Most Merciful.
'---------------------------------------------------------------------------------------
' This code is part of M(Mohammed)Sayed's code collection version 1
' Please use it for good :)
'---------------------------------------------------------------------------------------
' Module    : YomnaWallpaperChangerMod
' DateTime  : (Update 05/03/2007 05:20 : Add position support)
' Author    : Mohammed Sayed | http://www.freewebs.com/msayed
' Contact   : msayed2004@gmail.com
' Purpose   : Enable changing the Desktop wallpaper with the specified position
'             (Center, Stretch & Tile).
' Disclaimer: This source code provided as is with no guarantee of functionality or any unwanted results
'             use it at your own risk.
' Sup. OSs  : This code was tested on Win98 , XP pro (SP2) & 2003 server (SP1).
' License   : Free to use in exchange for adding credits to Mohammed Sayed & the KPD Team with your applications.
' Reports   : I will appreciate any bug reports , suggestions are welcome , send me e-mails at msayed2004@gmail.com
'---------------------------------------------------------------------------------------
' Yomna is my niece and I start my projects with here name.
' Thanks to Mohammed Ragab & Keith Stanier for teaching me programming.
'---------------------------------------------------------------------------------------

Option Explicit
Option Compare Text

Public Enum YWPosition
    Center
    Stretch
    Tile
End Enum

Private Const SPI_SETDESKWALLPAPER As Long = 20
Private Const SPIF_UPDATEINIFILE As Long = &H1
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Const KEYPath As String = "Control Panel\Desktop"
Private Const REG_SZ As Long = 1

Private Enum RegClasses
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
End Enum

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Private Sub SaveString(hKey As RegClasses, strPath As String, strValue As String, strData As String)
    Dim Ret As Long
    RegCreateKey hKey, strPath, Ret
    RegSetValueEx Ret, strValue, 0, REG_SZ, ByVal strData, Len(strData)
    RegCloseKey Ret
End Sub

Public Function SetWallpaper(ImagePath As String, Position As YWPosition) As Boolean
'---------------------------------------------------------------------------------------
'Parameters :
'   1-ImagePath : The path of the image that you want to set as a wallpaper
'                 Only bitmaps (*.bmp) are supported.
'   2-Postion   : Wallpaper position , can be (Center, Stretch & Tile).
'---------------------------------------------------------------------------------------
'Return values : True if succeeded and False if failed.
'---------------------------------------------------------------------------------------

Dim Ret As Long
If Right$(ImagePath, 3) <> "bmp" Then SetWallpaper = False: Exit Function

Select Case Position
    Case Center
        SaveString HKEY_CURRENT_USER, KEYPath, "TileWallpaper", 0
        SaveString HKEY_CURRENT_USER, KEYPath, "WallpaperStyle", 0
    Case Stretch
        SaveString HKEY_CURRENT_USER, KEYPath, "TileWallpaper", 0
        SaveString HKEY_CURRENT_USER, KEYPath, "WallpaperStyle", 2
    Case Tile
        SaveString HKEY_CURRENT_USER, KEYPath, "TileWallpaper", 1
        SaveString HKEY_CURRENT_USER, KEYPath, "WallpaperStyle", 0
End Select
Ret = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, ImagePath, SPIF_UPDATEINIFILE)
SetWallpaper = CBool(Ret)
End Function
