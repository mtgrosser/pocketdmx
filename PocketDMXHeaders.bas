Attribute VB_Name = "PocketDMXHeaders"
'*****************************************************************
' PocketDMX
' DMX Controller for GLP Pocket Scan and FreeDMX.com interface
' Created 2006-2008 by Matthias Grosser
' License: Creative Commons Non-Commercial Share-Alike 2.0
' Website: http://dev.brainkiller.org/pocketdmx/
'
'*****************************************************************

Option Explicit


Public Enum PocketScanDMXChannels
  chnPan
  chnTilt
  chnColor
  chnGobo
  chnShutter
  chnSpeed
  chnLaser
End Enum


Public Const VK_RBUTTON = &H2
Public Const VK_LBUTTON = &H1

Public Type POINTAPI
  X As Long
  Y As Long
End Type


Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (curCtr As Currency) As Boolean
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (curFreq As Currency) As Boolean
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Sub Out Lib "inpout32.dll" Alias "Out32" (ByVal iPortAddress As Integer, ByVal iValue As Integer)


Public Function FirstArg(ByRef rstrIn As String, Optional ByRef rstrDelimiter As String = "-", Optional ByVal lOffset As Long = 0, Optional ByVal bAllIfNone As Boolean = True) As String
  Dim lHash As Long

  lHash = InStr(lOffset + 1, rstrIn, rstrDelimiter)
  If lHash > 0 Then
    FirstArg = Mid$(rstrIn, lOffset + 1, lHash - lOffset - 1)
  Else
    If bAllIfNone Then FirstArg = Mid$(rstrIn, lOffset + 1)
  End If
End Function

Public Function ArgFrom(ByRef rstrIn As String, Optional ByRef rstrDelimiter As String = "-", Optional ByVal lOffset = 0, Optional ByVal bAllIfNone As Boolean = False) As String
  Dim lPos As Long
  
  lPos = InStr(lOffset + 1, rstrIn, rstrDelimiter)
  If lPos > 0 Then
    ArgFrom = Mid$(rstrIn, lPos + Len(rstrDelimiter))
  Else
    If bAllIfNone Then ArgFrom = (rstrIn)
  End If
End Function

