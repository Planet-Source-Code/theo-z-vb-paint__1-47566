Attribute VB_Name = "mdlAPI"
'*******************************************************************************
'** File Name  : mdlAPI.bas                                                   **
'** Language   : Visual Basic 6.0                                             **
'** References : -                                                            **
'** Components : -                                                            **
'** Modules    : -                                                            **
'** Developer  : Theo Zacharias (theo_yz@yahoo.com)                           **
'** Description: A modul to declare Windows API procedures, functions and     **
'**              types                                                        **
'** Last modified on August 14, 2003                                          **
'*******************************************************************************

Option Explicit

Public Type typPoint
  x As Long
  y As Long
End Type

' To fill an area
Public Declare Sub _
  ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, _
                            ByVal y As Long, ByVal crColor As Long, _
                            ByVal wFillType As Long)

' To retrieve color at the specified coordinates
' (this procedure equal with Point method in PictureBox object, only faster)
Public Declare Function _
  GetPixel Lib "gdi32" (ByVal hDC As Long, _
                        ByVal x As Long, ByVal y As Long) As Long

' To draw a bezier curve
Public Declare Sub _
  PolyBezier Lib "gdi32" (ByVal hDC As Long, _
                          lppt As typPoint, ByVal cPoints As Long)

' To draw a polygon
Public Declare Sub _
  Polygon Lib "gdi32" (ByVal hDC As Long, _
                       lpPoint As typPoint, ByVal nCount As Long)

' To draw a rounded rectangle
Public Declare Sub _
  RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, _
                         ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, _
                         ByVal X3 As Long, ByVal Y3 As Long)

' To set the pixel at the specified coordinates to the specified color
' (this function equal with PSet method in PictureBox object, only faster)
Public Declare Sub _
  SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, _
                        ByVal y As Long, ByVal crColor As Long)

' To perform an operation to specific file
' (in this program, this is used to send mail)
Public Declare Sub _
  ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, _
     ByVal lpFile As String, ByVal lpParameters As String, _
     ByVal lpDirectory As String, ByVal nShowCmd As Long)
