Attribute VB_Name = "mdlEffect"
'*******************************************************************************
'** File Name  : mdlEffect.bas                                                **
'** Language   : Visual Basic 6.0                                             **
'** References : -                                                            **
'** Components : -                                                            **
'** Modules    : * mdlAPI (GetPixel and SetPixel)                             **
'**              * frmPaint (AdjustPaintResizeBox, DrawSelectionRect,         **
'**                          Form_Resize                                      **
'**              * mdlGeneral (ShowErrMessage)                                **
'** Developer  : Theo Zacharias (theo_yz@yahoo.com)                           **
'** Description: A modul to handle image effect operations                    **
'** Last modified on August 15, 2003                                          **
'*******************************************************************************

Option Explicit

Public Enum enmEffect
  conEffFlipHorizontal = 0
  conEffFlipVertical = 1
  conEffResize = 2
  conEffRotate = 3
  conEffInvertColors = 4
End Enum

Public Const conZoomFactor = 1.25
Public Const conMaxImageWidth = 50000
Public Const conMaxImageHeight = 50000

'Properties for resize effect
Public sngResizeWidth As Single                             'resize width factor
Public sngResizeHeight As Single                           'resize height factor

'Properties for rotate effect
Public blnRotateClockWise As Boolean
Public sngRotateAngle As Single

' Purpose    : Apply effect intImageEffect to the selection (if any) or to the
'              paint area
' Assumption : These effect properties have been initiated:
'              - sngResizeWidth, sngResizeHeight (for resize effect)
'              - sngRotateAngle (degree type, for rotate effect)
' Effect     : As specified
' Inputs     : intEffect, pic, picTemp
' Returns    : pic (with effect applied)
Public Sub ApplyEffect(intEffect As enmEffect, _
                       ByRef pic As PictureBox, picTemp As PictureBox)
  Dim blnAutoSize As Boolean                     'to save picTemp.AutoSize value
  
  On Error GoTo ErrorHandler
  
  With pic
    blnAutoSize = picTemp.AutoSize
    picTemp.AutoSize = True
    picTemp.Width = .Width
    picTemp.Height = .Height
    picTemp.Picture = .Image
    .Picture = Nothing
    Select Case intEffect
      Case conEffFlipHorizontal
        .PaintPicture picTemp.Image, .ScaleWidth, 0, _
                      -.ScaleWidth, .ScaleHeight, , , , , vbSrcCopy
      Case conEffFlipVertical
        .PaintPicture picTemp.Image, 0, .ScaleHeight, _
                      .ScaleWidth, -.ScaleHeight, , , , , vbSrcCopy
      Case conEffInvertColors
        .PaintPicture picTemp.Image, 0, 0, _
                      .ScaleWidth, .ScaleHeight, , , , , vbSrcInvert
      Case conEffResize
        frmPaint.DrawSelectionRect
        .Visible = False
        .Width = .Width * sngResizeWidth
        .Height = .Height * sngResizeHeight
        .PaintPicture picTemp.Image, 0, 0, _
                      .ScaleWidth, .ScaleHeight, , , , , vbSrcCopy
        .Visible = True
        frmPaint.DrawSelectionRect
        frmPaint.AdjustPaintResizeBox
        frmPaint.Form_Resize
      Case conEffRotate
        If sngRotateAngle = 180 Then
          .PaintPicture picTemp.Image, .ScaleWidth, .ScaleHeight, _
                        -.ScaleWidth, -.ScaleHeight, , , , , vbSrcCopy
        Else
          ImageRotate picSource:=picTemp, picDestination:=pic, _
                      sngRotateAngle:=sngRotateAngle, _
                      blnClockWise:=blnRotateClockWise
        End If
    End Select
    picTemp.AutoSize = blnAutoSize
  End With
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Rotate image picSource sngRotateAngle degree and save the result
'              in picDestination
' Assumptions: -
' Effect     : As specified
' Inputs     : picSource, picDestination, sngRotateAngle
' Return     : picDestination
Private Sub ImageRotate(picSource As PictureBox, _
                        picDestination As PictureBox, _
                        sngRotateAngle As Single, blnClockWise As Boolean)
  Const conPi = 3.14159265358979
  
  Dim A As Single                                            'angle of R and dXd
  Dim intMaxXY As Single 'maximum width or height of picDestination
  Dim dXs As Long         'relative coordinate where the pixel color information
  Dim dYs As Long         '                     will be retrieved from picSource
  Dim dXd As Long         'relative coordinate where the pixel color information
  Dim dYd As Long         '                    will be written to picDestination
  Dim lngAdjustX As Long                 'to adjust the new pixel coordinates to
  Dim lngAdjustY As Long                 ' make sure the whole part of the image
                                         '  is shown (currently only for 90° and
                                         '                        270° rotation)
  Dim lngColor(3) As Long                              'pixel colors information
  Dim R As Integer                               'length of line (0,0)-(dXd,dYd)
  Dim Xs As Integer           'base coordinate where the pixel color information
  Dim Ys As Integer           '                 will be retrieved from picSource
  Dim Xd As Integer           'base coordinate where the pixel color information
  Dim Yd As Integer           '                will be written to picDestination
                              
  'On Error GoTo ErrorHandler
  
  If blnClockWise Then
    sngRotateAngle = 360 - sngRotateAngle
  End If
  Xs = picSource.ScaleWidth / 2
  Ys = picSource.ScaleHeight / 2
  Xd = picDestination.ScaleWidth / 2
  Yd = picDestination.ScaleHeight / 2
  intMaxXY = varIIf(picDestination.ScaleWidth > picDestination.ScaleHeight, _
                    picDestination.ScaleWidth / 2, _
                    picDestination.ScaleHeight / 2)
  If (sngRotateAngle = 90) Or (sngRotateAngle = 270) Then
    lngAdjustX = ((picDestination.ScaleHeight - _
                   picDestination.ScaleWidth) / 2) - 2
    lngAdjustY = ((picDestination.ScaleWidth - _
                   picDestination.ScaleHeight) / 2)
    frmPaint.DrawSelectionRect
    picDestination.Tag = CStr(picDestination.Width)
    picDestination.Width = picDestination.Height
    picDestination.Height = CLng(picDestination.Tag)
    With frmPaint
      .DrawSelectionRect
      .AdjustPaintResizeBox
      .Form_Resize
      .Refresh
    End With
  Else
    lngAdjustX = 0
    lngAdjustY = 0
  End If
  sngRotateAngle = sngRotateAngle * (conPi / 180)             'convert to radian
  'Write each pixels to picDestination with transformed coordinates
  '  to make rotation effect
  picDestination.DrawMode = vbCopyPen
  For dXd = 0 To intMaxXY
    For dYd = 0 To intMaxXY
      If dXd = 0 Then
        A = conPi / 2
      Else
        A = Atn(dYd / dXd)
      End If
      R = Sqr((dXd * dXd) + (dYd * dYd))
      dXs = R * Cos(A + sngRotateAngle)
      dYs = R * Sin(A + sngRotateAngle)
      'Get pixel colors information from picSource
      lngColor(0) = GetPixel(picSource.hDC, Xs + dXs, Ys + dYs)
      lngColor(1) = GetPixel(picSource.hDC, Xs - dXs, Ys - dYs)
      lngColor(2) = GetPixel(picSource.hDC, Xs + dYs, Ys - dXs)
      lngColor(3) = GetPixel(picSource.hDC, Xs - dYs, Ys + dXs)
      'Set pixel colors information to picDestination
      If lngColor(0) <> -1 Then
        SetPixel picDestination.hDC, Xd + dXd + lngAdjustX, _
                 Yd + dYd + lngAdjustY, lngColor(0)
      End If
      If lngColor(1) <> -1 Then
        SetPixel picDestination.hDC, Xd - dXd + lngAdjustX, _
                 Yd - dYd + lngAdjustY, lngColor(1)
      End If
      If lngColor(2) <> -1 Then
        SetPixel picDestination.hDC, Xd + dYd + lngAdjustX, _
                 Yd - dXd + lngAdjustY, lngColor(2)
      End If
      If lngColor(3) <> -1 Then
        SetPixel picDestination.hDC, Xd - dYd + lngAdjustX, _
                 Yd + dXd + lngAdjustY, lngColor(3)
      End If
    Next
    picDestination.Refresh
  Next
  picDestination.Refresh
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub


