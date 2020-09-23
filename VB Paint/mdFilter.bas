Attribute VB_Name = "mdlFilter"
'*******************************************************************************
'** File Name  : mdlFilter.bas                                                **
'** Language   : Visual Basic 6.0                                             **
'** References : -                                                            **
'** Components : -                                                            **
'** Modules    : * mdlAPI (GetPixel and SetPixel)                             **
'**              * frmPaint (UpdateStatusBar)                                 **
'**              * mdlGeneral (ShowErrMessage)                                **
'** Developer  : Theo Zacharias (theo_yz@yahoo.com)                           **
'** Description: A modul to handle filter operations                          **
'** Last modified on August 11, 2003                                          **
'*******************************************************************************

'Notes:
'* I define filtering a picture as operation to read every pixel in the picture,
'  specified new properties of the pixel (new location, new color, etc.) and
'  write it to the picture.
'* Each filter below has filter factor store in sngFilterFactor variable. There
'  are comments about this factor for each filter about what happen if you
'  increase or decrease the value and what is the minimum and maximum value of
'  this factor. Please note that I make the comments based on a very concise
'  experiment, not by analyze it with pencil and paper. So it may not be quite
'  accurate.

Option Explicit

Public Enum enmFilter
  conFltBlacknWhite = 0
  conFltBlur = 1
  conFltBrightness = 2
  conFltCrease = 3
  conFltDarkness = 4
  conFltDiffuse = 5
  conFltEmboss = 6
  conFltGrayBlacknWhite = 7
  conFltGrayscale = 8
  conFltInvertColors = 9
  conFltReplaceColors = 10
  conFltSharpen = 11
  conFltSnow = 12
  conFltWave = 13
End Enum

'Properties for "replace color" filter
Public lngReplacedColor As Long
Public lngReplaceWithColor As Long

' Purpose    : Apply filter intFilter to clip region (X1,Y1)-(X2,Y2) of picture
'              box pic (if clip region is omitted then the filter will be
'              applied to the whole picture)
' Assumptions: * These filter properties have been initiated:
'                  lngReplaceColor, lngReplaceWithcolor (for "replace color"
'                  filter)
'              * X2 > 0 and Y2 > 0
' Effects    : -
' Input      : intFilter, pic, X1, Y1, X2, Y2
' Return     : pic (with the filter applied)
Public Sub ApplyFilter(intFilter As enmFilter, ByRef pic As PictureBox, _
                       Optional X1 As Long = -1, Optional Y1 As Long = -1, _
                       Optional X2 As Long = -1, Optional Y2 As Long = -1)
  Dim blnSmallArea As Boolean            'Condition whether the filter operation
                                         '         only be applied to small area
  Dim intDrawMode As Integer                    'to keep current draw mode value
  Dim lngColor() As Long        'three dimensions array to save RGB color (first
                                '             dimension: R = 0, G = 1, B = 2) of
                                ' (X,Y) coordinate (second and third dimensions)
  Dim lngReadColor As Long                                 'current color readed
  Dim lngTransColor As Long                         'color transformation factor
  Dim lngWriteColor As Long                               'current color written
  Dim R As Long                                                     'current RGB
  Dim G As Long                                                     '      color
  Dim B As Long                                                     'information
  Dim sngFilterFactor As Single
  Dim X As Long                                              'current coordinate
  Dim Y As Long                                              '   pixel processed
  
  On Error GoTo ErrorHandler
  
  If (X1 = -1) And (Y1 = -1) And (X2 = -1) And (Y2 = -1) Then
    X1 = 0
    Y1 = 0
    X2 = pic.ScaleWidth
    Y2 = pic.ScaleHeight
  End If
  blnSmallArea = (((X2 - X1) * (Y2 - Y1)) < (16 * 16))
  With pic
    intDrawMode = .DrawMode
    .DrawMode = vbCopyPen
    Select Case intFilter
      Case conFltBlacknWhite
        sngFilterFactor = 192      'increase this value to get more black colors
                                   '     than white colors or decrease it to get
                                   '         more white colors than black colors
                                   '  0 for total white and 256 for total black)
        For X = X1 To X2
          For Y = Y1 To Y2
            lngReadColor = mdlAPI.GetPixel(hdc:=.hdc, X:=X, Y:=Y)
            R = lngReadColor Mod 256
            If (R >= sngFilterFactor) Then
              lngWriteColor = vbWhite
            Else
              G = (lngReadColor \ 256) Mod 256
              If (G >= sngFilterFactor) Then
                lngWriteColor = vbWhite
              Else
                B = (lngReadColor \ 256) \ 256
                If (B >= sngFilterFactor) Then
                  lngWriteColor = vbWhite
                Else
                  lngWriteColor = vbBlack
                End If
              End If
            End If
            mdlAPI.SetPixel hdc:=.hdc, X:=X, Y:=Y, crColor:=lngWriteColor
          Next
          If Not blnSmallArea Then
            pic.Refresh
            frmPaint.UpdateStatusBar intInfo:=conStFiltering, _
                                     intPercentage:=((X * 100) \ X2)
          End If
        Next
      Case conFltBlur
        sngFilterFactor = 10         'decrease this value to get more bright blur
                                     '       or increase it to get more dark blur
                                     '            (limit to 0 for total white and
                                     '                        256 for total black
        RetrieveColorInformation pic:=pic, lngColor:=lngColor, _
                                 X1:=X1, Y1:=Y1, X2:=X2, Y2:=Y2, _
                                 blnShowProgress:=(Not blnSmallArea)
        For X = X1 + 1 To X2 - 1
          For Y = Y1 + 1 To Y2 - 1
            R = lngColor(0, X - 1, Y - 1) + lngColor(0, X, Y - 1) + _
                lngColor(0, X + 1, Y - 1) + lngColor(0, X - 1, Y) + _
                lngColor(0, X, Y) + lngColor(0, X + 1, Y) + _
                lngColor(0, X - 1, Y + 1) + lngColor(0, X, Y + 1) + _
                lngColor(0, X + 1, Y + 1)
            G = lngColor(1, X - 1, Y - 1) + lngColor(1, X, Y - 1) + _
                lngColor(1, X + 1, Y - 1) + lngColor(1, X - 1, Y) + _
                lngColor(1, X, Y) + lngColor(1, X + 1, Y) + _
                lngColor(1, X - 1, Y + 1) + lngColor(1, X, Y + 1) + _
                lngColor(1, X + 1, Y + 1)
            B = lngColor(2, X - 1, Y - 1) + lngColor(2, X, Y - 1) + _
                lngColor(2, X + 1, Y - 1) + lngColor(2, X - 1, Y) + _
                lngColor(2, X, Y) + lngColor(2, X + 1, Y) + _
                lngColor(2, X - 1, Y + 1) + lngColor(2, X, Y + 1) + _
                lngColor(2, X + 1, Y + 1)
            lngWriteColor = RGB(Abs(R / sngFilterFactor), _
                                Abs(G / sngFilterFactor), _
                                Abs(B / sngFilterFactor))
            mdlAPI.SetPixel hdc:=.hdc, X:=X, Y:=Y, crColor:=lngWriteColor
          Next
          If Not blnSmallArea Then
            pic.Refresh
            frmPaint.UpdateStatusBar intInfo:=conStFiltering, _
                                    intPercentage:=(((X + 1) * 100) \ X2)
          End If
        Next
      Case conFltBrightness, conFltDarkness
        Select Case intFilter
          Case conFltBrightness
            If Not blnSmallArea Then
              sngFilterFactor = 32   'decrease this value to make more bright or
                                     '           increase it to make less bright
                                     '           (limit to 0 for total white and
                                     '                    256 for no brightness)
            Else
              sngFilterFactor = 2
            End If
          Case conFltDarkness
            If Not blnSmallArea Then
              sngFilterFactor = -32    'decrease this value to make more dark or
                                       '           increase it to make less dark
                                       '          (-256 for inverting colors and
                                       '               limit to for no darkness)
            Else
              sngFilterFactor = -2
            End If
        End Select
        For X = X1 To X2
          For Y = Y1 To Y2
            lngReadColor = mdlAPI.GetPixel(hdc:=.hdc, X:=X, Y:=Y)
            GetRGBColor lngColor:=lngReadColor, R:=R, G:=G, B:=B
            lngWriteColor = RGB(Abs(R + sngFilterFactor), _
                                Abs(G + sngFilterFactor), _
                                Abs(B + sngFilterFactor))
            mdlAPI.SetPixel hdc:=.hdc, X:=X, Y:=Y, crColor:=lngWriteColor
            
          Next
          If Not blnSmallArea Then
            pic.Refresh
            frmPaint.UpdateStatusBar intInfo:=conStFiltering, _
                                     intPercentage:=((X * 100) \ X2)
          End If
        Next
      Case conFltCrease, conFltWave
        Select Case intFilter
          Case conFltCrease
            sngFilterFactor = 512     'decrease this value to get more crease or
                                      '           increase it to get less crease
                                      '               (64 for maximum crease and
                                      '                     65536 for no crease)
          Case conFltWave
            sngFilterFactor = 4        'increase this value to get more wave or
                                       '            decrease it to get less wave
                                       ' (0 for no wave and 16 for maximum wave)
        End Select
        RetrieveColorInformation pic:=pic, lngColor:=lngColor, _
                                 X1:=X1, Y1:=Y1, X2:=X2, Y2:=Y2, blnAll:=True, _
                                 blnShowProgress:=(Not blnSmallArea)
        For X = X1 To X2
          For Y = Y1 To Y2
            lngWriteColor = lngColor(3, X, Y)
            mdlAPI.SetPixel hdc:=.hdc, X:=X, _
                            Y:=(Sin(X) * sngFilterFactor) + (Y), _
                            crColor:=lngWriteColor
          Next
          If Not blnSmallArea Then
            pic.Refresh
            frmPaint.UpdateStatusBar intInfo:=conStFiltering, _
                                    intPercentage:=((X * 100) \ X2)
          End If
        Next
      Case conFltDiffuse
        sngFilterFactor = 5
        RetrieveColorInformation pic:=pic, lngColor:=lngColor, _
                                 X1:=X1, Y1:=Y1, X2:=X2, Y2:=Y2, blnAll:=True, _
                                 blnShowProgress:=(Not blnSmallArea)
        For X = X1 + 2 To X2 - 3
          For Y = Y1 + 2 To Y2 - 3
            lngReadColor = lngColor(3, X, Y + Int((Rnd * sngFilterFactor) - 2))
            R = Abs(lngReadColor Mod 256)
            lngReadColor = lngColor(3, X + Int((Rnd * sngFilterFactor) - 2), Y)
            G = Abs((lngReadColor \ 256) Mod 256)
            lngReadColor = lngColor(3, X + Int((Rnd * sngFilterFactor) - 2), _
                                       Y + Int((Rnd * sngFilterFactor) - 2))
            B = Abs((lngReadColor \ 256) \ 256)
            lngWriteColor = RGB(Red:=R, Green:=G, Blue:=B)
            mdlAPI.SetPixel hdc:=.hdc, X:=X, Y:=Y, crColor:=lngWriteColor
          Next
          If Not blnSmallArea Then
            pic.Refresh
            frmPaint.UpdateStatusBar intInfo:=conStFiltering, _
                                     intPercentage:=(((X + 3) * 100) \ X2)
          End If
        Next
      Case conFltEmboss
        sngFilterFactor = -128      'increase this abs(value) to get more bright
                                    ' emboss decrease it to get more dark emboss
                                    '             (0 for maximum dark emboss and
                                    '              256 for maximum bright emboss
        RetrieveColorInformation pic:=pic, lngColor:=lngColor, _
                                 X1:=X1, Y1:=Y1, X2:=X2, Y2:=Y2, _
                                 blnShowProgress:=(Not blnSmallArea)
        For X = X1 To X2 - 1
          For Y = Y1 To Y2 - 1
            R = Abs(lngColor(0, X, Y) - lngColor(0, X + 1, Y + 1) + _
                    sngFilterFactor)
            G = Abs(lngColor(1, X, Y) - lngColor(1, X + 1, Y + 1) + _
                    sngFilterFactor)
            B = Abs(lngColor(2, X, Y) - lngColor(2, X + 1, Y + 1) + _
                    sngFilterFactor)
            lngWriteColor = RGB(Red:=R, Green:=G, Blue:=B)
            mdlAPI.SetPixel hdc:=.hdc, X:=X, Y:=Y, crColor:=lngWriteColor
          Next
          If Not blnSmallArea Then
            pic.Refresh
            frmPaint.UpdateStatusBar intInfo:=conStFiltering, _
                                     intPercentage:=(((X + 1) * 100) \ X2)
          End If
        Next
      Case conFltGrayBlacknWhite
        sngFilterFactor = 3         'increase this value to get more black colors
                                    '      or decrase it to get more white colors
                                    '                 (limit to 0 for total white
                                    '                     and 32 for total black)
        For X = X1 To X2
          For Y = Y1 To Y2
            lngReadColor = mdlAPI.GetPixel(hdc:=.hdc, X:=X, Y:=Y)
            GetRGBColor lngColor:=lngReadColor, R:=R, G:=G, B:=B
            R = Abs(R * (G - B + G + R)) / 256
            G = Abs(R * (B - G + B + R)) / 256
            B = Abs(G * (B - G + B + R)) / 256
            lngReadColor = RGB(Red:=R, Green:=G, Blue:=B)
            GetRGBColor lngColor:=lngReadColor, R:=R, G:=G, B:=B
            lngReadColor = (R + G + B) / sngFilterFactor
            lngWriteColor = RGB(Red:=lngReadColor, _
                                Green:=lngReadColor, Blue:=lngReadColor)
            mdlAPI.SetPixel hdc:=.hdc, X:=X, Y:=Y, crColor:=lngWriteColor
          Next
          If Not blnSmallArea Then
            pic.Refresh
            frmPaint.UpdateStatusBar intInfo:=conStFiltering, _
                                    intPercentage:=((X * 100) \ X2)
          End If
        Next
      Case conFltGrayscale
        sngFilterFactor = 0.32 'increase this value to get more bright grayscale
                               '       or decrease it to get more dark grayscale
                               '                (0 for total black and (256 / 6)
                               '                          for almost total white
        For X = X1 To X2
          For Y = Y1 To Y2
            lngReadColor = mdlAPI.GetPixel(hdc:=.hdc, X:=X, Y:=Y)
            GetRGBColor lngColor:=lngReadColor, R:=R, G:=G, B:=B
            lngTransColor = Abs((R * sngFilterFactor) + _
                                (G * sngFilterFactor) + (B * sngFilterFactor))
            lngWriteColor = RGB(Red:=lngTransColor, _
                                Green:=lngTransColor, Blue:=lngTransColor)
            mdlAPI.SetPixel hdc:=.hdc, X:=X, Y:=Y, crColor:=lngWriteColor
          Next
          If Not blnSmallArea Then
            pic.Refresh
            frmPaint.UpdateStatusBar intInfo:=conStFiltering, _
                            intPercentage:=((X * 100) \ X2)
          End If
        Next
      Case conFltReplaceColors
        For X = X1 To X2
          For Y = Y1 To Y2
            lngReadColor = mdlAPI.GetPixel(hdc:=.hdc, X:=X, Y:=Y)
            If lngReadColor = lngReplacedColor Then
              lngWriteColor = lngReplaceWithColor
              mdlAPI.SetPixel hdc:=.hdc, X:=X, Y:=Y, crColor:=lngWriteColor
            End If
          Next
          If Not blnSmallArea Then
            pic.Refresh
            frmPaint.UpdateStatusBar intInfo:=conStFiltering, _
                                     intPercentage:=((X * 100) \ X2)
          End If
        Next
      Case conFltSharpen, conFltSnow
        Select Case intFilter
          Case conFltSharpen
            sngFilterFactor = 0.5        'increase this value to get more sharp
                                         '     or decrease it to get less sharp
                                         '                (0 for no sharpen and
                                         '               2 for maximum sharpen)
          Case conFltSnow
            sngFilterFactor = 24          'increase this value to get more snow
                                          '     or decrease it to get less snow
                                          '            (4 for minimum snowy and
                                          '               64 for maximum snowy)
        End Select
        RetrieveColorInformation pic:=pic, lngColor:=lngColor, _
                                 X1:=X1, Y1:=Y1, X2:=X2, Y2:=Y2, _
                                 blnShowProgress:=(Not blnSmallArea)
        For X = X1 + 1 To X2
          For Y = Y1 + 1 To Y2
            R = lngColor(0, X, Y) + _
                (sngFilterFactor * _
                 (lngColor(0, X, Y) - lngColor(0, X - 1, Y - 1)))
            G = lngColor(1, X, Y) + _
                (sngFilterFactor * _
                 (lngColor(1, X, Y) - lngColor(1, X - 1, Y - 1)))
            B = lngColor(2, X, Y) + _
                (sngFilterFactor * _
                 (lngColor(2, X, Y) - lngColor(2, X - 1, Y - 1)))
            lngWriteColor = RGB(Abs(R), Abs(G), Abs(B))
            mdlAPI.SetPixel hdc:=.hdc, X:=X, Y:=Y, crColor:=lngWriteColor
          Next
          If Not blnSmallArea Then
            pic.Refresh
            frmPaint.UpdateStatusBar intInfo:=conStFiltering, _
                                     intPercentage:=((X * 100) \ X2)
          End If
        Next
    End Select
    .DrawMode = intDrawMode
    .Refresh
  End With
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Get each R (red), G (green), B (blue) information color from RGB
'              color lngColor
' Assumptions: -
' Effects    : -
' Inputs     : lngColor
' Return     : R, G, B
Private Sub GetRGBColor(lngColor As Long, ByRef R As Long, _
                        ByRef G As Long, ByRef B As Long)
  On Error GoTo ErrorHandler
  
  R = lngColor Mod 256
  G = (lngColor \ 256) Mod 256
  B = (lngColor \ 256) \ 256
  Exit Sub
  
ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Retrieve every pixels color information in region (X1,Y1)-(X2,Y2)
'              of picture box pic and save the result to lngColor()
' Assumptions: -
' Effects    : -
' Input      : * pic
'              * X1, Y1, X2, Y2
'              * blnAll (condition whether to retrieve all color in once
'                        or seperate it in Red, Green and Blue color information
'              * blnShowProgress (condition whether it needs to refresh for
'                                 every column filtered)
' Return     : lngColor() (three dimensions array to save RGB color (first
'                          dimension: R = 0, G = 1, B = 2, All = 3) of (X,Y)
'                          coordinate (second and third dimensions))
Private Sub RetrieveColorInformation( _
              pic As PictureBox, ByRef lngColor() As Long, _
              Optional X1 As Long = -1, Optional Y1 As Long = -1, _
              Optional X2 As Long = -1, Optional Y2 As Long = -1, _
              Optional blnAll As Boolean = False, _
              Optional blnShowProgress = True _
            )
  Dim R As Long                                                     'current RGB
  Dim G As Long                                                     '      color
  Dim B As Long                                                     'information
  Dim X As Long                                              'current coordinate
  Dim Y As Long                                              '   pixel processed
  
  On Error GoTo ErrorHandler
  
  If (X1 = -1) Or (Y1 = -1) Or (X2 = -1) Or (Y2 = -1) Then
    X1 = 0
    Y1 = 0
    X2 = pic.ScaleWidth
    Y2 = pic.ScaleHeight
  End If
  If blnAll Then
    ReDim lngColor(3, X2, Y2)
  Else
    ReDim lngColor(2, X2, Y2)
  End If
  For X = X1 To X2
    For Y = Y1 To Y2
      If blnAll Then
        lngColor(3, X, Y) = mdlAPI.GetPixel(pic.hdc, X, Y)
      Else
        GetRGBColor lngColor:=mdlAPI.GetPixel(pic.hdc, X, Y), R:=R, G:=G, B:=B
        lngColor(0, X, Y) = R
        lngColor(1, X, Y) = G
        lngColor(2, X, Y) = B
      End If
    Next
    If blnShowProgress Then
      frmPaint.UpdateStatusBar intInfo:=conStRetrieveingColor, _
                               intPercentage:=((X * 100) \ X2)
    End If
  Next
  Exit Sub
  
ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub
