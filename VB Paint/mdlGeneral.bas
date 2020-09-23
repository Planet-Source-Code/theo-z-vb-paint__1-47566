Attribute VB_Name = "mdlGeneral"
'*******************************************************************************
'** File Name  : mdlGeneral.bas                                               **
'** Language   : Visual Basic 6.0                                             **
'** Reference  : Microsoft Scripting Runtime (only for ForceSave sub)         **
'** Components : -                                                            **
'** Modules    : -                                                            **
'** Developer  : Theo Zacharias (theo_yz@yahoo.com)                           **
'** Description: A modul to handle other public operations                    **
'** Last modified on August 15, 2003                                          **
'*******************************************************************************

Option Explicit

Public Enum enmError
  conErrWrite = 1
  conErrPrint = 2
  conErrReadImage = 3
  conErrDrawing = 4
  conErrPermission = 70
  conErrCancel = 32755
  conErrOthers = 0
End Enum

' Purpose    : Determine whether file strFileName exist or not
' Assumptions: -
' Effects    : -
' Input      : strFileName
' Returns    : True if file strFileName exist, false otherwise
Public Function blnFileExist(strFileName As String) As Boolean
  'On Error GoTo ErrorHandler
  
  Dim blnReturn As Boolean
  Dim fso As Scripting.FileSystemObject

  Set fso = New Scripting.FileSystemObject
  blnReturn = fso.FileExists(strFileName)
  Set fso = Nothing
  blnFileExist = blnReturn
  Exit Function

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Function

' Purpose    : Remove read-only/hidden attributes of file strFileName if users
'              agree to
' Assumption : -
' Effect     : As specified
' Input      : strFileName
' Return     : True if users agree to remove the attributes, false otherwise
Public Function ForceSave(strFileName As String) As Boolean
  'On Error GoTo ErrorHandler

  Dim fso As Scripting.FileSystemObject

  If MsgBox("The file is read-only/hidden. " & vbNewLine & _
            "Are you sure you want to write into the file " & _
            "and remove the read-only/hidden property?", _
            vbYesNo + vbQuestion) = vbYes Then
    Set fso = New Scripting.FileSystemObject
    fso.GetFile(strFileName).Attributes = 0
    ForceSave = True
  Else
    ForceSave = False
  End If
  Exit Function

ErrorHandler:
  ForceSave = False
  ShowErrMessage intErr:=conErrWrite
End Function

' Purpose    : Show error message intErr
' Assumptions: -
' Effect     : The error message has just been showed
' Inputs     : * intErr (error number)
'              * strMessage (for intErr = conErrOthers)
' Returns    : -
Public Sub ShowErrMessage(intErr As enmError, Optional strErrMessage As String)
  Select Case intErr
    Case conErrWrite
      MsgBox "Cannot write to the disk." & vbNewLine & vbNewLine & _
               "Make sure the disk is not full or write-protected.", _
             vbOKOnly + vbCritical
    Case conErrPrint
      MsgBox "Cannot print the file." & vbNewLine & vbNewLine & _
               "Make sure the print is ready.", vbOKOnly + vbCritical
    Case conErrReadImage
      MsgBox "Cannot open the file." & vbNewLine & vbNewLine & _
               "The file may be corrupt or not a valid picture file.", _
             vbOKOnly + vbCritical
    Case conErrDrawing
      MsgBox "Cannot drawing using the selected tool." & _
               vbNewLine & vbNewLine & _
               "The needed file may be missing.", _
             vbOKOnly + vbCritical
    Case conErrOthers
      MsgBox strErrMessage, vbOKOnly + vbCritical
  End Select
End Sub

' Purpose    : Get file name with or without its extention from its path strPath
' Assumptions: -
' Effects    : -
' Inputs     : strPath. blnNoExt, blnNoPath
' Return     : As specified
Public Function strGetFileName(strPath As String, _
                               Optional blnNoExt As Boolean = True, _
                               Optional blnNoPath As Boolean = True) As String
  Dim intIxDot As Integer
  Dim strReturn As String
  
  'On Error GoTo ErrorHandler
  
  If blnNoPath Then
    strReturn = Dir(strPath)
  Else
    strReturn = strPath
  End If
  If blnNoExt Then
    intIxDot = InStrRev(strReturn, ".")
    If intIxDot <> 0 Then
      strReturn = Left(strReturn, intIxDot - 1)
    End If
  End If
  strGetFileName = strReturn
  Exit Function

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Function

' Purpose    : Return varTrue if blnCondition = true, or varFalse otherwise
' Assumptions: -
' Effects    : -
' Inputs     : blnCondition, varTrue, varFalse
' Returns    : As specified
Public Function varIIf(blnCondition As Boolean, _
                        varTrue As Variant, varFalse As Variant) As Variant
  'On Error GoTo ErrorHandler
  
  If blnCondition Then
    varIIf = varTrue
  Else
    varIIf = varFalse
  End If
  Exit Function

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Function
