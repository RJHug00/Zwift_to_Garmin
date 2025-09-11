Attribute VB_Name = "Zwift"
Option Explicit

' This macro downloads a FIT file created by Zwift, edits it, and writes a FIT file suitable for Garmin upload.
' [Remember also https://www.fitfiletools.com/#/top and https://www.fitfileviewer.com as diagnostic aids.]

Private Const GARMIN_DEVICE_NAME As String = "Venu-2"
Private Const GARMIN_DEVICE_PRODUCT As Integer = 3703       ' Venu 2
Private Const GARMIN_DEVICE_SERIALNo As String = "xxxxxxxxxx"

Private Const ZWIFT_USER = "your_username", ZWIFT_PW = "your_password"
' there's invariably a place to acquire the hardcoded components of this URI:
'                                                 ----------------------    and       -------
Private Const ZWIFT_REST_URI As String = "https://us-or-rly101.zwift.com/api/profiles/nnnnnnn/activities/?start=0&limit=5"
Private Const ZWIFT_RIDE_N As Integer = 1     ' 1 for most recent Zwift; 2 for previous ride, 3 for the ride before that, etc.

Private Const FIT_FILE_NAME As String = "%USERPROFILE%\Downloads\Zwift-~YYYY-mm-DD-HH-MM-SS~"
Private Const ZWIFT_TIMEZONE_BIAS As Double = (-4 / 24)      ' EST-5

Private Const INSERTED_STRING_SIZE As Integer = 100

Public Sub Garmin_from_Zwift()

  Dim i As Integer, j As Integer, k As Integer, IIndex As Long, LastIndex As Long, IJndex As Long
  Dim tpi As clsFITfield, msgDict(16) As clsFITmsg, msgDef As clsFITmsg, ffb As String
  Dim lclMsgId As Integer, mMsgId As Integer, accumulated_offset As Long, ffk As String
  Dim oHttpReq As Object, JSON As String, FIT() As Byte, fft As String, zwiftTime As Date
  Dim expandedFileName As String, ss As String, fNo As Integer: fNo = FreeFile

  Application.Cursor = xlWait
  Application.Calculation = xlCalculationManual
  Application.ScreenUpdating = False

  Set oHttpReq = CreateObject("MSXML2.ServerXMLHTTP")
  oHttpReq.SetTimeouts 30000, 30000, 30000, 30000

'=============================================
'  Get FIT file for the most recent Zwift ride
'=============================================

  oHttpReq.Open "POST", "https://secure.zwift.com/auth/realms/zwift/protocol/openid-connect/token", False
  oHttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded" ' application/json
  oHttpReq.SetRequestHeader "Accept", "application/json"
  oHttpReq.Send "username=" & ZWIFT_USER & "&password=" & ZWIFT_PW & "&grant_type=password&client_id=Zwift+Game+Client"

  JSON = IIf(oHttpReq.Status = 200, oHttpReq.responseText, "")

  IIndex = InStr(JSON, "token") + 8
  ss = Mid(JSON, IIndex): ss = Left(ss, InStr(ss, """") - 1)    ' The auth token

  oHttpReq.Open "GET", ZWIFT_REST_URI, False
  oHttpReq.SetRequestHeader "Accept", "application/json"
  oHttpReq.SetRequestHeader "Authorization", "Bearer " & ss
  oHttpReq.Send ""

  JSON = IIf(oHttpReq.Status = 200, oHttpReq.responseText, "")

  IIndex = 1: IJndex = ZWIFT_RIDE_N    ' normally 1 for 'most recent'
  While IJndex > 0
    IJndex = IJndex - 1: IIndex = InStr(IIndex, JSON, """name""") + 16
  Wend

  ss = Mid(JSON, IIndex, 100): fft = Left(ss, InStr(ss, """") - 1)

  IIndex = InStr(IIndex, JSON, "sport") + 8
  If Mid(JSON, IIndex, 7) <> "CYCLING" Then _
    MsgBox "Not a Zwift CYCLING activity", vbExclamation, "Error": GoTo finito
  
  IIndex = InStr(IIndex, JSON, "startDate") + 12
  zwiftTime = CDate(Replace(Mid(JSON, IIndex, 19), "T", " "))  ' UTC
  zwiftTime = zwiftTime + ZWIFT_TIMEZONE_BIAS
  
  IIndex = InStr(IIndex, JSON, "fitFileBucket") + 16
  ss = Mid(JSON, IIndex, 32): ffb = Left(ss, InStr(ss, """") - 1)

  IIndex = InStr(IIndex, JSON, "fitFileKey") + 13
  ss = Mid(JSON, IIndex, 64): ffk = Left(ss, InStr(ss, """") - 1)

  oHttpReq.Open "GET", "https://" & ffb & ".s3.amazonaws.com/" & ffk, False
  oHttpReq.SetRequestHeader "referer", "https://www.zwift.com/"
  oHttpReq.Send ""

  If oHttpReq.Status = 200 Then
    FIT = oHttpReq.responseBody
  Else
    MsgBox "Error retrieving Zwift file from AWS", vbExclamation, "Error": GoTo finito
  End If

  Set oHttpReq = Nothing

  LastIndex = UBound(FIT) - 2    ' stop looping short of the trailing checksum
  IIndex = FIT(0)                ' point past the FIT header to the first FIT record

'============================================
' Edit the FIT file for Garmin considerations
'============================================

NextRec:
  lclMsgId = FIT(IIndex)
  mMsgId = lclMsgId And &HF
  If (lclMsgId And &H40) > 0 Then             ' Definition rec
    IJndex = IIndex                           ' remember where this definition starts
    Set msgDict(mMsgId) = New clsFITmsg
    msgDict(mMsgId).Init FIT, IIndex          ' decodes the definition record
    Set msgDef = msgDict(mMsgId)

    If msgDef.Name = "Session" Then FIT(IJndex + 5) = FIT(IJndex + 5) + 1   ' we'll add one field definition later

  Else                                        ' Data
    Set msgDef = msgDict(mMsgId)              ' retrieve the decoded definition
    ss = msgDef.Name
    accumulated_offset = 1
    j = msgDef.NumFields - 1

    ' a bit confusing, I know.  We waited for the start of the Session DATA to amend the definition just passed.
    ' and at the same time, increase the size of FIT for both that amendment *and* the string we'll insert shortly.
    
    If msgDef.Name = "Session" Then
      ReDim Preserve FIT(UBound(FIT) + 3 + INSERTED_STRING_SIZE)
      LastIndex = LastIndex + 3 + INSERTED_STRING_SIZE
      IJndex = UBound(FIT) - INSERTED_STRING_SIZE
      While IJndex > IIndex
        FIT(IJndex) = FIT(IJndex - 3): IJndex = IJndex - 1    ' make a 3-byte hole right here
      Wend
      FIT(IIndex) = &H6E: FIT(IIndex + 1) = INSERTED_STRING_SIZE: FIT(IIndex + 2) = &H7
      IIndex = IIndex + 3      ' adjust the start of the data record
                          ' don't increment j to reflect the added field (msgDef.NumFields is one too few)
    End If

    For i = 0 To j
      Set tpi = msgDef.FieldDef(i): k = tpi.offset
      accumulated_offset = accumulated_offset + tpi.sz
      ss = tpi.GlobalCodeToHuman(msgDef.globalMsgId)    ' the name of the field

      If msgDef.Name = "FileID" Then
        Select Case ss
          Case "Manufacturer": Call Set_2(FIT, IIndex + k, 1)                 ' Garmin instead of Zwift
          Case "Product": Call Set_2(FIT, IIndex + k, GARMIN_DEVICE_PRODUCT)  ' Venu-2
        End Select
      End If

      If msgDef.Name = "DeviceInfo" Then
        Select Case ss
          Case "Manufacturer": Call Set_2(FIT, IIndex + k, 1)                 ' Garmin instead of Zwift
          Case "Product": Call Set_2(FIT, IIndex + k, GARMIN_DEVICE_PRODUCT)  ' Venu-2
          Case "Serial#": Call Set_4(FIT, IIndex + k, get_ULongFromLargeDecimal(GARMIN_DEVICE_SERIALNo))
          Case "Product Name": Call SetString(FIT, IIndex + k, GARMIN_DEVICE_NAME, 20)
        End Select
      End If
    Next i    ' next field in a data rec

    IIndex = IIndex + accumulated_offset

    If msgDef.Name = "Session" Then    ' if we just processed the Session data rec
      IJndex = UBound(FIT)
      While IJndex > IIndex
        FIT(IJndex) = FIT(IJndex - INSERTED_STRING_SIZE): IJndex = IJndex - 1
      Wend
      For IJndex = IIndex To IIndex + INSERTED_STRING_SIZE - 1
        FIT(IJndex) = 0   ' zero fill the string field we're setting
      Next IJndex
      IJndex = IIndex  ' SetString clobbers arg #2 and we need IIndex preserved
      Call SetString(FIT, IJndex, "Zwift:" & fft, 0)
      IIndex = IIndex + INSERTED_STRING_SIZE
    End If

  End If    ' definition vs data rec

  If IIndex < LastIndex Then GoTo NextRec

  Call Initialize_CRC_table
  Call Set_4(FIT, 4, UBound(FIT) - 15) ' update header with revised length
  Call Set_2(FIT, 12, 0)               ' force header checksum unused just in-case
  Call Set_2(FIT, LastIndex + 1, GenerateCkSum(FIT, 0, LastIndex))

  expandedFileName = ExpandFileName(FIT_FILE_NAME, zwiftTime) & ".FIT"

  On Error Resume Next: Kill expandedFileName: On Error GoTo 0
  Open expandedFileName For Binary Access Write As #fNo
  Put #fNo, , FIT
  Close #fNo

finito:
  Set oHttpReq = Nothing
  Application.Cursor = xlDefault
  Application.Calculation = xlCalculationAutomatic
  Application.ScreenUpdating = True
End Sub

Public Function ExpandFileName(inString As String, timestamp As Date) As String
  Dim v() As String, s As String: s = inString
  Dim token As String, i As Integer
  ' the filename can be patterned to include date and/or time
  i = InStr(s, "~")
  If i > 0 Then
    token = Mid(s, i, 32): token = Left(token, InStr(2, token, "~"))
    s = Replace(s, token, Format(timestamp, Mid(token, 2, Len(token) - 2)))
  End If
  ' tokens delimited by '%' are environment variables to be expanded
  v = Split(s, "%")
  While UBound(v) > 0
    s = Replace(s, "%" & v(1) & "%", ParamGet(v(1))) ' replace all occurrences
    v = Split(s, "%")                                ' look for more variables
  Wend
  ExpandFileName = s   ' return the substitution string
End Function
