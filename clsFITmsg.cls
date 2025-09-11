VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFITmsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public globalMsgId As Integer
Public bigEndian As Boolean
Private flds() As clsFITfield
Private hexRepresentation As String

' if VBA had a constructor like other languages, this would be there.
Public Sub Init(bytArr() As Byte, ByRef indx As Long)
  Dim ef As endian_flipper, i As Integer, nFields As Integer, offset As Integer
  Dim dData As Boolean, nDFields As Integer, idx As Integer

  i = bytArr(indx) And &HBF ' knock down the x40 bit
  If (i And &H20) > 0 Then
    dData = True
    'MsgBox "Unable to process developer data messages!", vbOKOnly, "Error": Err.Raise 514
    'Debug.Print "Developer definition; continuing anyway"
  End If

  bigEndian = bytArr(indx + 2) <> arch.LITTLE

  If bigEndian Then
    ef.b1 = bytArr(indx + 3): ef.b0 = bytArr(indx + 4)
  Else
    ef.b0 = bytArr(indx + 3): ef.b1 = bytArr(indx + 4)
  End If
  RtlCopyMemory globalMsgId, ef, 2

  nFields = bytArr(indx + 5) - 1: ReDim flds(nFields)

  hexRepresentation = ZHex(bytArr(indx)) & ZHex(bytArr(indx + 1)) & ZHex(bytArr(indx + 2)) & _
                      ZHex(bytArr(indx + 3)) & ZHex(bytArr(indx + 4)) & ZHex(bytArr(indx + 5))

  indx = indx + 6                 ' 1st field
  offset = 1                      ' pre-compensated for first message byte

  For idx = 0 To nFields
    Set flds(idx) = New clsFITfield
    flds(idx).Init bytArr(indx), bytArr(indx + 1), bytArr(indx + 2), offset
    hexRepresentation = hexRepresentation & ZHex(bytArr(indx)) & ZHex(bytArr(indx + 1)) & ZHex(bytArr(indx + 2))
    offset = offset + bytArr(indx + 1)
    indx = indx + 3
  Next idx
  
  ' if the developer data bit is on, more fields follow
  ' and I need to fabricate an additional flds() element for each one
  If dData Then
    nDFields = bytArr(indx): indx = indx + 1: hexRepresentation = hexRepresentation & ZHex(nDFields)
    ReDim Preserve flds(nFields + nDFields)
    For i = 1 To nDFields
      Set flds(idx) = New clsFITfield
      flds(idx).Init bytArr(indx), bytArr(indx + 1), bytArr(indx + 2), offset: idx = idx + 1
      offset = offset + bytArr(indx + 1)
      hexRepresentation = hexRepresentation & ZHex(bytArr(indx)) & ZHex(bytArr(indx + 1)) & ZHex(bytArr(indx + 2))
      indx = indx + 3
    Next i
  End If
  'Debug.Print hexRepresentation
End Sub

Public Property Get toHexString() As String
  toHexString = hexRepresentation
End Property

Public Property Get NumFields() As Integer
  NumFields = UBound(flds) + 1
End Property

Public Property Get FieldDef(indx As Integer) As clsFITfield
  Set FieldDef = flds(indx)
End Property

Public Property Get Name() As String
  Select Case globalMsgId              'Walk Run IRide
    Case 0: Name = "FileID"             'Y    Y   Y
    Case 1: Name = "Capabilities"
    Case 2: Name = "DeviceSettings"     'Y    Y   Y
    Case 3: Name = "UserProfile"        'Y    Y   Y
    Case 4: Name = "HRM_Profile"
    Case 5: Name = "Sdm_Profile"
    Case 6: Name = "BikeProfile"
    Case 7: Name = "ZonesTgt"           'Y    Y   Y
    Case 8: Name = "HR_Zone"
    Case 9: Name = "PowerZone"
    Case 10: Name = "Met_Zone"
    Case 12: Name = "Sport"             'Y    Y   Y
    Case 15: Name = "Goal"
    Case 18: Name = "Session"           'Y    Y   Y
    Case 19: Name = "Lap"               'Y    Y   Y
    Case 20: Name = "Record"            'Y    Y   Y
    Case 21: Name = "Event"             'Y    Y   Y
    Case 23: Name = "DeviceInfo"        'Y    Y   Y
    Case 26: Name = "Workout"
    Case 27: Name = "WorkoutStep"
    Case 28: Name = "Schedule"
    Case 30: Name = "WeightScale"
    Case 31: Name = "Course"
    Case 32: Name = "CoursePt"
    Case 33: Name = "Totals"
    Case 34: Name = "Activity"          'Y    Y   Y
    Case 35: Name = "Software"
    Case 37: Name = "File_Capabilities"
    Case 39: Name = "Field_Capabilities"
    Case 49: Name = "File_Creator"           'Y    Y   Y
    Case 51: Name = "BloodPressure"
    Case 53: Name = "SpeedZone"
    Case 55: Name = "Monitoring"
    Case 72: Name = "TrainingFile"
    Case 78: Name = "HRV"
    Case 80: Name = "ANT_RX"
    Case 81: Name = "ANT_TX"
    Case 82: Name = "ANT_Channel"
    Case 101: Name = "Length"
    Case 103: Name = "MonitorInfo"
    Case 105: Name = "Pad"
    Case 106: Name = "SlaveDev"
    Case 127: Name = "Connectivity"
    Case 128: Name = "WxConditions"
    Case 129: Name = "WxAlert"
    Case 131: Name = "CadenceZone"
    Case 132: Name = "HeartRate"
    Case 142: Name = "SegmentLap"
    Case 145: Name = "MemoGlob"
    Case 148: Name = "SegmentID"
    Case 149: Name = "SegmentLdrBd"
    Case 150: Name = "SegmentPt"
    Case 151: Name = "SegmentFile"
    Case 158: Name = "WorkoutSess"
    Case 159: Name = "WatchfaceSettings"
    Case 160: Name = "GPSmeta"          'Y    Y
    Case 161: Name = "CameraEvt"
    Case 162: Name = "TimestampCorr"
    Case 164: Name = "GyroData"
    Case 165: Name = "AccelerometerData"
    Case 169: Name = "VideoFrame"
    Case 174: Name = "OBDII"
    Case 177: Name = "NMEA"
    Case 178: Name = "AviationALT"
    Case 184: Name = "Video"
    Case 185: Name = "VideoTitle"
    Case 186: Name = "VideoDesc"
    Case 187: Name = "VideoClip"
    Case 188: Name = "OHR"
    Case 200: Name = "EXDScreen"
    Case 201: Name = "EXDData1"
    Case 202: Name = "EXDData2"
    Case 206: Name = "DeveloperField"
    Case 207: Name = "DeveloperData"
    Case 208: Name = "Magnetometer"
    Case 209: Name = "Barometer"
    Case 210: Name = "OneD_Calibration"
    Case 211: Name = "Monitoring_HR"
    Case 216: Name = "TimeInZone"       'Y    Y   Y
    Case 225: Name = "Set"
    Case 227: Name = "StressLevel"
    Case 229: Name = "MaxMETdata"
    Case 258: Name = "DiveSettings"
    Case 259: Name = "DiveGas"
    Case 262: Name = "DiveAlarm"
    Case 264: Name = "ExerciseTitle"
    Case 268: Name = "DiveSummary"
    Case 269: Name = "SpO2"
    Case 275: Name = "SleepLevel"
    Case 285: Name = "Jump"
    Case 290: Name = "BeatIntervals"
    Case 297: Name = "Respiration"
    Case 312: Name = "Split"
    Case 313: Name = "SplitSumm"
    Case 317: Name = "ClimbPRO"
    Case 319: Name = "TankUpd"
    Case 323: Name = "TankSumm"
    Case 346: Name = "SleepAssess"
    Case 370: Name = "HRVSummary"
    Case 371: Name = "HRV_Value"
    Case 375: Name = "DevAuxBattery"
    Case 393: Name = "DiveApnea"
    ' seen, but un-documented:
    Case 13, 22, 79: Name = "_0" & CStr(globalMsgId)
    Case 104, 113, 140, 141, 147: Name = "_" & CStr(globalMsgId)
    Case 233, 288: Name = "_" & CStr(globalMsgId)
    Case 324, 325, 326, 327, 358, 394: Name = "_" & CStr(globalMsgId)
    Case Else: Name = "**" & globalMsgId   ' never before seen
  End Select
End Property
