VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFITfield"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public gblIndex As Integer
Public sz As Integer
Public tpCode As Integer
Public offset As Integer

' if VBA had a constructor like other languages, this would be there.
Public Sub Init(idx As Byte, isz As Byte, itp As Byte, ioff As Integer)
  gblIndex = idx
  sz = isz
  tpCode = itp
  offset = ioff
End Sub

Public Property Get TypeToHuman() As String
  Dim s As String
  Select Case tpCode
    Case 0: s = "Enum"
    Case 1: s = "SInt8"
    Case 2: s = "UInt8"
    Case 7: s = "Str"
    Case 10: s = "ZInt8"
    Case 13: s = "Byte"
    Case 131: s = "SInt16"
    Case 132: s = "UInt16"
    Case 133: s = "SInt32"
    Case 134: s = "UInt32"
    Case 136: s = "Float"
    Case 137: s = "Flt64"
    Case 139: s = "ZInt16"
    Case 140: s = "ZInt32"
    Case 142: s = "SInt64"
    Case 143: s = "UInt64"
    Case 144: s = "ZInt64"
    Case 253: s = "TStamp"
    Case Else: s = CStr(tpCode)
  End Select
  Select Case sz
    Case 1, 2, 4:
    Case Else: s = s & "(" & sz & ")"   ' Str, Byte, UInt8, UInt8z have this
  End Select
  TypeToHuman = s
End Property

Public Property Get GlobalCodeToHuman(gblMsgCode As Integer) As String

 ' It *seems* like 253 is always a timestamp, but that isn't documented anywhere
 If gblIndex = 253 Then GlobalCodeToHuman = "Time": Exit Property
 Dim s As String

 Select Case gblMsgCode
  ' fast-path: these are almost every second, aren't defined, and aren't useful
  Case GARMIN_TYPE.gGPS_META, GARMIN_TYPE.gHRV
   s = "WhoTFC"
  Case GARMIN_TYPE.gRECORD  ' Most frequent useful; every second
   Select Case gblIndex
    Case 0: s = "Lat"
    Case 1: s = "Lon"
    Case 5: s = "Distance"
    Case 11: s = "Time From Course"
    Case 19: s = "Tot Cycles"
    Case 29: s = "Accumulated Power"
    Case 73: s = "Enh. Speed"
    Case 78: s = "Enh. Altitude"
    Case 2: s = "Altitude"
    Case 6: s = "Speed m/s"    ' deprecated in favor of #73
    Case 7: s = "Watts"
    Case 9: s = "Grade"
    Case 28: s = "Compressed Accumulated Power"
    Case 32: s = "Vert Speed"
    Case 33: s = "Calories"
    Case 39: s = "Vert Oscillation"
    Case 40: s = "Stance Time Pct"
    Case 41: s = "Stance Time"
    Case 51: s = "Ball Speed"
    Case 52: s = "Cadence256"
    Case 54: s = "Tot Hemoglobin Conc"
    Case 55: s = "Tot Hemoglobin Conc Min"
    Case 56: s = "Tot Hemoglobin Conc Max"
    Case 57: s = "Sat Hemoglobin Pct"
    Case 58: s = "Sat Hemoglobin Pct Min"
    Case 59: s = "Sat Hemoglobin Pct Max"
    Case 3: s = "HR"
    Case 4: s = "Cadence"
    Case 8: s = "Compressed Speed Dist"
    Case 10: s = "Resistance"
    Case 12: s = "Cycle Length"
    Case 13: s = "Temp"
    Case 17: s = "Speed 1S"
    Case 18: s = "Cycles"
    Case 30: s = "L R Balance"
    Case 31: s = "GPS Accuracy"
    Case 42: s = "Activity Type"
    Case 43: s = "L Torque Eff"
    Case 44: s = "R Torque Eff"
    Case 45: s = "L Pedal Smoothness"
    Case 46: s = "R Pedal Smoothness"
    Case 47: s = "Combined Pedal Smoothness"
    Case 48: s = "Time128"
    Case 49: s = "Stroke Type"
    Case 50: s = "Zone"
    Case 53: s = "Fractional Cadence"
    Case 62: s = "Device Index"
    Case 136: s = "*More HR"
    Case Else: s = "_" & CStr(gblIndex)
   End Select
  Case GARMIN_TYPE.gLAP
   Select Case gblIndex
    Case 2: s = "StartTime"
    Case 3: s = "StartLat"
    Case 4: s = "StartLon"
    Case 5: s = "EndLat"
    Case 6: s = "EndLon"
    Case 7: s = "Total Elapsed Time"
    Case 8: s = "Total Timer Time"
    Case 9: s = "TotalDist"
    Case 10: s = "Tot Cycles"
    Case 41: s = "Tot Work"
    Case 52: s = "Tot Moving Time"
    Case 57: s = "Time in HR Zone"
    Case 58: s = "Time in Speed Zone"
    Case 59: s = "Time in Cadence Zone"
    Case 60: s = "Time in Power Zone"
    Case 110: s = "Enh. Avg Speed"
    Case 111: s = "Enh. Max Speed"
    Case 112: s = "Enh. Avg Alt"
    Case 113: s = "Enh. Min Alt"
    Case 114: s = "Enh. Max Alt"
    Case 254: s = "Msg Index"
    Case 11: s = "Total Calories"
    Case 12: s = "Total Fat Calories"
    Case 13: s = "Avg Speed"
    Case 14: s = "Max Speed"
    Case 19: s = "Avg Power"
    Case 20: s = "Max Power"
    Case 21: s = "Total Ascent"
    Case 22: s = "Total Descent"
    Case 32: s = "Num Lengths"
    Case 33: s = "Normalized Power"
    Case 34: s = "Left Right Balance"
    Case 35: s = "1st Length Index"
    Case 37: s = "Avg Stroke Distance"
    Case 40: s = "Num Active Lengths"
    Case 42: s = "Avg Alt"
    Case 43: s = "Max Alt"
    Case 45: s = "Avg Grade"
    Case 46: s = "Avg Pos Grade"
    Case 47: s = "Avg Neg Grade"
    Case 48: s = "Max Pos Grade"
    Case 49: s = "Max Neg Grade"
    Case 53: s = "Avg Pos Vert Speed"
    Case 54: s = "Avg Neg Vert Speed"
    Case 55: s = "Max Pos Vert Speed"
    Case 56: s = "Max Neg Vert Speed"
    Case 61: s = "Repetition Num"
    Case 62: s = "Min Altitude"
    Case 71: s = "Wkt Step Index"
    Case 74: s = "Opponent Score"
    Case 75: s = "Stroke Count"
    Case 76: s = "Zone Count"
    Case 77: s = "Avg Vert Oscillation"
    Case 78: s = "Avg Stance Time Pct"
    Case 79: s = "Avg Stance Time"
    Case 83: s = "Player Score"
    Case 84: s = "Avg Total Hemoglobin Conc"
    Case 85: s = "Min Total Hemoglobin Conc"
    Case 86: s = "Max Total Hemoglobin Conc"
    Case 87: s = "Avg Saturated Hemoglobin Pct"
    Case 88: s = "Min Saturated Hemoglobin Pct"
    Case 89: s = "Max Saturated Hemoglobin Pct"
    Case 121: s = "Avg VAM"
    Case 0: s = "Event"
    Case 1: s = "Event Type"
    Case 15: s = "Avg HR"
    Case 16: s = "Max HR"
    Case 17: s = "Avg Cadence"
    Case 18: s = "Max Cadence"
    Case 23: s = "Intensity"
    Case 24: s = "Lap Trigger"
    Case 25: s = "Sport"
    Case 26: s = "Event Group"
    Case 38: s = "Swim Stroke"
    Case 39: s = "Sub Sport"
    Case 44: s = "GPS Accuracy"
    Case 50: s = "Avg Temp"
    Case 51: s = "Max Temp"
    Case 63: s = "Min HR"
    Case 80: s = "Avg Fractional Cadence"
    Case 81: s = "Max Fractional Cadence"
    Case 82: s = "Total Fractional Cycles"
    Case 98: s = "Total Time Standing"
    Case 102: s = "Avg L.Pwr Phase"
    Case 103: s = "Avg L.Pwr Phase Peak"
    Case 104: s = "Avg R.Pwr Phase"
    Case 105: s = "Avg R.Pwr Phase Peak"
    Case 106: s = "Avg Pwr Position"
    Case 107: s = "Max Pwr Position"
    Case 149: s = "Bike Grit Score"
    Case 150: s = "Bike Flow Score"
    Case 99: s = "Standing Count"
    Case 118: s = "Avg Vert Ratio"
    Case 119: s = "Avg Stance Time Balance"
    Case 120: s = "Avg Step Length"
    Case 136: s = "Enh. Avg Resp Rate"
    Case 137: s = "Enh. Max Resp Rate"
    Case 91: s = "Avg L.Torque Eff"
    Case 92: s = "Avg R.Torque Eff"
    Case 124: s = "Min Temp"
    Case 147: s = "Avg Resp Rate"
    Case 148: s = "Max Resp Rate"
    Case 93: s = "Avg L.Pedal Smooth"
    Case 94: s = "Avg R.Pedal Smooth"
    Case 95: s = "Avg Combined Pedal Smooth"
    Case 100: s = "Avg L.Plat C.Off"
    Case 101: s = "Avg R.Plat C.Off"
    Case 108: s = "Avg CadencePos"
    Case 109: s = "Max CadencePos"
    Case 156: s = "Tot Frac Ascent"
    Case 157: s = "Tot Frac Descent"
    Case Else: s = "_" & CStr(gblIndex)
   End Select
  Case GARMIN_TYPE.gEVENT
   Select Case gblIndex
    Case 3: s = "Data"
    Case 2: s = "Data16"
    Case 7: s = "Score"
    Case 8: s = "Opponent Score"
    Case 0: s = "Event"
    Case 1: s = "Event Type"
    Case 4: s = "Event Group"
    Case 9: s = "Front Gear #"
    Case 10: s = "Front Gear"
    Case 11: s = "Rear Gear #"
    Case 12: s = "Rear Gear"
    Case 21: s = "Radar Threat Level Max"
    Case 22: s = "Radar Threat Count"
    Case Else: s = "_" & CStr(gblIndex)
   End Select
  Case GARMIN_TYPE.gTIME_IN_ZONE, GARMIN_TYPE.gSPLIT  ' somewhat often but equally un-useful
   s = "WhoTFC"
  Case GARMIN_TYPE.gACTIVITY
   Select Case gblIndex
    Case 0: s = "TotalTimer"
    Case 5: s = "LocalTime"
    Case 1: s = "#Sessions"
    Case 2: s = "Type"
    Case 3: s = "Event"
    Case 4: s = "EventType"
    Case 6: s = "EventGrp"
    Case Else: s = "_" & CStr(gblIndex)
   End Select
  Case GARMIN_TYPE.gSPORT
   Select Case gblIndex
    Case 3: s = "Name"
    Case 0: s = "Sport#"
    Case 1: s = "SubSport"
    Case Else: s = "_" & CStr(gblIndex)
   End Select
  Case GARMIN_TYPE.gSESSION
   Select Case gblIndex
    Case 2: s = "StartTime"
    Case 3: s = "StartLat"
    Case 4: s = "StartLon"
    Case 7: s = "ElapsedTime"
    Case 8: s = "TimerTime"
    Case 9: s = "Distance"
    Case 10: s = "Cycles"
    Case 29: s = "NEC Lat"
    Case 30: s = "NEC Lon"
    Case 31: s = "SWC Lat"
    Case 32: s = "SWC Lon"
    Case 41: s = "Avg Stroke Count"
    Case 48: s = "Total Work"
    Case 59: s = "Total Moving Time"
    Case 65: s = "Time in HR Zone"
    Case 66: s = "Time in Speed Zone"
    Case 67: s = "Time in Cadence Zone"
    Case 68: s = "Time in Power Zone"
    Case 69: s = "Avg Lap Time"
    Case 124: s = "Enh. Avg Speed"
    Case 125: s = "Enh. Max Speed"
    Case 126: s = "Enh. Avg Altitude"
    Case 127: s = "Enh. Min Altitude"
    Case 128: s = "Enh. Max Altitude"
    Case 254: s = "Msg Index"
    Case 11: s = "Total Calories"
    Case 13: s = "Total Fat Calories"
    Case 14: s = "Avg Speed"
    Case 15: s = "Max Speed"
    Case 20: s = "Avg Power"
    Case 21: s = "Max Power"
    Case 22: s = "Total Ascent"
    Case 23: s = "Total Descent"
    Case 25: s = "1st Lap Index"
    Case 26: s = "Num Laps"
    Case 33: s = "Num Lengths"
    Case 34: s = "Normalized Power"
    Case 35: s = "Training Stress Score"
    Case 36: s = "Intensity Factor"
    Case 37: s = "Left Right Balance"
    Case 42: s = "Avg Stroke Distance"
    Case 44: s = "Pool Length"
    Case 45: s = "Threshold Power"
    Case 47: s = "Num Active Lengths"
    Case 49: s = "Avg Altitude"
    Case 50: s = "Max Altitude"
    Case 52: s = "Avg Grade"
    Case 53: s = "Avg Pos Grade"
    Case 54: s = "Avg Neg Grade"
    Case 55: s = "Max Pos Grade"
    Case 56: s = "Max Neg Grade"
    Case 60: s = "Avg Pos Vert Speed"
    Case 61: s = "Avg Neg Vert Speed"
    Case 62: s = "Max Pos Vert Speed"
    Case 63: s = "Max Neg Vert Speed"
    Case 70: s = "Best Lap Index"
    Case 71: s = "Min Altitude"
    Case 82: s = "Player Score"
    Case 83: s = "Opponent Score"
    Case 85: s = "Stroke Count"
    Case 86: s = "Zone Count"
    Case 87: s = "Max Ball Speed"
    Case 88: s = "Avg Ball Speed"
    Case 89: s = "Avg Vert Oscillation"
    Case 90: s = "Avg Stance Time Pct"
    Case 91: s = "Avg Stance Time"
    Case 139: s = "Avg VAM"
    Case 0: s = "Event"
    Case 1: s = "Event Type"
    Case 5: s = "Sport"
    Case 6: s = "Sub Sport"
    Case 16: s = "Avg HR"
    Case 17: s = "Max HR"
    Case 18: s = "Avg Cadence"
    Case 19: s = "Max Cadence"
    Case 24: s = "Total Aerobic TE"
    Case 27: s = "Event Group"
    Case 28: s = "Trigger"
    Case 43: s = "Swim Stroke"
    Case 46: s = "Pool Length Unit"
    Case 51: s = "GPS Accuracy"
    Case 57: s = "Avg Temp"
    Case 58: s = "Max Temp"
    Case 64: s = "Min HR"
    Case 84: s = "Opponent Name"
    Case 92: s = "Avg Fractional Cadence"
    Case 93: s = "Max Fractional Cadence"
    Case 94: s = "Total Fractional Cycles"
    Case 111: s = "Sport Index"
    Case 137: s = "Total Anaerobic TE"
    Case Else: s = "_" & CStr(gblIndex)
   End Select
  Case GARMIN_TYPE.gCREATOR
   Select Case gblIndex
    Case 0: s = "SWvsn"
    Case 1: s = "HWvsn"
    Case Else: s = "_" & CStr(gblIndex)
   End Select
  Case GARMIN_TYPE.gFILE_ID
   Select Case gblIndex
    Case 3: s = "Serial#"
    Case 4: s = "Created"  ' time
    Case 8: s = "ProdName" ' name string
    Case 1: s = "Manufacturer"
    Case 2: s = "Product"  ' code
    Case 5: s = "#"        ' only certain files
    Case 0: s = "Type"     ' File type
    Case Else: s = "_" & CStr(gblIndex)
   End Select
  Case GARMIN_TYPE.gDEVICE_INFO
   Select Case gblIndex
    Case 3: s = "Serial#"
    Case 7: s = "Cumulative Operating Time"
    Case 27: s = "Product Name"
    Case 2: s = "Manufacturer"
    Case 4: s = "Product"
    Case 5: s = "SW Vsn"
    Case 10: s = "Battery Voltage"
    Case 21: s = "ANT Device#"
    Case 0: s = "Device Index"
    Case 1: s = "Device Type"
    Case 6: s = "HW Vsn"
    Case 11: s = "Battery Status"
    Case 18: s = "Sensor Position"
    Case 19: s = "Descriptor"
    Case 20: s = "ANT Transmission Type"
    Case 22: s = "ANT Network"
    Case 25: s = "Source Type"
    Case Else: s = "_" & CStr(gblIndex)
   End Select
  Case GARMIN_TYPE.gDEVICE_SETTINGS
   Select Case gblIndex
    Case 1: s = "UTC Offset"
    Case 2: s = "Time Offset"
    Case 39: s = "Clock Time"
    Case 40: s = "Pages Enabled"
    Case 57: s = "Default Page"
    Case 58: s = "Autosync Min Steps"
    Case 59: s = "Autosync Min Time"
    Case 0: s = "Active Time Zone"
    Case 4: s = "Time Mode"
    Case 5: s = "TimeZone Offset"
    Case 12: s = "Backlight Mode"
    Case 36: s = "Activity Tracker Enabled"
    Case 46: s = "Move Alert Enabled"
    Case 47: s = "Date Mode"
    Case 55: s = "Display Orientation"
    Case 56: s = "Mounting Side"
    Case 174: s = "Tap Sensitivity"
    Case Else: s = "_" & CStr(gblIndex)
   End Select
  Case GARMIN_TYPE.gUSER_PROFILE
   Select Case gblIndex
    Case 0: s = "Friendly Name"
    Case 254: s = "Msg Index"
    Case 4: s = "Weight"
    Case 22: s = "Local ID"
    Case 31: s = "Run Step Length"
    Case 32: s = "Walk Step Length"
    Case 1: s = "Gender"
    Case 2: s = "Age"
    Case 3: s = "Height"
    Case 5: s = "Language"
    Case 6: s = "Elev Setting"
    Case 7: s = "Weight Setting"
    Case 8: s = "Resting HR"
    Case 9: s = "Default Max Run HR"
    Case 10: s = "Default Max Bike HR"
    Case 11: s = "Default Max HR"
    Case 12: s = "HR Setting"
    Case 13: s = "Speed Setting"
    Case 14: s = "Distance Setting"
    Case 16: s = "Power Setting"
    Case 17: s = "Activity Class"
    Case 18: s = "Position Setting"
    Case 21: s = "Temp Setting"
    Case 23: s = "Global ID"
    Case 28: s = "Wake Time"
    Case 29: s = "Sleep Time"
    Case 30: s = "Height Setting"
    Case 47: s = "Depth Setting"
    Case Else: s = "_" & CStr(gblIndex)
   End Select
  Case GARMIN_TYPE.gZONES_TARGET
   Select Case gblIndex
    Case 3: s = "FTP"
    Case 1: s = "Max HR"
    Case 2: s = "Thresh HR"
    Case 5: s = "HR Calc Type"
    Case 7: s = "PWR Calc Type"
    Case Else: s = "_" & CStr(gblIndex)
   End Select
  ' when the specified global message is undocumented or not implemented
  Case Else: s = CStr(gblMsgCode) & ":_" & CStr(gblIndex)
 End Select
 GlobalCodeToHuman = s
End Property
