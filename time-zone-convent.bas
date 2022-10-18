'Original answer:
'https://stackoverflow.com/questions/3120915/get-timezone-information-in-vba-excel

'Format timezones
'https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/format-function-visual-basic-for-applications



Sub TimeZone()

Dim sourceTZString As String
Dim destTZString As String
sourceTZString = "GMT Standard Time"
destTZString = "Tokyo Standard Time"
'destTZString = "Pacific Standard Time"
dateTimeFormat = "dd mmm yyyy hh:mm:ss"

Dim OutlookApp As Object
Dim TZones As TimeZones
Dim convertedTime As Date
Dim inputTime As Date

Dim sourceTZ As TimeZone
Dim destTZ As TimeZone
Dim secNum As Integer

Set OutlookApp = CreateObject("Outlook.Application")
Set TZones = OutlookApp.TimeZones

Set sourceTZ = TZones.Item(sourceTZString)
Set destTZ = TZones.Item(destTZString)

inputTime = Now

Debug.Print sourceTZString & ": " & Format(inputTime, dateTimeFormat)
'' the outlook rounds the seconds to the nearest minute
'' thus, we store the seconds, convert the truncated time and add them later
secNum = Second(inputTime)
inputTime = DateAdd("s", -secNum, inputTime)
convertedTime = TZones.ConvertTime(inputTime, sourceTZ, destTZ)
convertedTime = DateAdd("s", secNum, convertedTime)
Debug.Print destTZString & ": " & Format(convertedTime, dateTimeFormat)
Debug.Print

End Sub

