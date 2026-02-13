Const olFolderCalendar = 9
Const olAppointmentItem = 1
Const olOutOfOffice = 3
'Const vbYesNo = 4
'Const vbQuestion = 32
'Const vbYes = 6

' --- Confirmation Dialog ---
intResponse = MsgBox("This script will add 2026 COA Holidays to your Outlook calendar." & vbCrLf & vbCrLf & _
                     "Do you want to continue?", vbYesNo + vbQuestion, "Confirm Calendar Update")

If intResponse <> vbYes Then
    MsgBox "Operation cancelled.", vbInformation, "Cancelled"
    WScript.Quit
End If

' --- Outlook Setup ---
Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objCalendar = objNamespace.GetDefaultFolder(olFolderCalendar) 

Set objDictionary = CreateObject("Scripting.Dictionary")
objDictionary.Add "January 1, 2026", "New Year's Day"    
objDictionary.Add "January 19, 2026", "Martin Luther King Day"
objDictionary.Add "February 16, 2026", "President's Day"
objDictionary.Add "May 25, 2026", "Memorial Day"
objDictionary.Add "June 19, 2026", "Juneteenth"
objDictionary.Add "July 3, 2026", "Independence Day Observed"
objDictionary.Add "September 07, 2026", "Labor Day"
objDictionary.Add "November 11, 2026", "Veterans Day"
objDictionary.Add "November 26, 2026", "Thanksgiving Day"
objDictionary.Add "November 27, 2026", "Thanksgiving Friday"
objDictionary.Add "December 24, 2026", "Christmas Eve"
objDictionary.Add "December 25, 2026", "Christmas Day"

colKeys = objDictionary.Keys

For Each strKey in colKeys
    dtmHolidayDate = strKey
    strHolidayName = objDictionary.Item(strKey)

    Set objHoliday = objOutlook.CreateItem(olAppointmentItem)  
    objHoliday.Subject = strHolidayName
    objHoliday.Start = dtmHolidayDate & " 12:00 AM"
    objHoliday.End = dtmHolidayDate & " 11:59 PM"
    objHoliday.ReminderSet = True
    objHoliday.ReminderMinutesBeforeStart = 10080
    objHoliday.BusyStatus = olOutOfOffice
    objHoliday.Save
Next

MsgBox("Your Outlook Calendar has been updated!")