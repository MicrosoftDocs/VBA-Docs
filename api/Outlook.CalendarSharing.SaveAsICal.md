---
title: CalendarSharing.SaveAsICal method (Outlook)
keywords: vbaol11.chm2411
f1_keywords:
- vbaol11.chm2411
ms.prod: outlook
api_name:
- Outlook.CalendarSharing.SaveAsICal
ms.assetid: 2314f751-77c5-9b95-05fb-c3075f512508
ms.date: 06/08/2017
localization_priority: Normal
---


# CalendarSharing.SaveAsICal method (Outlook)

Exports calendar information from the parent  **[Folder](Outlook.Folder.md)** of the **[CalendarSharing](Outlook.CalendarSharing.md)** object as an iCalendar calendar (.ics) file.


## Syntax

_expression_. `SaveAsICal`( `_Path_` )

 _expression_ An expression that returns a [CalendarSharing](Outlook.CalendarSharing.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Path_|Required| **String**|The path and file name of the iCalendar file.|

## Remarks

The level of detail provided in the iCalendar file is determined by a combination of values in the following  **CalendarSharing** properties:


-  **[CalendarDetail](Outlook.CalendarSharing.CalendarDetail.md)**
    
-  **[IncludeAttachments](Outlook.CalendarSharing.IncludeAttachments.md)**
    
-  **[IncludePrivateDetails](Outlook.CalendarSharing.IncludePrivateDetails.md)**
    
-  **[RestrictToWorkingHours](Outlook.CalendarSharing.RestrictToWorkingHours.md)**
    
You can set the  **[IncludeWholeCalendar](Outlook.CalendarSharing.IncludeWholeCalendar.md)** property to **True** to export all items contained in the folder, or you can set the **[StartDate](Outlook.CalendarSharing.StartDate.md)** and **[EndDate](Outlook.CalendarSharing.EndDate.md)** properties to limit the exported items to a date range between a specified start date and end date, respectively.


## Example

The following Visual Basic for Applications (VBA) example creates a  **CalendarSharing** object for the Calendar folder, then exports the contents of the entire folder (including attachments and private items) to an iCalendar calendar (.ics) file.


```vb
Public Sub ExportEntireCalendar() 
 
 
 
 Dim oNamespace As NameSpace 
 
 Dim oFolder As Folder 
 
 Dim oCalendarSharing As CalendarSharing 
 
 
 
 On Error GoTo ErrRoutine 
 
 
 
 ' Get a reference to the Calendar default folder 
 
 Set oNamespace = Application.GetNamespace("MAPI") 
 
 Set oFolder = oNamespace.GetDefaultFolder(olFolderCalendar) 
 
 
 
 ' Get a CalendarSharing object for the Calendar default folder. 
 
 Set oCalendarSharing = oFolder.GetCalendarExporter 
 
 
 
 ' Set the CalendarSharing object to export the contents of 
 
 ' the entire Calendar folder, including attachments and 
 
 ' private items, in full detail. 
 
 With oCalendarSharing 
 
 .CalendarDetail = olFullDetails 
 
 .IncludeWholeCalendar = True 
 
 .IncludeAttachments = True 
 
 .IncludePrivateDetails = True 
 
 .RestrictToWorkingHours = False 
 
 End With 
 
 
 
 ' Export calendar to an iCalendar calendar (.ics) file. 
 
 oCalendarSharing.SaveAsICal "C:\SampleCalendar.ics" 
 
 
 
EndRoutine: 
 
 On Error GoTo 0 
 
 Set oCalendarSharing = Nothing 
 
 Set oFolder = Nothing 
 
 Set oNamespace = Nothing 
 
Exit Sub 
 
 
 
ErrRoutine: 
 
 Select Case Err.Number 
 
 Case 287 ' &H0000011F 
 
 ' The user denied access to the Address Book. 
 
 ' This error occurs if the code is run by an 
 
 ' untrusted application, and the user chose not to 
 
 ' allow access. 
 
 MsgBox "Access to Outlook was denied by the user.", _ 
 
 vbOKOnly, _ 
 
 Err.Number & " - " & Err.Source 
 
 Case -2147467259 ' &H80004005 
 
 ' Export failed. 
 
 ' This error typically occurs if the CalendarSharing 
 
 ' method cannot export the calendar information because 
 
 ' of conflicting property settings. 
 
 MsgBox Err.Description, _ 
 
 vbOKOnly, _ 
 
 Err.Number & " - " & Err.Source 
 
 Case -2147221233 ' &H8004010F 
 
 ' Operation failed. 
 
 ' This error typically occurs if the GetCalendarExporter method 
 
 ' is called on a folder that doesn't contain calendar items. 
 
 MsgBox Err.Description, _ 
 
 vbOKOnly, _ 
 
 Err.Number & " - " & Err.Source 
 
 Case Else 
 
 ' Any other error that may occur. 
 
 MsgBox Err.Description, _ 
 
 vbOKOnly, _ 
 
 Err.Number & " - " & Err.Source 
 
 End Select 
 
 
 
 GoTo EndRoutine 
 
End Sub
```


## See also


[CalendarSharing Object](Outlook.CalendarSharing.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]