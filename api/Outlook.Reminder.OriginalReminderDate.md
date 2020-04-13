---
title: Reminder.OriginalReminderDate property (Outlook)
keywords: vbaol11.chm564
f1_keywords:
- vbaol11.chm564
ms.prod: outlook
api_name:
- Outlook.Reminder.OriginalReminderDate
ms.assetid: ecc3f0c4-0e20-1d02-94b5-40807523ad2d
ms.date: 06/08/2017
localization_priority: Normal
---


# Reminder.OriginalReminderDate property (Outlook)

Returns a **Date** that specifies the original date and time that the specified reminder is set to occur. Read-only.


## Syntax

_expression_. `OriginalReminderDate`

_expression_ A variable that represents a [Reminder](Outlook.Reminder.md) object.


## Remarks

This value corresponds to the original date and time value before the  **[Snooze](Outlook.Reminder.Snooze.md)** method is executed or the user clicks the **Snooze** button.


## Example

The following Microsoft Visual Basic for Applications (VBA) example creates a report of all reminders in the  **[Reminders](Outlook.Reminders.md)** collection and the dates at which they are scheduled to occur. The subroutine concatenates the **[Caption](Outlook.Reminder.Caption.md)** and **OriginalReminderDate** properties of all **[Reminder](Outlook.Reminder.md)** objects in the collection into a string and displays the string in a dialog box.


```vb
Sub DisplayOriginalDateReport() 
 
 'Displays the time at which all reminders will be displayed. 
 
 Dim objRems As Outlook.Reminders 
 
 Dim objRem As Outlook.Reminder 
 
 Dim strTitle As String 
 
 Dim strReport As String 
 
 
 
 Set objRems = Application.Reminders 
 
 strTitle = "Original Reminder Schedule:" 
 
 strReport = "" 
 
 'Check if any reminders exist. 
 
 If objRems.Count = 0 Then 
 
 MsgBox "There are no current reminders." 
 
 Else 
 
 For Each objRem In objRems 
 
 'Add info to string 
 
 strReport = strReport & objRem.Caption & vbTab & vbTab & _ 
 
 objRem.OriginalReminderDate & vbCr 
 
 Next objRem 
 
 'Display report in dialog 
 
 MsgBox strTitle & vbCr & vbCr & strReport 
 
 End If 
 
End Sub
```


## See also


[Reminder Object](Outlook.Reminder.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]