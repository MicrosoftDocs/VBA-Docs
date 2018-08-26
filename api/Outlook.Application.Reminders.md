---
title: Application.Reminders Property (Outlook)
keywords: vbaol11.chm731
f1_keywords:
- vbaol11.chm731
ms.prod: outlook
api_name:
- Outlook.Application.Reminders
ms.assetid: 1f5428f0-6362-a691-2fad-c80e48dce3f5
ms.date: 06/08/2017
---


# Application.Reminders Property (Outlook)

Returns a  **[Reminders](Outlook.Reminders.md)** collection that represents all current reminders. Read-only.


## Syntax

 _expression_. `Reminders`

 _expression_ A variable that represents an [Application](./Outlook.Application.md) object.


## Example

The following example returns the  **Reminders** collection and displays the captions of all reminders in the collection. If no current reminders are available, a message is displayed to the user.


```vb
Sub ViewReminderInfo() 
 
 'Lists reminder caption information 
 
 Dim objRem As Outlook.Reminder 
 
 Dim objRems As Outlook.Reminders 
 
 Dim strTitle As String 
 
 Dim strReport As String 
 
 
 
 Set objRems = Application.Reminders 
 
 strTitle = "Current Reminders:" 
 
 strReport = "" 
 
 'If there are reminders, display message 
 
 If Application.Reminders.Count <> 0 Then 
 
 For Each objRem In objRems 
 
 'Add information to string 
 
<<<<<<< HEAD
 strReport = strReport &; objRem.Caption &; vbCr 
=======
 strReport = strReport & objRem.Caption & vbCr 
>>>>>>> master
 
 Next objRem 
 
 'Display report in dialog 
 
<<<<<<< HEAD
 MsgBox strTitle &; vbCr &; vbCr &; strReport 
=======
 MsgBox strTitle & vbCr & vbCr & strReport 
>>>>>>> master
 
 Else 
 
 MsgBox "There are no reminders in the collection." 
 
 End If 
 
End Sub
```


## See also


[Application Object](Outlook.Application.md)

