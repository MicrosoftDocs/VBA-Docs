---
title: Reminder.Dismiss method (Outlook)
keywords: vbaol11.chm558
f1_keywords:
- vbaol11.chm558
ms.prod: outlook
api_name:
- Outlook.Reminder.Dismiss
ms.assetid: cc757453-5eab-4e9f-5dd2-2b7620506d11
ms.date: 06/08/2017
localization_priority: Normal
---


# Reminder.Dismiss method (Outlook)

Dismisses the current reminder.


## Syntax

_expression_. `Dismiss`

_expression_ A variable that represents a [Reminder](Outlook.Reminder.md) object.


## Remarks

The  **Dismiss** method will fail if there is no visible reminder.


## Example

The following example dismisses all active reminders. A reminder is active if its  **[IsVisible](Outlook.Reminder.IsVisible.md)** property is set to **True**.


```vb
Sub DismissReminders() 
 
'Dismisses any active reminders. 
 
 
 
 Dim objRems As Outlook.Reminders 
 
 Dim objRem As Outlook.Reminder 
 
 Dim i As Integer 
 
 
 
 Set objRems = Application.Reminders 
 
 
 
 For i = objRems.Count To 1 Step -1 
 
 If objRems(i).IsVisible = True Then 
 
 objRems(i).Dismiss 
 
 End If 
 
 Next 
 
 Set olApp = Nothing 
 
 Set objRems = Nothing 
 
 Set objRem = Nothing 
 
End Sub
```


## See also


[Reminder Object](Outlook.Reminder.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]