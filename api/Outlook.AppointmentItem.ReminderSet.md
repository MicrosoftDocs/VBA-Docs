---
title: AppointmentItem.ReminderSet property (Outlook)
keywords: vbaol11.chm895
f1_keywords:
- vbaol11.chm895
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.ReminderSet
ms.assetid: 575d5fb2-1672-ddae-832c-7dcc7d1da2d6
ms.date: 06/08/2017
localization_priority: Normal
---


# AppointmentItem.ReminderSet property (Outlook)

Returns or sets a **Boolean** value that is **True** if a reminder has been set for this item. Read/write.


## Syntax

_expression_. `ReminderSet`

_expression_ A variable that represents an [AppointmentItem](Outlook.AppointmentItem.md) object.


## Example

This example creates an appointment item and sets the  **ReminderSet** property before saving it.


```vb
Sub AddAppointment() 
 
 Dim apti As Outlook.AppointmentItem 
 
 
 
 Set apti = Application.CreateItem(olAppointmentItem) 
 
 apti.Subject = "Car Servicing" 
 
 apti.Start = DateAdd("n", 16, Now) 
 
 apti.End = DateAdd("n", 60, apti.Start) 
 
 apti.ReminderSet = True 
 
 apti.ReminderMinutesBeforeStart = 60 
 
 apti.Save 
 
End Sub
```


## See also


[AppointmentItem Object](Outlook.AppointmentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]