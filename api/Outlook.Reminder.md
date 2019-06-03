---
title: Reminder object (Outlook)
keywords: vbaol11.chm3014
f1_keywords:
- vbaol11.chm3014
ms.prod: outlook
api_name:
- Outlook.Reminder
ms.assetid: b7364e48-51bc-b360-2154-e85e7779ece4
ms.date: 06/08/2017
localization_priority: Normal
---


# Reminder object (Outlook)

Represents an Outlook reminder.


## Remarks

Reminders allow users to keep track of upcoming appointments by scheduling a pop-up dialog box to appear at a given time. In addition to appointments, reminders can occur for tasks, contacts and email messages.

Use  **[Reminders](Outlook.Application.Reminders.md)** (_index_), where _index_ is the name or index number of the reminder, to return a single **Reminder** object.

Reminders are created programmatically when a new Microsoft Outlook item, such as an  **[AppointmentItem](Outlook.AppointmentItem.md)** object, is created and the item 's **[ReminderSet](Outlook.AppointmentItem.ReminderSet.md)** property is set to **True**.

Use the  **Reminders** collection's **[Remove](Outlook.Reminders.Remove.md)** method to remove a **Reminder** object from the collection. Once a reminder is removed from its associated item, the **AppointmentItem** object's **ReminderSet** property is set to **False**.


## Example

The following example displays the caption of the first reminder in the collection.


```vb
Sub ViewReminderInfo() 
 
 'Displays information about first reminder in collection 
 
 
 
 Dim colReminders As Outlook.Reminders 
 
 Dim objRem As Reminder 
 
 
 
 Set colReminders = Application.Reminders 
 
 'If there are reminders, display message 
 
 If colReminders.Count <> 0 Then 
 
 Set objRem = colReminders.Item(1) 
 
 MsgBox "The caption of the first reminder in the collection is: " & _ 
 
 objRem.Caption 
 
 Else 
 
 MsgBox "There are no reminders in the collection." 
 
 
 
 End If 
 
 
 
End Sub
```

The following example creates a new appointment item and sets the  **ReminderSet** property to **True**, adding a new **Reminder** object to the **Reminders** collection.




```vb
Sub AddAppt() 
 
 'Adds a new appointment and reminder to the reminders collection 
 
 Dim objApt As AppointmentItem 
 
 
 
 Set objApt = Application.CreateItem(olAppointmentItem) 
 
 objApt.ReminderSet = True 
 
 objApt.Subject = "Tuesday's meeting" 
 
 objApt.Save 
 
End Sub
```


## Methods



|Name|
|:-----|
|[Dismiss](Outlook.Reminder.Dismiss.md)|
|[Snooze](Outlook.Reminder.Snooze.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.Reminder.Application.md)|
|[Caption](Outlook.Reminder.Caption.md)|
|[Class](Outlook.Reminder.Class.md)|
|[IsVisible](Outlook.Reminder.IsVisible.md)|
|[Item](Outlook.Reminder.Item.md)|
|[NextReminderDate](Outlook.Reminder.NextReminderDate.md)|
|[OriginalReminderDate](Outlook.Reminder.OriginalReminderDate.md)|
|[Parent](Outlook.Reminder.Parent.md)|
|[Session](Outlook.Reminder.Session.md)|

## See also


[Reminder Object Members](overview/Outlook.md)
[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]