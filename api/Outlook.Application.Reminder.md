---
title: Application.Reminder event (Outlook)
keywords: vbaol11.chm431
f1_keywords:
- vbaol11.chm431
ms.prod: outlook
api_name:
- Outlook.Application.Reminder
ms.assetid: f8c9fa87-3daa-58e1-7b8d-3c819cd4cab2
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Reminder event (Outlook)

Occurs immediately before a reminder is displayed.


## Syntax

_expression_. `Reminder`( `_Item_` )

_expression_ A variable that represents an **[Application](Outlook.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|The  **[AppointmentItem](Outlook.AppointmentItem.md)**, **[MailItem](Outlook.MailItem.md)**, **[ContactItem](Outlook.ContactItem.md)**, or **[TaskItem](Outlook.TaskItem.md)** associated with the reminder. If the appointment associated with the reminder is a recurring appointment, _Item_ is the specific occurrence of the appointment that displayed the reminder, not the master appointment.|

## Example

This Microsoft Visual Basic for Applications (VBA) example displays the item that fired the  **Reminder** event when the event fires. The sample code must be placed in a class module, and the `Initialize_handler` routine must be called before the event procedure can be called by Microsoft Outlook.


```vb
Dim WithEvents myolapp As Outlook.Application 
 
 
 
Sub Initialize_handler() 
 
 Set myolapp = Outlook.Application 
 
End Sub 
 
 
 
Private Sub myolapp_Reminder(ByVal Item As Object) 
 
 Item.Display 
 
End Sub
```


## See also


[Application Object](Outlook.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]