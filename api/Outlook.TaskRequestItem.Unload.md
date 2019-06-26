---
title: TaskRequestItem.Unload event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem.Unload
ms.assetid: 9a8d0aad-8d39-32ae-945e-757c346e80d6
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestItem.Unload event (Outlook)

Occurs before an Outlook item is unloaded from memory, either programmatically or by user action. 


## Syntax

_expression_. `Unload`

_expression_ A variable that represents a [TaskRequestItem](Outlook.TaskRequestItem.md) object.


## Remarks

This event occurs after the  **Close** event for the Outlook item occurs, but before the Outlook item is unloaded from memory, allowing an add-in to release any resources related to the object. Although the event occurs before the Outlook item is unloaded from memory, this event cannot be canceled.


> [!NOTE] 
> This event is meant only as a notification event, so that an add-in can dereference the object. An error occurs if any property or method for this object is called within the  **Unload** event.


## See also


[TaskRequestItem Object](Outlook.TaskRequestItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]