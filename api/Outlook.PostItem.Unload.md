---
title: PostItem.Unload Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.PostItem.Unload
ms.assetid: 42dea931-c3dd-b8ff-5ace-0744b17650e0
ms.date: 06/08/2017
localization_priority: Normal
---


# PostItem.Unload Event (Outlook)

Occurs before an Outlook item is unloaded from memory, either programmatically or by user action. 


## Syntax

_expression_. `Unload`

_expression_ A variable that represents a [PostItem](./Outlook.PostItem.md) object.


## Remarks

This event occurs after the  **Close** event for the Outlook item occurs, but before the Outlook item is unloaded from memory, allowing an add-in to release any resources related to the object. Although the event occurs before the Outlook item is unloaded from memory, this event cannot be canceled.


 **Note**  This event is meant only as a notification event, so that an add-in can dereference the object. An error occurs if any property or method for this object is called within the  **Unload** event.


## See also


[PostItem Object](Outlook.PostItem.md)

