---
title: DrawingControl.BeforeWindowClosed event (Visio)
ms.prod: visio
api_name:
- Visio.DrawingControl.BeforeWindowClosed
ms.assetid: eb253857-42a1-3371-484d-e01fea21c8cc
ms.date: 06/08/2017
localization_priority: Normal
---


# DrawingControl.BeforeWindowClosed event (Visio)

Occurs before a window is closed.


## Syntax

_expression_.**BeforeWindowClosed** (_Window_)

_expression_ A variable that represents a **[DrawingControl](Visio.DrawingControl.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window that is going to be closed.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]