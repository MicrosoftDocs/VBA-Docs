---
title: DrawingControl.WindowTurnedToPage event (Visio)
ms.prod: visio
api_name:
- Visio.DrawingControl.WindowTurnedToPage
ms.assetid: 6cfab282-62db-75e6-9988-b6f9b4a0367e
ms.date: 06/08/2017
localization_priority: Normal
---


# DrawingControl.WindowTurnedToPage event (Visio)

Occurs after a window shows a different page.


## Syntax

_expression_.**WindowTurnedToPage** (_Window_)

_expression_ A variable that represents a **[DrawingControl](Visio.DrawingControl.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window that shows a different page.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]