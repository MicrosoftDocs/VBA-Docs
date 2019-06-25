---
title: DrawingControl.BeforeStyleDelete event (Visio)
ms.prod: visio
api_name:
- Visio.DrawingControl.BeforeStyleDelete
ms.assetid: 7b6fb188-d625-3133-f7c0-2f0c55dfe083
ms.date: 06/08/2017
localization_priority: Normal
---


# DrawingControl.BeforeStyleDelete event (Visio)

Occurs before a style is deleted.


## Syntax

_expression_.**BeforeStyleDelete** (_Style_)

_expression_ A variable that represents a **[DrawingControl](Visio.DrawingControl.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Style_|Required| **[IVSTYLE]**|The style that is going to be deleted.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]