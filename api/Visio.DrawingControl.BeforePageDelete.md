---
title: DrawingControl.BeforePageDelete event (Visio)
ms.prod: visio
api_name:
- Visio.DrawingControl.BeforePageDelete
ms.assetid: 18b80e01-a323-bce3-7107-1979f5ecc5a3
ms.date: 06/08/2017
localization_priority: Normal
---


# DrawingControl.BeforePageDelete event (Visio)

Occurs before a page is deleted.


## Syntax

_expression_.**BeforePageDelete** (_Page_)

_expression_ A variable that represents a **[DrawingControl](Visio.DrawingControl.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Page_|Required| **[IVPAGE]**|The page that is going to be deleted.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]