---
title: InvisibleApp.MustFlushScopeEnded event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.MustFlushScopeEnded
ms.assetid: de6fbf0d-25cf-e1e4-7d9f-e0a0e302e729
ms.date: 06/26/2019
localization_priority: Normal
---


# InvisibleApp.MustFlushScopeEnded event (Visio)

Occurs after the Microsoft Visio instance is forced to flush its event queue.


## Syntax

_expression_.**MustFlushScopeEnded** (_app_)

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Visio that is forced to flush its event queue.|

## Remarks

This event, along with the **MustFlushScopeBeginning** event, can be used to identify whether an event is being fired because Visio is forced to flush its event queue.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]