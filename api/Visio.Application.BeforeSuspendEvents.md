---
title: Application.BeforeSuspendEvents event (Visio)
ms.prod: visio
api_name:
- Visio.Application.BeforeSuspendEvents
ms.assetid: a6879424-40d8-e517-aad0-f31aa84a49f6
ms.date: 06/25/2019
localization_priority: Normal
---


# Application.BeforeSuspendEvents event (Visio)

Occurs before firing of events is suspended.


## Syntax

_expression_.**BeforeSuspendEvents** (_app_)

_expression_ An expression that returns an **[Application](Visio.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Microsoft Visio in which firing of events is going to be suspended.|

## Return value

**Nothing**


## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]