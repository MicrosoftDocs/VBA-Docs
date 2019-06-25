---
title: InvisibleApp.AfterResumeEvents event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.AfterResumeEvents
ms.assetid: 33117394-135e-0f44-79e8-d16531cd0ca5
ms.date: 06/24/2019
localization_priority: Normal
---


# InvisibleApp.AfterResumeEvents event (Visio)

Occurs after firing of events is resumed.


## Syntax

_expression_.**AfterResumeEvents** (_app_)

_expression_ An expression that returns an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Microsoft Visio in which firing of events resumed.|

## Return value

**Nothing**


## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]