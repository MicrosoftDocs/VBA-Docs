---
title: Application.AppObjActivated event (Visio)
ms.prod: visio
api_name:
- Visio.Application.AppObjActivated
ms.assetid: ab27fad1-5afb-534c-987f-e5401603aa52
ms.date: 06/24/2019
localization_priority: Normal
---


# Application.AppObjActivated event (Visio)

Occurs after a Microsoft Visio instance becomes active.


## Syntax

_expression_.**AppObjActivated** (_app_)

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Visio that has become the active instance.|

## Remarks

The **AppObjActivated** event indicates that an instance of Visio has become active (the instance of Visio that is retrieved by the **GetObject** method in a Microsoft Visual Basic program). The **AppObjActivated** event is different from the **AppActivated** event, which occurs after an instance of Visio becomes the active application on the Microsoft Windows desktop.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]