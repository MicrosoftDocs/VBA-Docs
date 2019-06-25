---
title: InvisibleApp.BeforeSuspend event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.BeforeSuspend
ms.assetid: f7c84f82-6c44-2053-f28e-c4810d58eedf
ms.date: 06/25/2019
localization_priority: Normal
---


# InvisibleApp.BeforeSuspend event (Visio)

Occurs before the operating system enters a suspended state.


## Syntax

_expression_.**BeforeSuspend** (_app_)

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Visio that is going to be suspended.|

## Remarks

Client programs should close any open network files when this event is fired.

If your solution runs outside the Microsoft Visio process, you cannot be assured of receiving this event. For this reason, you should monitor window messages in your program.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]