---
title: InvisibleApp.SuspendCanceled event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.SuspendCanceled
ms.assetid: 5c266211-8686-85e8-f059-38e3cdab4211
ms.date: 06/08/2017
localization_priority: Normal
---


# InvisibleApp.SuspendCanceled event (Visio)

Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelSuspend** event.


## Syntax

_expression_.**SuspendCanceled** (_app_)

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Microsoft Visio that was going to be suspended.|

## Remarks

If your solution runs outside the Visio process, you cannot be assured of receiving this event. For this reason, you should monitor window messages in your program.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]