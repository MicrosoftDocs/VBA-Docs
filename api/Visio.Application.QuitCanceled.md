---
title: Application.QuitCanceled event (Visio)
ms.prod: visio
api_name:
- Visio.Application.QuitCanceled
ms.assetid: 0861a2ea-f4d7-dc57-7642-2e7642fd2afe
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.QuitCanceled event (Visio)

Occurs after an event handler has returned **True** (cancel) to a **QueryCancelQuit** event.


## Syntax

_expression_.**QuitCanceled** (_app_)

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Microsoft Visio that was going to be terminated.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]