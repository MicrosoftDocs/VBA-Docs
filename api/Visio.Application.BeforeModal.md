---
title: Application.BeforeModal event (Visio)
ms.prod: visio
api_name:
- Visio.Application.BeforeModal
ms.assetid: 505d3e54-c8f7-7f02-90d2-43f73573b296
ms.date: 06/25/2019
localization_priority: Normal
---


# Application.BeforeModal event (Visio)

Occurs before a Microsoft Visio instance enters a modal state.


## Syntax

_expression_.**BeforeModal** (_app_)

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Visio that is going to enter a modal state.|

## Remarks

Visio becomes modal when it displays a dialog box. A modal instance of Visio does not handle Automation calls. The **BeforeModal** event indicates that an instance is about to become modal, and the **AfterModal** event indicates that the instance is no longer modal.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]