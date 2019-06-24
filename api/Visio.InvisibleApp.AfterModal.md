---
title: InvisibleApp.AfterModal event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.AfterModal
ms.assetid: 52c6926d-0b94-8425-729c-9494f5c7879e
ms.date: 06/24/2019
localization_priority: Normal
---


# InvisibleApp.AfterModal event (Visio)

Occurs after the Microsoft Visio instance leaves a modal state.


## Syntax

_expression_.**AfterModal** (_app_)

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance that is no longer modal.|

## Remarks

Visio becomes modal when it displays a dialog box. A modal instance of Visio does not handle Automation calls. The **BeforeModal** event indicates that the instance is about to become modal, and the **AfterModal** event indicates that the instance is no longer modal.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]