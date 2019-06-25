---
title: Application.BeforeWindowPageTurn event (Visio)
ms.prod: visio
api_name:
- Visio.Application.BeforeWindowPageTurn
ms.assetid: ddb79c04-7cb4-61fd-a37d-d5969e445d5c
ms.date: 06/25/2019
localization_priority: Normal
---


# Application.BeforeWindowPageTurn event (Visio)

Occurs before a window is about to show a different page.


## Syntax

_expression_.**BeforeWindowPageTurn** (_Window_)

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window that is going to show a different page.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]