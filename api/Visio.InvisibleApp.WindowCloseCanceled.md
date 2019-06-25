---
title: InvisibleApp.WindowCloseCanceled event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.WindowCloseCanceled
ms.assetid: 472119a3-c6e8-15bf-5266-6cfba9207ce5
ms.date: 06/08/2017
localization_priority: Normal
---


# InvisibleApp.WindowCloseCanceled event (Visio)

Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelWindowClose** event.


## Syntax

_expression_.**WindowCloseCanceled** (_Window_)

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window that was going to be closed.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]