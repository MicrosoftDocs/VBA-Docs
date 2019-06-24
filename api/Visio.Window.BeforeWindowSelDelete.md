---
title: Window.BeforeWindowSelDelete event (Visio)
keywords: vis_sdr.chm11619085
f1_keywords:
- vis_sdr.chm11619085
ms.prod: visio
api_name:
- Visio.Window.BeforeWindowSelDelete
ms.assetid: 450bd22a-ceef-dcf4-90c0-b7511c3506dc
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.BeforeWindowSelDelete event (Visio)

Occurs before the shapes in the selection of a window are deleted.


## Syntax

_expression_.**BeforeWindowSelDelete** (_Window_)

_expression_ A variable that represents a **[Window](Visio.Window.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window that contains the selection that is going to be deleted.|

## Remarks

The  **BeforeWindowSelDelete** event fires if user interactions cause shapes in a window to be deleted. It doesn't fire if a program deletes shapes in a window by using the **Cut** method, for example.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]