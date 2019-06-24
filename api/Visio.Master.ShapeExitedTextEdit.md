---
title: Master.ShapeExitedTextEdit event (Visio)
keywords: vis_sdr.chm10719385
f1_keywords:
- vis_sdr.chm10719385
ms.prod: visio
api_name:
- Visio.Master.ShapeExitedTextEdit
ms.assetid: 401f6d32-d1fb-f019-52a3-d553b8516ecf
ms.date: 06/08/2017
localization_priority: Normal
---


# Master.ShapeExitedTextEdit event (Visio)

Occurs after a shape is no longer open for interactive text editing.


## Syntax

_expression_.**ShapeExitedTextEdit** (_Shape_)

_expression_ A variable that represents a **[Master](Visio.Master.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Shape_|Required| **[IVSHAPE]**|The shape that was closed for text editing.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]