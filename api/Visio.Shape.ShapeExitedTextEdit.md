---
title: Shape.ShapeExitedTextEdit Event (Visio)
keywords: vis_sdr.chm11219385
f1_keywords:
- vis_sdr.chm11219385
ms.prod: visio
api_name:
- Visio.Shape.ShapeExitedTextEdit
ms.assetid: ba707fd6-2a5a-65f6-6db4-ed3b5250a103
ms.date: 06/08/2017
---


# Shape.ShapeExitedTextEdit Event (Visio)

Occurs after a shape is no longer open for interactive text editing.


## Syntax

Private Sub  _expression_ _'ShapeExitedTextEdit'(**_ByVal Shape As [IVSHAPE]_**)

 _expression_ A variable that represents a [Shape](./Visio.Shape.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Shape_|Required| **[IVSHAPE]**|The shape that was closed for text editing.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


