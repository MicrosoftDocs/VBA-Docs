---
title: DrawingControl.BeforeShapeTextEdit Event (Visio)
ms.prod: visio
api_name:
- Visio.DrawingControl.BeforeShapeTextEdit
ms.assetid: a499b10a-3163-b734-91b1-5985613712d0
ms.date: 06/08/2017
localization_priority: Normal
---


# DrawingControl.BeforeShapeTextEdit Event (Visio)

Occurs before a shape is opened for text editing in the user interface.


## Syntax

Private Sub  _expression_ _'BeforeShapeTextEdit'(**_ByVal Shape As [IVSHAPE]_**)

 _expression_ A variable that represents a [DrawingControl](./Visio.DrawingControl.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Shape_|Required| **[IVSHAPE]**|The shape that is going to be opened for text editing.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


