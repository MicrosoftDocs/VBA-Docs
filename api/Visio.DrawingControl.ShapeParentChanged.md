---
title: DrawingControl.ShapeParentChanged Event (Visio)
ms.prod: visio
api_name:
- Visio.DrawingControl.ShapeParentChanged
ms.assetid: f3e0a789-bb27-10b4-f363-effc8f480a54
ms.date: 06/08/2017
---


# DrawingControl.ShapeParentChanged Event (Visio)

Occurs after shapes are grouped or a group is ungrouped.


## Syntax

Private Sub  _expression_ _'ShapeParentChanged'(**_ByVal shape As [IVSHAPE]_**)

 _expression_ A variable that represents a [DrawingControl](./Visio.DrawingControl.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Shape_|Required| **[IVSHAPE]**|The shape whose parent changed.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


