---
title: Application.ShapeParentChanged Event (Visio)
ms.prod: visio
api_name:
- Visio.Application.ShapeParentChanged
ms.assetid: 321f937c-27e0-be80-9d6a-78e4e85629ec
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ShapeParentChanged Event (Visio)

Occurs after shapes are grouped or a group is ungrouped.


## Syntax

Private Sub  _expression_ _'ShapeParentChanged'(**_ByVal Shape As [IVSHAPE]_**)

 _expression_ A variable that represents an [Application](./Visio.Application.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Shape_|Required| **[IVSHAPE]**|The shape whose parent changed.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]