---
title: Master.ShapeParentChanged event (Visio)
keywords: vis_sdr.chm10719235
f1_keywords:
- vis_sdr.chm10719235
ms.prod: visio
api_name:
- Visio.Master.ShapeParentChanged
ms.assetid: 37de7351-969b-5b24-fde2-e4473e92b344
ms.date: 06/08/2017
localization_priority: Normal
---


# Master.ShapeParentChanged event (Visio)

Occurs after shapes are grouped or a group is ungrouped.


## Syntax

_expression_.**ShapeParentChanged** (_Shape_)

_expression_ A variable that represents a **[Master](Visio.Master.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Shape_|Required| **[IVSHAPE]**|The shape whose parent changed.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]