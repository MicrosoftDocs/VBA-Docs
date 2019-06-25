---
title: Documents.ShapeDataGraphicChanged event (Visio)
keywords: vis_sdr.chm10662010
f1_keywords:
- vis_sdr.chm10662010
ms.prod: visio
api_name:
- Visio.Documents.ShapeDataGraphicChanged
ms.assetid: 47fa7996-9b1f-d530-adf1-8dbb6d69aa5c
ms.date: 06/08/2017
localization_priority: Normal
---


# Documents.ShapeDataGraphicChanged event (Visio)

Occurs after a data graphic is applied to or deleted from a shape.


> [!NOTE] 
> This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

_expression_.**ShapeDataGraphicChanged** (_Shape_)

 _expression_ An expression that returns a **[Documents](Visio.Documents.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Shape_|Required| **[IVSHAPE]**|The shape to which the data graphic was applied or from which it was deleted.|

## Remarks

A data graphic is a  **Master** object of type **visTypeDataGraphic**. When the same master that is already applied to a shape is re-applied to the shape, the **ShapeDataGraphicChanged** event does not fire, even if the master has been modified since it was applied originally. If, however, a different data graphic master is applied to the shape, the event does fire.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]