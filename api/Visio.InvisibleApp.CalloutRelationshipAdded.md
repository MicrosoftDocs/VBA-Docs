---
title: InvisibleApp.CalloutRelationshipAdded event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.CalloutRelationshipAdded
ms.assetid: dafc001b-d85e-416d-8f6e-5617969d9f15
ms.date: 06/25/2019
localization_priority: Normal
---


# InvisibleApp.CalloutRelationshipAdded event (Visio)

Occurs when a new callout relationship is added to the application.


## Syntax

_expression_.**CalloutRelationshipAdded** (_ShapePair_)

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ShapePair_|Required| **[RelatedShapePairEvent](Visio.RelatedShapePairEvent.md)**|An object that represents the new callout shape-pair relationship.|

## Remarks

The **RelatedShapePairEvent** object returned by this event contains two shapes: the callout, represented by the **[FromShapeID](Visio.RelatedShapePairEvent.FromShapeID.md)** property of the **RelatedShapePairEvent** object, and the target shape, represented by the **[ToShapeID](Visio.RelatedShapePairEvent.ToShapeID.md)** property.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]