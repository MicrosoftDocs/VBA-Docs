---
title: InvisibleApp.ContainerRelationshipAdded event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.ContainerRelationshipAdded
ms.assetid: 18d47470-f361-dc9d-2547-4b7bf81db453
ms.date: 06/25/2019
localization_priority: Normal
---


# InvisibleApp.ContainerRelationshipAdded event (Visio)

Occurs when a new container relationship is added to the document.


## Syntax

_expression_.**ContainerRelationshipAdded** (_ShapePair_)

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ShapePair_|Required| **[RelatedShapePairEvent](Visio.RelatedShapePairEvent.md)**|An object that represents the new container shape-pair relationship.|

## Remarks

The **RelatedShapePairEvent** object that this event returns contains two shapes: the container and the member, represented by the **[FromShapeID](Visio.RelatedShapePairEvent.FromShapeID.md)** and the **[ToShapeID](Visio.RelatedShapePairEvent.ToShapeID.md)** properties respectively.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]