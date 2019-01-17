---
title: InvisibleApp.ContainerRelationshipDeleted Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.ContainerRelationshipDeleted
ms.assetid: 689cb7e6-48a4-6438-ba9d-e1b554ac0bca
ms.date: 06/08/2017
localization_priority: Normal
---


# InvisibleApp.ContainerRelationshipDeleted Event (Visio)

Occurs when a container relationship is deleted from the document.


## Syntax

Private Sub  _expression_ _'ContainerRelationshipDeleted'(**_By Val ShapePair As RelatedShapePairEvent_**)

 _expression_ A variable that represents an '[InvisibleApp](Visio.InvisibleApp.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ShapePair_|Required| **[RelatedShapePairEvent](Visio.RelatedShapePairEvent.md)**|An object that represents the container shape-pair relationship.|

## Remarks

The  **RelatedShapePairEvent** object that this event returns contains two shapes: the container and the member, represented by the **[RelatedShapePairEvent.FromShapeID](Visio.RelatedShapePairEvent.FromShapeID.md)** and the **[RelatedShapePairEvent.ToShapeID](Visio.RelatedShapePairEvent.ToShapeID.md)** properties respectively. When Microsoft Visio deletes both shapes in the container relationship simultaneously, for example when both shapes are deleted from a drawing as part of a selection, the event does not fire.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **[Event](Visio.Event.md)** objects, use the **[EventList.Add](Visio.EventList.Add.md)** or **[EventList.AddAdvise](Visio.EventList.AddAdvise.md)** method. To create an **Event** object that runs an add-on, use the **EventList.Add** method. To create an **Event** object that receives notification, use the **EventList.AddAdvise** method. To find an event code for the event you want to create, see [Event Codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]