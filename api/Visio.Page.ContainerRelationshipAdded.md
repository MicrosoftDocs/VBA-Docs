---
title: Page.ContainerRelationshipAdded Event (Visio)
keywords: vis_sdr.chm10962065
f1_keywords:
- vis_sdr.chm10962065
ms.prod: visio
api_name:
- Visio.Page.ContainerRelationshipAdded
ms.assetid: 4cd95f23-baaa-3987-05f3-a379670efd02
ms.date: 06/08/2017
---


# Page.ContainerRelationshipAdded Event (Visio)

Occurs when a new container relationship is added to the document.


## Syntax

Private Sub  _expression_ _'ContainerRelationshipAdded'(**_By Val ShapePair As RelatedShapePairEvent_**)

 _expression_ A variable that represents a '[Page](Visio.Page.md)' object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ShapePair_|Required| **[RelatedShapePairEvent](Visio.RelatedShapePairEvent.md)**|An object that represents the new container shape-pair relationship.|

## Remarks

The  **RelatedShapePairEvent** object that this event returns contains two shapes: the container and the member, represented by the **[RelatedShapePairEvent.FromShapeID](Visio.RelatedShapePairEvent.FromShapeID.md)** and the **[RelatedShapePairEvent.ToShapeID](Visio.RelatedShapePairEvent.ToShapeID.md)** properties respectively.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **[Event](Visio.Event.md)** objects, use the **[EventList.Add](Visio.EventList.Add.md)** or **[EventList.AddAdvise](Visio.EventList.AddAdvise.md)** method. To create an **Event** object that runs an add-on, use the **EventList.Add** method. To create an **Event** object that receives notification, use the **EventList.AddAdvise** method. To find an event code for the event you want to create, see [Event Codes](../visio/Concepts/event-codesvisio.md).


