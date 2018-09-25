---
title: Documents.ContainerRelationshipAdded Event (Visio)
keywords: vis_sdr.chm10662065
f1_keywords:
- vis_sdr.chm10662065
ms.prod: visio
api_name:
- Visio.Documents.ContainerRelationshipAdded
ms.assetid: 8cc692b3-f079-0c9a-b2fe-f0e3fac83ec6
ms.date: 06/08/2017
---


# Documents.ContainerRelationshipAdded Event (Visio)

Occurs when a new container relationship is added to the document.


## Syntax

Private Sub  _expression_ _'ContainerRelationshipAdded'(**_By Val ShapePair As RelatedShapePairEvent_**)

 _expression_ A variable that represents a '[Documents](Visio.Documents.md)' object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ShapePair_|Required| **[RelatedShapePairEvent](Visio.RelatedShapePairEvent.md)**|An object that represents the new container shape-pair relationship.|

## Remarks

The  **RelatedShapePairEvent** object that this event returns contains two shapes: the container and the member, represented by the **[RelatedShapePairEvent.FromShapeID](Visio.RelatedShapePairEvent.FromShapeID.md)** and the **[RelatedShapePairEvent.ToShapeID](Visio.RelatedShapePairEvent.ToShapeID.md)** properties respectively.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **[Event](Visio.Event.md)** objects, use the **[EventList.Add](Visio.EventList.Add.md)** or **[EventList.AddAdvise](Visio.EventList.AddAdvise.md)** method. To create an **Event** object that runs an add-on, use the **EventList.Add** method. To create an **Event** object that receives notification, use the **EventList.AddAdvise** method. To find an event code for the event you want to create, see [Event Codes](../visio/Concepts/event-codesvisio.md).


