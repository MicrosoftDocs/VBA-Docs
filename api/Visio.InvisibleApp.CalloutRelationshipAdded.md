---
title: InvisibleApp.CalloutRelationshipAdded Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.CalloutRelationshipAdded
ms.assetid: dafc001b-d85e-416d-8f6e-5617969d9f15
ms.date: 06/08/2017
---


# InvisibleApp.CalloutRelationshipAdded Event (Visio)

Occurs when a new callout relationship is added to the application.


## Syntax

Private Sub  _expression_ _'CalloutRelationshipAdded'(**_By Val ShapePair As RelatedShapePairEvent_**)

 _expression_ A variable that represents an '[InvisibleApp](Visio.InvisibleApp.md)' object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ShapePair_|Required| **[RelatedShapePairEvent](Visio.RelatedShapePairEvent.md)**|An object that represents the new callout shape-pair relationship.|

## Remarks

The  **RelatedShapePairEvent** object returned by this event contains two shapes: the callout, represented by the **[FromShapeID](Visio.RelatedShapePairEvent.FromShapeID.md)** property of the **RelatedShapePairEvent** object; and the target shape, represented by the **[ToShapeID](Visio.RelatedShapePairEvent.ToShapeID.md)** property.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **[Event](Visio.Event.md)** objects, use the **[EventList.Add](Visio.EventList.Add.md)** or **[EventList.AddAdvise](Visio.EventList.AddAdvise.md)** method. To create an **Event** object that runs an add-on, use the **EventList.Add** method. To create an **Event** object that receives notification, use the **EventList.AddAdvise** method. To find an event code for the event you want to create, see [Event Codes](../visio/Concepts/event-codesvisio.md).


