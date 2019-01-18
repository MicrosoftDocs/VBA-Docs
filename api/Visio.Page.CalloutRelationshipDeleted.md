---
title: Page.CalloutRelationshipDeleted Event (Visio)
keywords: vis_sdr.chm10962080
f1_keywords:
- vis_sdr.chm10962080
ms.prod: visio
api_name:
- Visio.Page.CalloutRelationshipDeleted
ms.assetid: 06ab7df2-c2a9-2b86-4dd3-817f56dddf6c
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.CalloutRelationshipDeleted Event (Visio)

Occurs when a callout relationship is deleted from the page.


## Syntax

Private Sub  _expression_ _'CalloutRelationshipDeleted'(**_By Val ShapePair As RelatedShapePairEvent_**)

 _expression_ A variable that represents a '[Page](Visio.Page.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ShapePair_|Required| **[RelatedShapePairEvent](Visio.RelatedShapePairEvent.md)**|An object that represents the callout shape-pair relationship.|

## Remarks

The  **RelatedShapePairEvent** object that this event returns contains two shapes: the container and the member, represented by the **[RelatedShapePairEvent.FromShapeID](Visio.RelatedShapePairEvent.FromShapeID.md)** and the **[RelatedShapePairEvent.ToShapeID](Visio.RelatedShapePairEvent.ToShapeID.md)** properties respectively. When Microsoft Visio deletes both shapes in the container relationship simultaneously, for example when both shapes are deleted from a drawing as part of a selection, the event does not fire.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **[Event](Visio.Event.md)** objects, use the **[EventList.Add](Visio.EventList.Add.md)** or **[EventList.AddAdvise](Visio.EventList.AddAdvise.md)** method. To create an **Event** object that runs an add-on, use the **EventList.Add** method. To create an **Event** object that receives notification, use the **EventList.AddAdvise** method. To find an event code for the event that you want to create, see [Event Codes](../visio/Concepts/event-codesvisio.md).


