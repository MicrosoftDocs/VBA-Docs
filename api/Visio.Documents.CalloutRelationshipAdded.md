---
title: Documents.CalloutRelationshipAdded event (Visio)
keywords: vis_sdr.chm10662075
f1_keywords:
- vis_sdr.chm10662075
ms.prod: visio
api_name:
- Visio.Documents.CalloutRelationshipAdded
ms.assetid: c88e5921-60ea-cf87-a630-ac857f1159a8
ms.date: 06/08/2017
localization_priority: Normal
---


# Documents.CalloutRelationshipAdded event (Visio)

Occurs when a new callout relationship is added to a document.


## Syntax

_expression_.**CalloutRelationshipAdded** (_ShapePair_)

_expression_ A variable that represents a **[Documents](Visio.Documents.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ShapePair_|Required| **[RelatedShapePairEvent](Visio.RelatedShapePairEvent.md)**|An object that represents the new callout shape-pair relationship.|

## Remarks

The  **RelatedShapePairEvent** object that this event returns contains two shapes: the container and the member, represented by the **[RelatedShapePairEvent.FromShapeID](Visio.RelatedShapePairEvent.FromShapeID.md)** and the **[RelatedShapePairEvent.ToShapeID](Visio.RelatedShapePairEvent.ToShapeID.md)** properties respectively.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **[Event](Visio.Event.md)** objects, use the **[EventList.Add](Visio.EventList.Add.md)** or **[EventList.AddAdvise](Visio.EventList.AddAdvise.md)** method. To create an **Event** object that runs an add-on, use the **EventList.Add** method. To create an **Event** object that receives notification, use the **EventList.AddAdvise** method. To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]