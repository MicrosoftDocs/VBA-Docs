---
title: Pages.CalloutRelationshipAdded event (Visio)
keywords: vis_sdr.chm11062075
f1_keywords:
- vis_sdr.chm11062075
ms.prod: visio
api_name:
- Visio.Pages.CalloutRelationshipAdded
ms.assetid: 45f350ca-05ed-b775-d5da-b0d9c8a5c885
ms.date: 06/08/2017
localization_priority: Normal
---


# Pages.CalloutRelationshipAdded event (Visio)

Occurs when a new callout relationship is added to a page.


## Syntax

_expression_.**CalloutRelationshipAdded** (_ShapePair_)

_expression_ A variable that represents a **[Pages](Visio.Pages.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ShapePair_|Required| **[RelatedShapePairEvent](Visio.RelatedShapePairEvent.md)**|An object that represents the new callout shape-pair relationship.|

## Remarks

The  **RelatedShapePairEvent** object that this event returns contains two shapes: the container and the member, represented by the **[RelatedShapePairEvent.FromShapeID](Visio.RelatedShapePairEvent.FromShapeID.md)** and the **[RelatedShapePairEvent.ToShapeID](Visio.RelatedShapePairEvent.ToShapeID.md)** properties respectively.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]