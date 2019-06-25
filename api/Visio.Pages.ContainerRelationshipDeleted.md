---
title: Pages.ContainerRelationshipDeleted event (Visio)
keywords: vis_sdr.chm11062070
f1_keywords:
- vis_sdr.chm11062070
ms.prod: visio
api_name:
- Visio.Pages.ContainerRelationshipDeleted
ms.assetid: ed72e2e1-00c8-9ae0-eb53-57fe75035345
ms.date: 06/08/2017
localization_priority: Normal
---


# Pages.ContainerRelationshipDeleted event (Visio)

Occurs when a container relationship is deleted from the document.


## Syntax

_expression_.**ContainerRelationshipDeleted** (_ShapePair_)

_expression_ A variable that represents a **[Pages](Visio.Pages.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ShapePair_|Required| **[RelatedShapePairEvent](Visio.RelatedShapePairEvent.md)**|An object that represents the container shape-pair relationship.|

## Remarks

The **RelatedShapePairEvent** object that this event returns contains two shapes: the container and the member, represented by the **[FromShapeID](Visio.RelatedShapePairEvent.FromShapeID.md)** and the **[ToShapeID](Visio.RelatedShapePairEvent.ToShapeID.md)** properties respectively. When Microsoft Visio deletes both shapes in the container relationship simultaneously (for example, when both shapes are deleted from a drawing as part of a selection), the event does not fire.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]