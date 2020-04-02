---
title: RelatedShapePairEvent object (Visio)
keywords: vis_sdr.chm61045
f1_keywords:
- vis_sdr.chm61045
ms.prod: visio
api_name:
- Visio.RelatedShapePairEvent
ms.assetid: 8a59ae03-ed45-21e3-73ad-8fdbe4c53299
ms.date: 06/19/2019
localization_priority: Normal
---


# RelatedShapePairEvent object (Visio)

Holds information about the shapes that are involved in a container relationship or a callout relationship.


## Remarks

A related shape pair consists of two shapes—typically a container and a member, or a callout and a target shape.

When you add or remove a member shape from a container, Microsoft Visio fires a **ContainerRelationshipAdded** or **ContainerRelationshipDeleted** event respectively, and in each case, returns a **RelatedShapePairEvent** object that encapsulates information about the event.

When you add or remove a callout relationship from a document, Microsoft Visio fires a **CalloutRelationshipAdded** or **CalloutRelationshipDeleted** event respectively, and in each case, returns a **RelatedShapePairEvent** object that encapsulates information about the event.

## Properties

- [ContainingPage](Visio.RelatedShapePairEvent.ContainingPage.md)
- [ContainingPageID](Visio.RelatedShapePairEvent.ContainingPageID.md)
- [Document](Visio.RelatedShapePairEvent.Document.md)
- [FromShapeID](Visio.RelatedShapePairEvent.FromShapeID.md)
- [ObjectType](Visio.RelatedShapePairEvent.ObjectType.md)
- [Stat](Visio.RelatedShapePairEvent.Stat.md)
- [ToShapeID](Visio.RelatedShapePairEvent.ToShapeID.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]