---
title: InvisibleApp.DeferRelationshipRecalc property (Visio)
keywords: vis_sdr.chm17562425
f1_keywords:
- vis_sdr.chm17562425
ms.prod: visio
api_name:
- Visio.InvisibleApp.DeferRelationshipRecalc
ms.assetid: 77c7842c-1dc0-fea9-2cdc-0381aab251d2
ms.date: 06/25/2019
localization_priority: Normal
---


# InvisibleApp.DeferRelationshipRecalc property (Visio)

Determines whether Microsoft Visio defers recalculating shape sizes and relationships when a member of a relationship pair is moved or resized. Read/write.


## Syntax

_expression_.**DeferRelationshipRecalc**

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Return value

**Boolean**


## Remarks

For example, if you resize a shape that is a member of a container in a structured diagram, Visio will not adjust the size of the container if **DeferRelationshipRecalc** is **True**. 

When you set **DeferRelationshipRecalc** to **False**, Visio recalculates the container size and adjusts it accordingly. (In each case, the container's **[ResizeAsNeeded](Visio.ContainerProperties.ResizeAsNeeded.md)** property must be set to **visContainerAutoResizeExpandContract**.)

Setting **DeferRelationshipRecalc** to **False** causes Visio to immediately process all deferred actions.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]