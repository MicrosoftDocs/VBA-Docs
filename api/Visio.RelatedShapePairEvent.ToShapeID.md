---
title: RelatedShapePairEvent.ToShapeID property (Visio)
keywords: vis_sdr.chm17762585
f1_keywords:
- vis_sdr.chm17762585
ms.prod: visio
api_name:
- Visio.RelatedShapePairEvent.ToShapeID
ms.assetid: cdf61ad1-244e-5605-225b-4f919c923af8
ms.date: 06/08/2017
localization_priority: Normal
---


# RelatedShapePairEvent.ToShapeID property (Visio)

Returns the identifier of the second (member) shape in the related shape pair. Read-only.


## Syntax

_expression_. `ToShapeID`

_expression_ A variable that represents a **[RelatedShapePairEvent](Visio.RelatedShapePairEvent.md)** object.


## Return value

 **Long**


## Remarks

The first shape in the related shape pair is the container or callout shape. The second shape is the member shape.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]