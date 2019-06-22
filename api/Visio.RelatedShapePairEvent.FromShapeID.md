---
title: RelatedShapePairEvent.FromShapeID property (Visio)
keywords: vis_sdr.chm17762580
f1_keywords:
- vis_sdr.chm17762580
ms.prod: visio
api_name:
- Visio.RelatedShapePairEvent.FromShapeID
ms.assetid: d4f8c389-0a47-40e1-e60b-147daf789738
ms.date: 06/08/2017
localization_priority: Normal
---


# RelatedShapePairEvent.FromShapeID property (Visio)

Returns the identifier of the first (container or callout) shape in the related shape pair. Read-only.


## Syntax

_expression_. `FromShapeID`

_expression_ A variable that represents a **[RelatedShapePairEvent](Visio.RelatedShapePairEvent.md)** object.


## Return value

 **Long**


## Remarks

The first shape in the related shape pair is the container or callout shape. The second shape is the member shape.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]