---
title: Shape.CalloutsAssociated property (Visio)
keywords: vis_sdr.chm11262480
f1_keywords:
- vis_sdr.chm11262480
ms.prod: visio
api_name:
- Visio.Shape.CalloutsAssociated
ms.assetid: c1e32bb2-c946-3919-4f6e-064b5be50d6c
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.CalloutsAssociated property (Visio)

Returns an array of  **Long** values that represent the collection of identifiers for all of the callout shapes that are associated with the target shape by a callout relationship. Read-only.


## Syntax

_expression_. `CalloutsAssociated`

_expression_ A variable that represents a **[Shape](Visio.Shape.md)** object.


## Return value

 **Long()**


## Remarks

If there are no callouts associated with the target shape, the  **CalloutsAssociated** property returns an empty array.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]