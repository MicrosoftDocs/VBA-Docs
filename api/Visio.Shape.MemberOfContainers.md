---
title: Shape.MemberOfContainers property (Visio)
keywords: vis_sdr.chm11262465
f1_keywords:
- vis_sdr.chm11262465
ms.prod: visio
api_name:
- Visio.Shape.MemberOfContainers
ms.assetid: e8ed57eb-4031-5718-07ce-3641bda00186
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.MemberOfContainers property (Visio)

Returns an array of  **Long** values that represent the identifiers of the container shapes that include the shape as a member. Read-only.


## Syntax

_expression_. `MemberOfContainers`

_expression_ A variable that represents a **[Shape](Visio.Shape.md)** object.


## Return value

 **Long()**


## Remarks

The  **MemberOfContainers** property returns **Nothing** if the shape is not a member of any container.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]