---
title: Shape.ContainingPageID property (Visio)
keywords: vis_sdr.chm11260135
f1_keywords:
- vis_sdr.chm11260135
ms.prod: visio
api_name:
- Visio.Shape.ContainingPageID
ms.assetid: fd33d0d6-571d-47b5-28a7-6fa4aa671312
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.ContainingPageID property (Visio)

Returns the ID of the page that contains an object. Read-only.


## Syntax

_expression_. `ContainingPageID`

_expression_ A variable that represents a **[Shape](Visio.Shape.md)** object.


## Return value

Long


## Remarks

If the object is not in a  **Page** object, the **ContainingPageID** property returns -1. For example, if a **Shape** object belongs to a **Masters** collection, the **ContainingPageID** property returns -1.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]