---
title: Selection.Intersect method (Visio)
keywords: vis_sdr.chm11116375
f1_keywords:
- vis_sdr.chm11116375
ms.prod: visio
api_name:
- Visio.Selection.Intersect
ms.assetid: 5dc63a77-62de-3892-6ed2-bcb5cb0a29f1
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Intersect method (Visio)

Creates one closed shape from the area in which selected shapes overlap or intersect.


## Syntax

_expression_. `Intersect`

_expression_ A variable that represents a **[Selection](Visio.Selection.md)** object.


## Return value

Nothing


## Remarks

Calling the  **Intersect** method is equivalent to clicking **Intersect** in the Microsoft Visio user interface (click **Operations** in the **Shape Design** group on the [Developer](../visio/How-to/run-visio-in-developer-mode.md) tab). The produced shape will be the topmost shape in its containing shape and will inherit the text and formatting of the first selected shape.

The original shapes are deleted and no shapes are selected when the operation is complete.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]