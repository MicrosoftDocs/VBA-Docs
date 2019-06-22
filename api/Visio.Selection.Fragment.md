---
title: Selection.Fragment method (Visio)
keywords: vis_sdr.chm11116305
f1_keywords:
- vis_sdr.chm11116305
ms.prod: visio
api_name:
- Visio.Selection.Fragment
ms.assetid: e648675f-e60a-6a21-182e-32aa913df335
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Fragment method (Visio)

Breaks selected shapes into smaller shapes.


## Syntax

_expression_. `Fragment`

_expression_ A variable that represents a **[Selection](Visio.Selection.md)** object.


## Return value

Nothing


## Remarks

Calling the  **Fragment** method is equivalent to clicking **Fragment** in the Microsoft Visio user interface (click **Operations** in the **Shape Design** group on the [Developer](../visio/How-to/run-visio-in-developer-mode.md) tab). The produced shapes are the topmost shapes in the containing shape of the selected shapes. They inherit the formatting of the first selected shape and have no text.

The original shapes are deleted and there aren't any shapes selected when the operation is complete.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]