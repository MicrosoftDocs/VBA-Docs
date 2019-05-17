---
title: ThreeDFormat.Perspective property (Excel)
keywords: vbaxl10.chm119008
f1_keywords:
- vbaxl10.chm119008
ms.prod: excel
api_name:
- Excel.ThreeDFormat.Perspective
ms.assetid: 9f31508e-c723-e55a-07a9-cef1bc526136
ms.date: 05/17/2019
localization_priority: Normal
---


# ThreeDFormat.Perspective property (Excel)

Returns or sets an **[MsoTriState](Office.MsoTriState.md)** value that determines whether the extrusion appears in perspective.


## Syntax

_expression_.**Perspective**

_expression_ A variable that represents a **[ThreeDFormat](Excel.ThreeDFormat.md)** object.


## Remarks

Returns **msoFalse** if the extrusion is a parallel, or orthographic, projection—that is, the walls don't narrow toward a vanishing point.

Returns **msoTrue** if the extrusion appears in perspective—that is, the walls of the extrusion narrow toward a vanishing point.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]