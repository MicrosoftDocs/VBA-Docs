---
title: Cell.InheritedFormulaSource property (Visio)
keywords: vis_sdr.chm10150680
f1_keywords:
- vis_sdr.chm10150680
ms.prod: visio
api_name:
- Visio.Cell.InheritedFormulaSource
ms.assetid: 62aedef3-06b1-2fc3-5fd2-03f77668548f
ms.date: 06/08/2017
localization_priority: Normal
---


# Cell.InheritedFormulaSource property (Visio)

Returns the cell from which this cell inherited its formula. Read-only.


## Syntax

_expression_.**InheritedFormulaSource**

_expression_ A variable that represents a **[Cell](Visio.Cell.md)** object.


## Return value

Cell


## Remarks

If the formula in this cell is a local formula, the  **InheritedFormulaSource** property returns itself.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]