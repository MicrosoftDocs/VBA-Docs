---
title: Cell.InheritedValueSource property (Visio)
keywords: vis_sdr.chm10150685
f1_keywords:
- vis_sdr.chm10150685
ms.prod: visio
api_name:
- Visio.Cell.InheritedValueSource
ms.assetid: 1ffa8293-80a9-a43b-c6e1-b90cb2648efa
ms.date: 06/08/2017
localization_priority: Normal
---


# Cell.InheritedValueSource property (Visio)

Returns the cell from which this cell inherited its value. Read-only.


## Syntax

_expression_.**InheritedValueSource**

_expression_ A variable that represents a **[Cell](Visio.Cell.md)** object.


## Return value

Cell


## Remarks

If the value in this cell is a local value, the  **InheritedValueSource** property returns itself.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]