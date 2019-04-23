---
title: Cell.Height property (Word)
keywords: vbawd10.chm156106759
f1_keywords:
- vbawd10.chm156106759
ms.prod: word
api_name:
- Word.Cell.Height
ms.assetid: 746d61a9-d3e2-c28d-3dac-a892c33be2c7
ms.date: 06/08/2017
localization_priority: Normal
---


# Cell.Height property (Word)

Returns or sets the height of the specified table cell. .


## Syntax

_expression_.**Height**

 _expression_ An expression that returns a [Cell](./Word.Cell.md) object.


## Remarks

If the  **HeightRule** property of the specified row is **wdRowHeightAuto**, **Height** returns **wdUndefined**; setting the **Height** property sets **HeightRule** to **wdRowHeightAtLeast**. Read/write **Single**.


## See also


[Cell Object](Word.Cell.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]