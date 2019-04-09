---
title: Cells.PreferredWidth property (Word)
keywords: vbawd10.chm155844711
f1_keywords:
- vbawd10.chm155844711
ms.prod: word
api_name:
- Word.Cells.PreferredWidth
ms.assetid: 3f52069b-0fb2-0379-7f64-39d2ef9c02e1
ms.date: 06/08/2017
localization_priority: Normal
---


# Cells.PreferredWidth property (Word)

Returns or sets the preferred width (in points or as a percentage of the window width) for the specified cells. Read/write  **Single**.


## Syntax

_expression_. `PreferredWidth`

_expression_ Required. A variable that represents a '[Cells](Word.cells.md)' collection.


## Remarks

If the  **PreferredWidthType** property is set to **wdPreferredWidthPoints**, the **PreferredWidth** property returns or sets the width in points. If the **PreferredWidthType** property is set to **wdPreferredWidthPercent**, the **PreferredWidth** property returns or sets the width as a percentage of the window width.


## See also


[Cells Collection Object](Word.cells.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]