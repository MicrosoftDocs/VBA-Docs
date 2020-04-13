---
title: Cell.PreferredWidth property (Word)
keywords: vbawd10.chm156106861
f1_keywords:
- vbawd10.chm156106861
ms.prod: word
api_name:
- Word.Cell.PreferredWidth
ms.assetid: 2b59ace4-bd3e-8a30-b81e-0f57d29f8a02
ms.date: 06/08/2017
localization_priority: Normal
---


# Cell.PreferredWidth property (Word)

Returns or sets the preferred width (in points or as a percentage of the window width) for the specified cell. Read/write  **Single**.


## Syntax

_expression_. `PreferredWidth`

_expression_ Required. A variable that represents a '[Cell](Word.Cell.md)' object.


## Remarks

If the **PreferredWidthType** property is set to **wdPreferredWidthPoints**, the **PreferredWidth** property returns or sets the width in points. If the **PreferredWidthType** property is set to **wdPreferredWidthPercent**, the **PreferredWidth** property returns or sets the width as a percentage of the window width.


## See also


[Cell Object](Word.Cell.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]