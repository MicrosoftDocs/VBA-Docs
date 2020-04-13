---
title: Column.PreferredWidth property (Word)
keywords: vbawd10.chm156172394
f1_keywords:
- vbawd10.chm156172394
ms.prod: word
api_name:
- Word.Column.PreferredWidth
ms.assetid: b275a938-c0a0-3f92-f67e-6b3bead43466
ms.date: 06/08/2017
localization_priority: Normal
---


# Column.PreferredWidth property (Word)

Returns or sets the preferred width (in points or as a percentage of the window width) for the specified column. Read/write  **Single**.


## Syntax

 _expression_.**PreferredWidth**

_expression_ Required. An expression that returns a '[Column](Word.Column.md)' object.


## Remarks

If the **[PreferredWidthType](Word.Column.PreferredWidthType.md)** property is set to **wdPreferredWidthPoints**, the **PreferredWidth** property returns or sets the width in points. If the **PreferredWidthType** property is set to **wdPreferredWidthPercent**, the **PreferredWidth** property returns or sets the width as a percentage of the window width.


## See also


[Column Object](Word.Column.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]