---
title: Table.PreferredWidthType property (Word)
keywords: vbawd10.chm156303472
f1_keywords:
- vbawd10.chm156303472
ms.prod: word
api_name:
- Word.Table.PreferredWidthType
ms.assetid: 92954057-5ecd-3d43-c547-e1e1a6c83904
ms.date: 06/08/2017
localization_priority: Normal
---


# Table.PreferredWidthType property (Word)

Returns or sets the preferred unit of measurement to use for the width of the specified table. Read/write  **[WdPreferredWidthType](Word.WdPreferredWidthType.md)**.


## Syntax

_expression_. `PreferredWidthType`

_expression_ Required. A variable that represents a '[Table](Word.Table.md)' object.


## Example

This example sets Microsoft Word to accept widths as a percentage of window width, and then it sets the width of the first table in the document to 50% of the window width.


```vb
With ActiveDocument.Tables(1) 
 .PreferredWidthType = wdPreferredWidthPercent 
 .PreferredWidth = 50 
End With
```


## See also


[Table Object](Word.Table.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]