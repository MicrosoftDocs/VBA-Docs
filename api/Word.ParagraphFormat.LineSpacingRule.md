---
title: ParagraphFormat.LineSpacingRule property (Word)
keywords: vbawd10.chm156434542
f1_keywords:
- vbawd10.chm156434542
ms.prod: word
api_name:
- Word.ParagraphFormat.LineSpacingRule
ms.assetid: a08e9eeb-1b85-7cd8-a497-ac7d63234267
ms.date: 06/08/2017
localization_priority: Normal
---


# ParagraphFormat.LineSpacingRule property (Word)

Returns or sets the line spacing for the specified paragraph formatting. Read/write  **[WdLineSpacing](Word.WdLineSpacing.md)**.


## Syntax

_expression_. `LineSpacingRule`

_expression_ Required. A variable that represents a '[ParagraphFormat](Word.ParagraphFormat.md)' object.


## Remarks

Use  **wdLineSpaceSingle**, **wdLineSpace1pt5**, or **wdLineSpaceDouble** to set the line spacing to one of these values. To set the line spacing to an exact number of points or to a multiple number of lines, you must also set the **[LineSpacing](Word.ParagraphFormat.LineSpacing.md)** property.


## Example

This example double-spaces the lines in the first paragraph of the active document.


```vb
ActiveDocument.Paragraphs(1).LineSpacingRule = _ 
 wdLineSpaceDouble
```

This example returns the line spacing rule used for the first paragraph in the selection.




```vb
lrule = Selection.Paragraphs(1).LineSpacingRule
```


## See also


[ParagraphFormat Object](Word.ParagraphFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]