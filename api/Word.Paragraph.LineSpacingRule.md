---
title: Paragraph.LineSpacingRule property (Word)
keywords: vbawd10.chm156696686
f1_keywords:
- vbawd10.chm156696686
ms.prod: word
api_name:
- Word.Paragraph.LineSpacingRule
ms.assetid: 02bf5c99-fe6d-3bc4-9388-e8b372d00549
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph.LineSpacingRule property (Word)

Returns or sets the line spacing for the specified paragraph. Read/write  **[WdLineSpacing](Word.WdLineSpacing.md)**.


## Syntax

_expression_. `LineSpacingRule`

_expression_ Required. A variable that represents a '[Paragraph](Word.Paragraph.md)' object.


## Remarks

Use  **wdLineSpaceSingle**, **wdLineSpace1pt5**, or **wdLineSpaceDouble** to set the line spacing to one of these values. To set the line spacing to an exact number of points or to a multiple number of lines, you must also set the **[LineSpacing](Word.Paragraph.LineSpacing.md)** property.


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


[Paragraph Object](Word.Paragraph.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]