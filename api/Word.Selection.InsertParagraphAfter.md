---
title: Selection.InsertParagraphAfter method (Word)
keywords: vbawd10.chm158662817
f1_keywords:
- vbawd10.chm158662817
ms.prod: word
api_name:
- Word.Selection.InsertParagraphAfter
ms.assetid: ae97fbab-417a-14e2-0154-f0361826f903
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.InsertParagraphAfter method (Word)

Inserts a paragraph mark after a selection.


## Syntax

_expression_. `InsertParagraphAfter`

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Remarks

After using this method, the selection expands to include the new paragraph.


## Example

This example inserts a new paragraph after the current paragraph.


```vb
With Selection 
 .Move Unit:=wdParagraph 
 .InsertParagraphAfter 
 .Collapse Direction:=wdCollapseStart 
End With
```

This example inserts a paragraph at the end of the active document. The  **Content** property returns a **Range** object.




```vb
ActiveDocument.Content.InsertParagraphAfter
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]