---
title: Selection.InsertParagraph method (Word)
keywords: vbawd10.chm158662816
f1_keywords:
- vbawd10.chm158662816
ms.prod: word
api_name:
- Word.Selection.InsertParagraph
ms.assetid: bceda293-7294-8769-75fe-4792199439c1
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.InsertParagraph method (Word)

Replaces the specified selection with a new paragraph.


## Syntax

_expression_. `InsertParagraph`

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Remarks

After using method, the selection contains the new paragraph. If you don't want to replace the current selection, use the **Collapse** method before using this method. You can also use the **[InsertParagraphBefore](Word.Selection.InsertParagraphBefore.md)** or **[InsertParagraphAfter](Word.Selection.InsertParagraphAfter.md)** method to insert a new paragraph before or after a selection.


## Example

This example collapses the selection and then inserts a paragraph mark at the insertion point.


```vb
With Selection 
 .Collapse Direction:=wdCollapseStart 
 .InsertParagraph 
 .Collapse Direction:=wdCollapseEnd 
End With
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]