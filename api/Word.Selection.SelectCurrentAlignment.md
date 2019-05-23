---
title: Selection.SelectCurrentAlignment method (Word)
keywords: vbawd10.chm158663174
f1_keywords:
- vbawd10.chm158663174
ms.prod: word
api_name:
- Word.Selection.SelectCurrentAlignment
ms.assetid: 89d76005-c75a-7548-c1da-da292183d5ab
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.SelectCurrentAlignment method (Word)

Extends the selection forward until text with a different paragraph alignment is encountered.


## Syntax

_expression_. `SelectCurrentAlignment`

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Remarks

There are four types of paragraph alignment: left, centered, right, and justified.


## Example

This example positions the insertion point at the beginning of the first paragraph after the current paragraph that doesn't have the same alignment as the current paragraph. If the alignment is the same from the selection to the end of the document, the example moves the selection to the end of the document and displays a message.


```vb
Selection.SelectCurrentAlignment 
Selection.Collapse Direction:=wdCollapseEnd 
If Selection.End = ActiveDocument.Content.End - 1 Then 
 MsgBox "No change in alignment found." 
End If
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]