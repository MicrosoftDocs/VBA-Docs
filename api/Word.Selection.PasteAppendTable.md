---
title: Selection.PasteAppendTable method (Word)
keywords: vbawd10.chm158663666
f1_keywords:
- vbawd10.chm158663666
ms.prod: word
api_name:
- Word.Selection.PasteAppendTable
ms.assetid: 60e12397-563f-f8bc-160f-f24a12794d01
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.PasteAppendTable method (Word)

Merges pasted cells into an existing table by inserting the pasted rows between the selected rows. No cells are overwritten.


## Syntax

_expression_. `PasteAppendTable`

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Example

This example pastes table cells by inserting rows into the current table at the insertion point. This example assumes that the Clipboard contains a collection of table cells.


```vb
Sub PasteAppend 
 Selection.PasteAppendTable 
End Sub
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]