---
title: Range.PreviousSubdocument method (Word)
keywords: vbawd10.chm157155548
f1_keywords:
- vbawd10.chm157155548
ms.prod: word
api_name:
- Word.Range.PreviousSubdocument
ms.assetid: 542149f4-1a0c-bf1b-1cf6-9e8097af321e
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.PreviousSubdocument method (Word)

Moves the range to the previous subdocument.


## Syntax

_expression_. `PreviousSubdocument`

_expression_ Required. A variable that represents a **[Range](Word.Range.md)** object.


## Remarks

If there isn't another subdocument, an error occurs.


## Example

This example switches the active document to master document view and selects the previous subdocument.


```vb
If ActiveDocument.Subdocuments.Count >= 1 Then 
 ActiveDocument.ActiveWindow.View.Type = wdMasterView 
 Selection.EndKey Unit:=wdStory, Extend:=wdMove 
 Selection.PreviousSubdocument 
End If
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]