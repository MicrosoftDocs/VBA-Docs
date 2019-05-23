---
title: Range.NextSubdocument method (Word)
keywords: vbawd10.chm157155547
f1_keywords:
- vbawd10.chm157155547
ms.prod: word
api_name:
- Word.Range.NextSubdocument
ms.assetid: 4c048cc7-a2f6-38b1-e675-4d8870947130
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.NextSubdocument method (Word)

Moves the range to the next subdocument.


## Syntax

_expression_. `NextSubdocument`

_expression_ Required. A variable that represents a **[Range](Word.Range.md)** object.


## Remarks

If there isn't another subdocument, an error occurs.


## Example

This example switches the active document to master document view and selects the first subdocument.


```vb
If ActiveDocument.Subdocuments.Count >= 1 Then 
 ActiveDocument.ActiveWindow.View.Type = wdMasterView 
 Selection.HomeKey unit:=wdStory, Extend:=wdMove 
 Selection.Range.NextSubdocument 
End If
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]