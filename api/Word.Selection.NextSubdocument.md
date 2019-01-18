---
title: Selection.NextSubdocument method (Word)
keywords: vbawd10.chm158663170
f1_keywords:
- vbawd10.chm158663170
ms.prod: word
api_name:
- Word.Selection.NextSubdocument
ms.assetid: e8527994-23f4-c9a1-d96c-c2018e07efad
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.NextSubdocument method (Word)

Moves the selection to the next subdocument.


## Syntax

 _expression_. `NextSubdocument`

 _expression_ Required. A variable that represents a '[Selection](Word.Selection.md)' object.


## Remarks

If there isn't another subdocument, an error occurs.


## Example

This example switches the active document to master document view and selects the first subdocument.


```vb
If ActiveDocument.Subdocuments.Count >= 1 Then 
 ActiveDocument.ActiveWindow.View.Type = wdMasterView 
 Selection.HomeKey unit:=wdStory, Extend:=wdMove 
 Selection.NextSubdocument 
End If
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]