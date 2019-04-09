---
title: Revisions.AcceptAll method (Word)
keywords: vbawd10.chm159383653
f1_keywords:
- vbawd10.chm159383653
ms.prod: word
api_name:
- Word.Revisions.AcceptAll
ms.assetid: bf1fa0d5-22ab-d426-9411-ae3147277448
ms.date: 06/08/2017
localization_priority: Normal
---


# Revisions.AcceptAll method (Word)

Accepts all the tracked changes in a document or range, removes all revision marks, and incorporates the changes into the document.


## Syntax

_expression_. `AcceptAll`

_expression_ Required. A variable that represents a '[Revisions](Word.revisions.md)' collection.


## Remarks

Use the  **AcceptAllRevisions** method to accept all revisions in a document.


## Example

The following code example accepts all the tracked changes in the active document.


```vb
If ActiveDocument.Revisions.Count >= 1 Then _ 
 ActiveDocument.Revisions.AcceptAll
```

The following code example accepts all the tracked changes in the selection.




```vb
Selection.Range.Revisions.AcceptAll
```


## See also


[Revisions Collection Object](Word.revisions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]