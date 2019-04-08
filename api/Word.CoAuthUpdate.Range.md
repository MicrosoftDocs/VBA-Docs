---
title: CoAuthUpdate.Range property (Word)
keywords: vbawd10.chm247791617
f1_keywords:
- vbawd10.chm247791617
ms.prod: word
api_name:
- Word.CoAuthUpdate.Range
ms.assetid: 786bc4aa-bae2-9ef5-59b2-02eeb6445846
ms.date: 06/08/2017
localization_priority: Normal
---


# CoAuthUpdate.Range property (Word)

Returns a [Range](Word.Range.md) object that represents the portion of a document that is contained in the specified object. Read-only.


## Syntax

_expression_.**Range**

 _expression_ An expression that returns a [CoAuthUpdate](./Word.CoAuthUpdate.md) object.


## Example

The following code example gets the document range for the first update in the active document.


```vb
Dim rng As Range 
 
Set rng = ActiveDocument.CoAuthoring.Updates(1).Range 

```


## See also


[CoAuthUpdate Object](Word.CoAuthUpdate.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]