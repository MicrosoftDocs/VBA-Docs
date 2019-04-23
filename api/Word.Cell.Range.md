---
title: Cell.Range property (Word)
keywords: vbawd10.chm156106752
f1_keywords:
- vbawd10.chm156106752
ms.prod: word
api_name:
- Word.Cell.Range
ms.assetid: 579a25ad-91fa-a7c9-7eb8-4307521aeddd
ms.date: 03/26/2019
localization_priority: Normal
---


# Cell.Range property (Word)

Returns a **[Range](Word.Range.md)** object that represents the portion of a document that's contained in the specified object.


## Syntax

_expression_.**Range**

_expression_ A variable that represents a **[Cell](Word.Cell.md)** object.


## Example

This example copies the contents of the first cell in the first row in the first table.

```vb
If ActiveDocument.Tables.Count >= 1 Then _ 
 ActiveDocument.Tables(1).Rows(1).Cells(1).Range.Copy
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]