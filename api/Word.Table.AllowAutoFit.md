---
title: Table.AllowAutoFit property (Word)
keywords: vbawd10.chm156303470
f1_keywords:
- vbawd10.chm156303470
ms.prod: word
api_name:
- Word.Table.AllowAutoFit
ms.assetid: e8894734-68b3-60bb-7623-9497e4e99e10
ms.date: 06/08/2017
localization_priority: Normal
---


# Table.AllowAutoFit property (Word)

Allows Microsoft Word to automatically resize cells in a table to fit their contents. Read/write  **Boolean**.


## Syntax

_expression_. `AllowAutoFit`

_expression_ A variable that represents a '[Table](Word.Table.md)' object.


## Example

This example sets the first table in the active document to automatically resize based on its contents.


```vb
Sub AllowFit() 
 ActiveDocument.Tables(1).AllowAutoFit = True 
End Sub
```


## See also


[Table Object](Word.Table.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]