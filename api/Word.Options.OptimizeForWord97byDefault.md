---
title: Options.OptimizeForWord97byDefault property (Word)
keywords: vbawd10.chm162988455
f1_keywords:
- vbawd10.chm162988455
ms.prod: word
api_name:
- Word.Options.OptimizeForWord97byDefault
ms.assetid: 6d129c8d-24ed-d21c-70a6-f5cd79273b4f
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.OptimizeForWord97byDefault property (Word)

 **True** if Microsoft Word optimizes all new documents for viewing in Word 97 by disabling any incompatible formatting. Read/write **Boolean**.


## Syntax

_expression_. `OptimizeForWord97byDefault`

_expression_ A variable that represents a '[Options](Word.Options.md)' object.


## Remarks

To optimize a single document for Word 97, use the **[OptimizeForWord97](Word.Document.OptimizeForWord97.md)** property.


## Example

This example sets Word to disable all formatting in new documents that's incompatible with Word 97, and then it creates a new document whose  **OptimizeForWord97** property is automatically set to True.


```vb
Options.OptimizeForWord97byDefault = True 
MsgBox Documents.Add(DocumentType:=wdNewBlankDocument) _ 
 .OptimizeForWord97
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]