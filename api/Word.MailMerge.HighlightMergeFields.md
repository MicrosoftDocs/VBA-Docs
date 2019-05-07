---
title: MailMerge.HighlightMergeFields property (Word)
keywords: vbawd10.chm153092107
f1_keywords:
- vbawd10.chm153092107
ms.prod: word
api_name:
- Word.MailMerge.HighlightMergeFields
ms.assetid: 1002b34a-4492-97df-bb16-bd2c4319e055
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMerge.HighlightMergeFields property (Word)

 **True** to highlight the merge fields in a document. Read/write **Boolean**.


## Syntax

_expression_. `HighlightMergeFields`

_expression_ A variable that represents a '[MailMerge](Word.MailMerge.md)' object.


## Example

This example turns off highlighting merge fields in the active document.


```vb
Sub HighlightFields() 
 ActiveDocument.MailMerge.HighlightMergeFields = False 
End Sub
```


## See also


[MailMerge Object](Word.MailMerge.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]