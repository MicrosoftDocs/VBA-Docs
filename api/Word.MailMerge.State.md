---
title: MailMerge.State property (Word)
keywords: vbawd10.chm153092098
f1_keywords:
- vbawd10.chm153092098
ms.prod: word
api_name:
- Word.MailMerge.State
ms.assetid: eeee1112-91fb-ec32-a9ea-ab999f0c28e9
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMerge.State property (Word)

Returns the current state of a mail merge operation. Read-only  **WdMailMergeState**.


## Syntax

 _expression_. `State`

 _expression_ Required. A variable that represents a '[MailMerge](Word.MailMerge.md)' object.


## Example

This example executes a mail merge if the active document is a main document with an attached data source.


```vb
Set myMerge = ActiveDocument.MailMerge 
If myMerge.State = wdMainAndDataSource Then myMerge.Execute
```


## See also


[MailMerge Object](Word.MailMerge.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]