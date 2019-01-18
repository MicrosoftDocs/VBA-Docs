---
title: Tasks.ExitWindows method (Word)
keywords: vbawd10.chm159580163
f1_keywords:
- vbawd10.chm159580163
ms.prod: word
api_name:
- Word.Tasks.ExitWindows
ms.assetid: c2af5fdf-948d-c9cb-1a6a-8cde29ab630c
ms.date: 06/08/2017
localization_priority: Normal
---


# Tasks.ExitWindows method (Word)

Closes all open applications, quits Microsoft Windows, and logs the current user off.


## Syntax

 _expression_. `ExitWindows`

 _expression_ Required. A variable that represents a '[Tasks](Word.tasks.md)' collection.


## Remarks

This method does not save changes to open Microsoft Word documents; however, it does prompt you to save changes to open documents in other Windows-based applications.


## Example

This example saves all open Word documents, closes Word, and then quits Microsoft Windows.


```vb
Documents.Save NoPrompt:=True, _ 
 OriginalFormat:=wdOriginalDocumentFormat 
Tasks.ExitWindows
```


## See also


[Tasks Collection Object](Word.tasks.md)

