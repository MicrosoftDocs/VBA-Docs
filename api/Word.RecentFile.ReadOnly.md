---
title: RecentFile.ReadOnly property (Word)
keywords: vbawd10.chm157548546
f1_keywords:
- vbawd10.chm157548546
ms.prod: word
api_name:
- Word.RecentFile.ReadOnly
ms.assetid: 69c413bb-9758-06f8-05f1-318ec320fa82
ms.date: 06/08/2017
localization_priority: Normal
---


# RecentFile.ReadOnly property (Word)

 **True** if changes to the document cannot be saved to the original document. Read/write **Boolean**.


## Syntax

 _expression_. `ReadOnly`

 _expression_ Required. A variable that represents a '[RecentFile](Word.RecentFile.md)' object.


## Example

This example opens the most recently used file as a read-only document.


```vb
With RecentFiles(1) 
 .ReadOnly = True 
 .Open 
End With
```


## See also


[RecentFile Object](Word.RecentFile.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]