---
title: RecentFile.Open method (Word)
keywords: vbawd10.chm157548548
f1_keywords:
- vbawd10.chm157548548
ms.prod: word
api_name:
- Word.RecentFile.Open
ms.assetid: bdcc49b7-3511-d625-be46-72dc26a927d0
ms.date: 06/08/2017
localization_priority: Normal
---


# RecentFile.Open method (Word)

Opens the specified object. Returns a  **Document** object representing the opened document.


## Syntax

_expression_.**Open**

_expression_ Required. A variable that represents a **[RecentFile](Word.RecentFile.md)** object.


## Return value

Document


## Example

This example opens each document in the **RecentFiles** collection.


```vb
Sub OpenRecentFiles() 
 Dim rFile As RecentFile 
 For Each rFile In RecentFiles 
 rFile.Open 
 Next rFile 
End Sub
```


## See also


[RecentFile Object](Word.RecentFile.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]