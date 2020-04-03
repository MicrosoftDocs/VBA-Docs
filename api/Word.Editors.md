---
title: Editors object (Word)
keywords: vbawd10.chm140
f1_keywords:
- vbawd10.chm140
ms.prod: word
api_name:
- Word.Editors
ms.assetid: acce718a-e3c1-deac-8b7f-fd8a5a9e47c6
ms.date: 06/08/2017
localization_priority: Normal
---


# Editors object (Word)

A collection of  **Editor** objects that represents a collection of users or groups of users who have been given specific permissions to edit portions of a document.


## Remarks

Use the **Add** method to give a specified user or group permission to modify a range or selection within a document. The following example gives the current user editing permission to modify the active selection.


```vb
Dim objEditor As Editor 
 
Set objEditor = Selection.Editors.Add(wdEditorCurrent)
```


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]