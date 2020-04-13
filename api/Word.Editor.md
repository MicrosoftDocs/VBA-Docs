---
title: Editor object (Word)
keywords: vbawd10.chm3442
f1_keywords:
- vbawd10.chm3442
ms.prod: word
api_name:
- Word.Editor
ms.assetid: af0c80f5-8c8a-be0e-4475-d3b3b3bacd0d
ms.date: 06/08/2017
localization_priority: Normal
---


# Editor object (Word)

Represents a single user who has been given specific permissions to edit portions of a document. 


## Remarks

Users who can be given permissions include individual contributors and groups of users as defined for Document Workspace sites.

The permissions you assign to ranges and selections go into effect only after a document is protected. Use the **Editors** collection and the **Editor** object to assign specific permissions to sections of a document. Then use the **Protect** method to protect the document.

Use the **Add** method of the **Editors** collection to give a specified user or group permission to modify a range or selection within a document. The following example gives the current user editing permission to modify the active selection.




```vb
Dim objEditor As Editor 
 
Set objEditor = Selection.Editors.Add(wdEditorCurrent)
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]