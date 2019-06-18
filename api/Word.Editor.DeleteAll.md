---
title: Editor.DeleteAll method (Word)
keywords: vbawd10.chm225575413
f1_keywords:
- vbawd10.chm225575413
ms.prod: word
api_name:
- Word.Editor.DeleteAll
ms.assetid: 81e69276-99f8-6525-2b45-c9e63feb1c53
ms.date: 06/08/2017
localization_priority: Normal
---


# Editor.DeleteAll method (Word)

Deletes all editing permissions in a document for a specific user.


## Syntax

_expression_.**DeleteAll**

_expression_ Required. A variable that represents an '[Editor](Word.Editor.md)' object.


## Example

The following example deletes all editing permissions for the first user in the  **Editors** collection in the active document.


```vb
Dim objEditor As Editor 
 
Set objEditor = Selection.Editors(1) 
 
objEditor.DeleteAll
```


## See also


[Editor Object](Word.Editor.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]