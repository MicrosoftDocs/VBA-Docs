---
title: ProtectedViewWindow.SourceName property (Word)
keywords: vbawd10.chm231735306
f1_keywords:
- vbawd10.chm231735306
ms.prod: word
api_name:
- Word.ProtectedViewWindow.SourceName
ms.assetid: 744639ae-dd9f-cf85-f15f-f2c753fc9d9d
ms.date: 06/08/2017
localization_priority: Normal
---


# ProtectedViewWindow.SourceName property (Word)

Returns the name of the source file for the specified Protected View window. Read-only  **String**.


## Syntax

_expression_.**SourceName**

 _expression_ An expression that returns a '[ProtectedViewWindow](Word.ProtectedViewWindow.md)' object.


## Remarks

This property does not return the path for the source file.


## Example

The following code example returns the path and name of the document associated with the specified Protected View window.


```vb
MsgBox ActiveProtectedViewWindow.SourcePath & "\" _ 
 & ActiveProtectedViewWindow.SourceName 

```


## See also


[ProtectedViewWindow Object](Word.ProtectedViewWindow.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]