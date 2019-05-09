---
title: ProtectedViewWindows.Count property (Word)
keywords: vbawd10.chm82313217
f1_keywords:
- vbawd10.chm82313217
ms.prod: word
api_name:
- Word.ProtectedViewWindows.Count
ms.assetid: edd30c3f-6890-be71-57c0-0aa3b1dea1a5
ms.date: 06/08/2017
localization_priority: Normal
---


# ProtectedViewWindows.Count property (Word)

Returns a  **Long** that represents the number of Protected View windows in the collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ An expression that returns a **[ProtectedViewWindows](Word.ProtectedViewWindows.md)** object.


## Example

The following code example displays the number of Protected View windows that are currently open.


```vb
MsgBox ProtectedViewWindows.Count
```


## See also


[ProtectedViewWindows Object](Word.ProtectedViewWindows.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]