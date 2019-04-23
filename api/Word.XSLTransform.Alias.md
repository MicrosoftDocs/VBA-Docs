---
title: XSLTransform.Alias property (Word)
keywords: vbawd10.chm76742658
f1_keywords:
- vbawd10.chm76742658
ms.prod: word
api_name:
- Word.XSLTransform.Alias
ms.assetid: 38615e8f-cb40-6e83-f29c-520430f16ada
ms.date: 06/08/2017
localization_priority: Normal
---


# XSLTransform.Alias property (Word)

Returns a  **String** that represents the display name for the specified object.


## Syntax

_expression_. `Alias`

_expression_ Required. A variable that represents a '[XSLTransform](Word.XSLTransform.md)' object.


## Example

The following example shows the display name for the first schema attached to the active document.


```vb
MsgBox Application.XMLNamespaces(1).Alias
```


## See also


[XSLTransform Object](Word.XSLTransform.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]