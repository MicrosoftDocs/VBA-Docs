---
title: Selection.Cut method (PowerPoint)
keywords: vbapp10.chm508003
f1_keywords:
- vbapp10.chm508003
ms.prod: powerpoint
api_name:
- PowerPoint.Selection.Cut
ms.assetid: 305103ad-f4d1-8173-e331-17750587d865
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Cut method (PowerPoint)

Deletes the specified object and places it on the Clipboard.


## Syntax

_expression_.**Cut**

_expression_ A variable that represents a [Selection](PowerPoint.Selection.md) object.


## Example

This example deletes the selection in window one and places a copy of it on the Clipboard.


```vb
Windows(1).Selection.Cut
```


## See also


[Selection Object](PowerPoint.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]