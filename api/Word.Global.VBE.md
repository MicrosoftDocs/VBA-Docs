---
title: Global.VBE property (Word)
keywords: vbawd10.chm163119165
f1_keywords:
- vbawd10.chm163119165
ms.prod: word
api_name:
- Word.Global.VBE
ms.assetid: 20a5da58-0e00-9cb2-59ae-cb94178f79c8
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.VBE property (Word)

Returns a  **VBE** object that represents the Visual Basic Editor.


## Syntax

_expression_.**VBE**

_expression_ Required. A variable that represents a '[Global](Word.Global.md)' object.


## Example

This example displays the number of references available for the active project.


```vb
MsgBox "References = " & VBE.ActiveVBProject.References.Count
```


## See also


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]