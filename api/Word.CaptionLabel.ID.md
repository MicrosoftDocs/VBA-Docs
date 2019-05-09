---
title: CaptionLabel.ID property (Word)
keywords: vbawd10.chm158924802
f1_keywords:
- vbawd10.chm158924802
ms.prod: word
api_name:
- Word.CaptionLabel.ID
ms.assetid: ddbbbc0b-8f83-041b-8a80-c0600e1c5231
ms.date: 06/08/2017
localization_priority: Normal
---


# CaptionLabel.ID property (Word)

Returns a  **WdCaptionLabelID** constant that represents the type for the specified caption label if the **BuiltIn** property of the **CaptionLabel** object is **True**. Read-only.


## Syntax

_expression_.**ID**

_expression_ Required. A variable that represents a '[CaptionLabel](Word.CaptionLabel.md)' object.


## Example

This example displays the built-in caption label names and ID values.


```vb
For Each cl In CaptionLabels 
 If cl.BuiltIn = True Then MsgBox cl.Name & " " & cl.ID 
Next cl
```


## See also


[CaptionLabel Object](Word.CaptionLabel.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]