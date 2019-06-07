---
title: Font.Scaling property (Publisher)
keywords: vbapb10.chm5373977
f1_keywords:
- vbapb10.chm5373977
ms.prod: publisher
api_name:
- Publisher.Font.Scaling
ms.assetid: 4ff0c484-12f8-38e3-72fd-dfd34507aec1
ms.date: 06/08/2019
localization_priority: Normal
---


# Font.Scaling property (Publisher)

Returns or sets a **Variant** value used to scale the width of the characters in the text range as a percentage of the current font size. Read/write.


## Syntax

_expression_.**Scaling**

_expression_ A variable that represents a **[Font](Publisher.Font.md)** object.


## Return value

Variant


## Remarks

Valid range is 0.1 to 600.0, where the number represents the percentage of the current font size. Indeterminate values are returned as -2.


## Example

This example scales the width of the text in the second story by 200%. For this example to work, a second story with text must exist in the active document.

```vb
Sub ScaleUp() 
 
 Application.ActiveDocument.Stories(2).TextRange.Font.Scaling = 200 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]