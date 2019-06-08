---
title: Font.Underline property (Publisher)
keywords: vbapb10.chm5373987
f1_keywords:
- vbapb10.chm5373987
ms.prod: publisher
api_name:
- Publisher.Font.Underline
ms.assetid: a01a943e-274d-725e-3f78-aa76c51d5c46
ms.date: 06/08/2019
localization_priority: Normal
---


# Font.Underline property (Publisher)

Returns or sets a **[PbUnderlineType](publisher.pbunderlinetype.md)** constant that indicates the type of underline for the selected characters in the specified font in a text range. Read/write.


## Syntax

_expression_.**Underline**

_expression_ A variable that represents a **[Font](Publisher.Font.md)** object.


## Return value

PbUnderlineType


## Remarks

The **Underline** property value can be one of the **PbUnderlineType** constants declared in the Microsoft Publisher type library.


## Example

This example formats the characters of the first story with a dashed and heavy underline.

```vb
Sub DashHeavy() 
 
 Application.ActiveDocument.Stories(1).TextRange.Font.Underline = pbUnderlineDashHeavy 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]