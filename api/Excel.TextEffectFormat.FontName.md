---
title: TextEffectFormat.FontName property (Excel)
keywords: vbaxl10.chm118005
f1_keywords:
- vbaxl10.chm118005
ms.prod: excel
api_name:
- Excel.TextEffectFormat.FontName
ms.assetid: d5aee022-b60b-f747-3c6b-7ae7e70cf6f8
ms.date: 05/17/2019
localization_priority: Normal
---


# TextEffectFormat.FontName property (Excel)

Returns or sets the name of the font in the specified WordArt. Read/write **String**.


## Syntax

_expression_.**FontName**

_expression_ A variable that represents a **[TextEffectFormat](Excel.TextEffectFormat.md)** object.


## Example

This example sets the font name to Courier New for shape three on _myDocument_ if the shape is WordArt.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3) 
 If .Type = msoTextEffect Then 
 .TextEffect.FontName = "Courier New" 
 End If 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]