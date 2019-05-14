---
title: ShapeRange.TextEffect property (Excel)
keywords: vbaxl10.chm640115
f1_keywords:
- vbaxl10.chm640115
ms.prod: excel
api_name:
- Excel.ShapeRange.TextEffect
ms.assetid: 95c2ab5d-061e-f50e-fc2b-7c44ffca7ce9
ms.date: 05/14/2019
localization_priority: Normal
---


# ShapeRange.TextEffect property (Excel)

Returns a **[TextEffectFormat](Excel.TextEffectFormat.md)** object that contains text-effect formatting properties for the specified shape. Read-only.


## Syntax

_expression_.**TextEffect**

_expression_ A variable that represents a **[ShapeRange](Excel.shaperange.md)** object.


## Example

This example sets the font style to bold for shape three on _myDocument_ if the shape is WordArt.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3) 
 If .Type = msoTextEffect Then 
 .TextEffect.FontBold = True 
 End If 
End With
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]