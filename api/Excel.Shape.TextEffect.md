---
title: Shape.TextEffect property (Excel)
keywords: vbaxl10.chm636108
f1_keywords:
- vbaxl10.chm636108
ms.prod: excel
api_name:
- Excel.Shape.TextEffect
ms.assetid: 4e2920c3-340c-c113-2667-4d4779cfb59f
ms.date: 05/14/2019
localization_priority: Normal
---


# Shape.TextEffect property (Excel)

Returns a **[TextEffectFormat](Excel.TextEffectFormat.md)** object that contains text-effect formatting properties for the specified shape. Read-only.


## Syntax

_expression_.**TextEffect**

_expression_ A variable that represents a **[Shape](Excel.Shape.md)** object.


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