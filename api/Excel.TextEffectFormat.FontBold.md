---
title: TextEffectFormat.FontBold property (Excel)
keywords: vbaxl10.chm118003
f1_keywords:
- vbaxl10.chm118003
ms.prod: excel
api_name:
- Excel.TextEffectFormat.FontBold
ms.assetid: 19773cce-32d3-b07f-4650-5a19a4aa469a
ms.date: 05/17/2019
localization_priority: Normal
---


# TextEffectFormat.FontBold property (Excel)

Returns **msoTrue** if the font in the specified WordArt is bold. Read/write **[MsoTriState](Office.MsoTriState.md)**.


## Syntax

_expression_.**FontBold**

_expression_ A variable that represents a **[TextEffectFormat](Excel.TextEffectFormat.md)** object.


## Example

This example sets the font to bold for shape three on _myDocument_ if the shape is WordArt.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3) 
    If .Type = msoTextEffect Then 
        .TextEffect.FontBold = msoTrue 
    End If 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]