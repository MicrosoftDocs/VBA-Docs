---
title: TextEffectFormat.KernedPairs property (Excel)
keywords: vbaxl10.chm118007
f1_keywords:
- vbaxl10.chm118007
ms.prod: excel
api_name:
- Excel.TextEffectFormat.KernedPairs
ms.assetid: 107889be-57eb-7fcf-17a1-6a1393009701
ms.date: 05/17/2019
localization_priority: Normal
---


# TextEffectFormat.KernedPairs property (Excel)

Returns **msoTrue** if character pairs in the specified WordArt are kerned. Read/write **[MsoTriState](Office.MsoTriState.md)**.


## Syntax

_expression_.**KernedPairs**

_expression_ A variable that represents a **[TextEffectFormat](Excel.TextEffectFormat.md)** object.


## Example

This example turns on character pair kerning for shape three on _myDocument_ if the shape is WordArt.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3) 
 If .Type = msoTextEffect Then 
 .TextEffect.KernedPairs = msoTrue 
 End If 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]