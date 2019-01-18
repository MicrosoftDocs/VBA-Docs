---
title: TextEffectFormat.KernedPairs property (Excel)
keywords: vbaxl10.chm118007
f1_keywords:
- vbaxl10.chm118007
ms.prod: excel
api_name:
- Excel.TextEffectFormat.KernedPairs
ms.assetid: 107889be-57eb-7fcf-17a1-6a1393009701
ms.date: 06/08/2017
localization_priority: Normal
---


# TextEffectFormat.KernedPairs property (Excel)

 **True** if character pairs in the specified WordArt are kerned. Read/write **MsoTriState**.


## Syntax

_expression_. `KernedPairs`

_expression_ A variable that represents a [TextEffectFormat](./Excel.TextEffectFormat.md) object.


## Remarks





| **MsoTriState** can be one of these **MsoTriState** constants.|
| **msoCTrue**|
| **msoFalse**|
| **msoTriStateMixed**|
| **msoTriStateToggle**|
| **msoTrue** Character pairs in the specified WordArt are kerned.|

## Example

This example turns on character pair kerning for shape three on  `myDocument` if the shape is WordArt.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3) 
 If .Type = msoTextEffect Then 
 .TextEffect.KernedPairs = msoTrue 
 End If 
End With
```


## See also


[TextEffectFormat Object](Excel.TextEffectFormat.md)

