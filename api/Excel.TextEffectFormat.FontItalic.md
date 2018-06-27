---
title: TextEffectFormat.FontItalic Property (Excel)
keywords: vbaxl10.chm118004
f1_keywords:
- vbaxl10.chm118004
ms.prod: excel
api_name:
- Excel.TextEffectFormat.FontItalic
ms.assetid: 5c1f9cd5-e994-3bed-f8ad-ab2ee2d64e7a
ms.date: 06/08/2017
---


# TextEffectFormat.FontItalic Property (Excel)

Returns  **msoTrue** if the font in the specified WordArt is italic. Read/write **[MsoTriState](./Office.MsoTriState.md)** .


## Syntax

 _expression_. `FontItalic`

 _expression_ A variable that represents a [TextEffectFormat](./Excel.TextEffectFormat.md) object.


## Remarks





| **MsoTriState** can be one of these **MsoTriState** constants.|
| **msoCTrue** Does not apply to this property.|
| **msoFalse** The specified WordArt is not italic.|
| **msoTriStateMixed** Does not apply to this property.|
| **msoTriStateToggle** Does not apply to this property.|
| **msoTrue** The specified WordArt is italic.|

## Example

This example sets the font to italic for the shape named "WordArt 4" in  `myDocument`.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes("WordArt 4").TextEffect.FontItalic = msoTrue
```


## See also


[TextEffectFormat Object](Excel.TextEffectFormat.md)

