---
title: TextEffectFormat.Alignment property (Excel)
keywords: vbaxl10.chm118002
f1_keywords:
- vbaxl10.chm118002
ms.prod: excel
api_name:
- Excel.TextEffectFormat.Alignment
ms.assetid: 0a86ac22-9496-d801-0cfb-a9fca5c30fec
ms.date: 05/17/2019
localization_priority: Normal
---


# TextEffectFormat.Alignment property (Excel)

Returns or sets an **[MsoTextEffectAlignment](Office.MsoTextEffectAlignment.md)** value that represents the alignment for WordArt.


## Syntax

_expression_.**Alignment**

_expression_ A variable that represents a **[TextEffectFormat](Excel.TextEffectFormat.md)** object.


## Example

This example adds a WordArt object to worksheet one and then right aligns the WordArt.

```vb
Set mySh = Worksheets(1).Shapes 
Set myTE = mySh.AddTextEffect(PresetTextEffect:=msoTextEffect1, _ 
    Text:="Test Text", FontName:="Palatino", FontSize:=54, _ 
    FontBold:=True, FontItalic:=False, Left:=100, Top:=50) 
myTE.TextEffect.Alignment = msoTextEffectAlignmentRight
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]