---
title: TextEffectFormat.NormalizedHeight property (Excel)
keywords: vbaxl10.chm118008
f1_keywords:
- vbaxl10.chm118008
ms.prod: excel
api_name:
- Excel.TextEffectFormat.NormalizedHeight
ms.assetid: 25c9c1ed-971d-3a9f-bb3c-5059f2dd80db
ms.date: 05/17/2019
localization_priority: Normal
---


# TextEffectFormat.NormalizedHeight property (Excel)

Returns **msoTrue** if all characters (both uppercase and lowercase) in the specified WordArt are the same height. Read/write **[MsoTriState](Office.MsoTriState.md)**.


## Syntax

_expression_.**NormalizedHeight**

_expression_ A variable that represents a **[TextEffectFormat](Excel.TextEffectFormat.md)** object.


## Example

This example adds WordArt that contains the text Test Effect to _myDocument_ and gives the new WordArt the name texteff1. The code then makes all characters in the shape named texteff1 the same height.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes.AddTextEffect( _ 
 PresetTextEffect:=msoTextEffect1, _ 
 Text:="Test Effect", FontName:="Courier New", _ 
 FontSize:=44, FontBold:=True, _ 
 FontItalic:=False, Left:=10, Top:=10).Name = "texteff1" 
myDocument.Shapes("texteff1").TextEffect.NormalizedHeight = msoTrue
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]