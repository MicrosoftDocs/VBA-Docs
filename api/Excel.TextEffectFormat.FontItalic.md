---
title: TextEffectFormat.FontItalic property (Excel)
keywords: vbaxl10.chm118004
f1_keywords:
- vbaxl10.chm118004
ms.prod: excel
api_name:
- Excel.TextEffectFormat.FontItalic
ms.assetid: 5c1f9cd5-e994-3bed-f8ad-ab2ee2d64e7a
ms.date: 05/17/2019
localization_priority: Normal
---


# TextEffectFormat.FontItalic property (Excel)

Returns **msoTrue** if the font in the specified WordArt is italic. Read/write **[MsoTriState](Office.MsoTriState.md)**.


## Syntax

_expression_.**FontItalic**

_expression_ A variable that represents a **[TextEffectFormat](Excel.TextEffectFormat.md)** object.


## Example

This example sets the font to italic for the shape named WordArt 4 in _myDocument_.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes("WordArt 4").TextEffect.FontItalic = msoTrue
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]