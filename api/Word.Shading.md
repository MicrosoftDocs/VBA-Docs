---
title: Shading object (Word)
keywords: vbawd10.chm2362
f1_keywords:
- vbawd10.chm2362
ms.prod: word
api_name:
- Word.Shading
ms.assetid: e136509a-1be1-29e4-7b37-1faf659e37ba
ms.date: 06/08/2017
localization_priority: Normal
---


# Shading object (Word)

Contains shading attributes for an object.


## Remarks

Use the **Shading** property to return the **Shading** object. The following example applies fine gray shading to the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).Shading.Texture = wdTexture10Percent
```

The following example applies shading with different foreground and background colors to the selection.




```vb
With Selection.Shading 
 .Texture = wdTexture20Percent 
 .ForegroundPatternColorIndex = wdBlue 
 .BackgroundPatternColorIndex = wdYellow 
End With
```

The following example applies a vertical line texture to the first row in the first table in the active document.




```vb
ActiveDocument.Tables(1).Rows(1).Shading.Texture = _ 
 wdTextureVertical
```


## Properties



|Name|
|:-----|
|[Application](Word.Shading.Application.md)|
|[BackgroundPatternColor](Word.Shading.BackgroundPatternColor.md)|
|[BackgroundPatternColorIndex](Word.Shading.BackgroundPatternColorIndex.md)|
|[Creator](Word.Shading.Creator.md)|
|[ForegroundPatternColor](Word.Shading.ForegroundPatternColor.md)|
|[ForegroundPatternColorIndex](Word.Shading.ForegroundPatternColorIndex.md)|
|[Parent](Word.Shading.Parent.md)|
|[Texture](Word.Shading.Texture.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]