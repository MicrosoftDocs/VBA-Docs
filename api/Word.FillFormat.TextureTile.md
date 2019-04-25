---
title: FillFormat.TextureTile property (Word)
keywords: vbawd10.chm164102264
f1_keywords:
- vbawd10.chm164102264
ms.prod: word
api_name:
- Word.FillFormat.TextureTile
ms.assetid: 670db5f6-8543-2c6e-6aeb-98f240716421
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat.TextureTile property (Word)

Returns or sets whether the texture fill is tiled or centered. Read/write.


## Syntax

_expression_.**TextureTile**

 _expression_ An expression that returns a **[FillFormat](word.fillformat.md)** object.


## Remarks

The value returned by the  **TextureTile** property can be one of the following[MsoTriState](Office.MsoTriState.md) constants.



|Constant|Description|
|:-----|:-----|
| **msoFalse**|The texture fill is centered.|
| **msoTrue**|The texture fill is tiled.|

The setting of the  **TextureTile** property corresponds to the setting of the **Tile picture as texture** box under **Tiling Options** on the **Fill** pane of the **Format Picture** dialog box in the Microsoft Word user interface (under **Drawing Tools**, on the  **Format** tab, expand the **Shape Styles** group.)


## See also


[FillFormat Object](Word.FillFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]