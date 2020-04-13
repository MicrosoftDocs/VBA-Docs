---
title: ShadowFormat.IncrementOffsetX method (Word)
keywords: vbawd10.chm164364298
f1_keywords:
- vbawd10.chm164364298
ms.prod: word
api_name:
- Word.ShadowFormat.IncrementOffsetX
ms.assetid: 0d564836-550d-30fa-e519-c6dc571d538d
ms.date: 06/08/2017
localization_priority: Normal
---


# ShadowFormat.IncrementOffsetX method (Word)

Changes the horizontal offset of the shadow by the specified number of points.


## Syntax

_expression_.**IncrementOffsetX** (_Increment_)

_expression_ Required. A variable that represents a **[ShadowFormat](Word.ShadowFormat.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|Specifies how far the shadow offset is to be moved horizontally, in points. A positive value moves the shadow to the right; a negative value moves it to the left.|

## Remarks

Use the **[OffsetX](Word.ShadowFormat.OffsetX.md)** property to set the absolute horizontal shadow offset.


## Example

This example moves the shadow on the third shape in the active document to the left by 3 points.


```vb
ActiveDocument.Shapes(3).Shadow.IncrementOffsetX -3
```


## See also


[ShadowFormat Object](Word.ShadowFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]