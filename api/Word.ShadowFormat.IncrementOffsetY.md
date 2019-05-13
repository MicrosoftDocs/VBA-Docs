---
title: ShadowFormat.IncrementOffsetY method (Word)
keywords: vbawd10.chm164364299
f1_keywords:
- vbawd10.chm164364299
ms.prod: word
api_name:
- Word.ShadowFormat.IncrementOffsetY
ms.assetid: e0859dd3-9058-32ec-37d8-d14187b69666
ms.date: 06/08/2017
localization_priority: Normal
---


# ShadowFormat.IncrementOffsetY method (Word)

Changes the vertical offset of the shadow by the specified number of points.


## Syntax

_expression_.**IncrementOffsetY** (_Increment_)

_expression_ Required. A variable that represents a **[ShadowFormat](Word.ShadowFormat.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|Specifies how far the shadow offset is to be moved vertically, in points. A positive value moves the shadow down; a negative value moves it up.|

## Remarks

Use the  **[OffsetY](Word.ShadowFormat.OffsetY.md)** property to set the absolute vertical shadow offset.


## Example

This example moves the shadow on the third shape in the active document up by 3 points.


```vb
ActiveDocument.Shapes(3).Shadow.IncrementOffsetY -3
```


## See also


[ShadowFormat Object](Word.ShadowFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]