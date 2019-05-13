---
title: ShadowFormat.IncrementOffsetX method (Excel)
keywords: vbaxl10.chm114020
f1_keywords:
- vbaxl10.chm114020
ms.prod: excel
api_name:
- Excel.ShadowFormat.IncrementOffsetX
ms.assetid: eaa71500-16dd-5df1-cf32-920ab71d77bb
ms.date: 05/14/2019
localization_priority: Normal
---


# ShadowFormat.IncrementOffsetX method (Excel)

Changes the horizontal offset of the shadow by the specified number of [points](../language/glossary/vbe-glossary.md#point). Use the **[OffsetX](Excel.ShadowFormat.OffsetX.md)** property to set the absolute horizontal shadow offset.


## Syntax

_expression_.**IncrementOffsetX** (_Increment_)

_expression_ A variable that represents a **[ShadowFormat](Excel.ShadowFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|Specifies how far the shadow offset is to be moved horizontally, in points. A positive value moves the shadow to the right; a negative value moves it to the left.|

## Example

This example moves the shadow on shape three on _myDocument_ to the left by 3 points.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(3).Shadow.IncrementOffsetX -3
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]