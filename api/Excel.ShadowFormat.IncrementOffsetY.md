---
title: ShadowFormat.IncrementOffsetY method (Excel)
keywords: vbaxl10.chm114021
f1_keywords:
- vbaxl10.chm114021
ms.prod: excel
api_name:
- Excel.ShadowFormat.IncrementOffsetY
ms.assetid: 0479d9a1-aae1-069c-f692-276291ec54ef
ms.date: 05/14/2019
localization_priority: Normal
---


# ShadowFormat.IncrementOffsetY method (Excel)

Changes the vertical offset of the shadow by the specified number of [points](../language/glossary/vbe-glossary.md#point). Use the **[OffsetY](Excel.ShadowFormat.OffsetY.md)** property to set the absolute vertical shadow offset.


## Syntax

_expression_.**IncrementOffsetY** (_Increment_)

_expression_ A variable that represents a **[ShadowFormat](Excel.ShadowFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|Specifies how far the shadow offset is to be moved vertically, in points. A positive value moves the shadow down; a negative value moves it up.|

## Example

This example moves the shadow on shape three on _myDocument_ up by 3 points.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(3).Shadow.IncrementOffsetY -3
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]