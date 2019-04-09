---
title: Patterned method (Excel Graph)
keywords: vbagr10.chm67164
f1_keywords:
- vbagr10.chm67164
ms.prod: excel
api_name:
- Excel.Patterned
ms.assetid: a492f089-cd6e-e7c3-2b25-7bcfadde4319
ms.date: 04/09/2019
localization_priority: Normal
---


# Patterned method (Excel Graph)

Sets a pattern for the specified fill.

## Syntax

_expression_.**Patterned** (_Pattern_)

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Pattern_ | Required |**[MsoPatternType](office.msopatterntype.md)** |The type of pattern. Can be one of the **MsoPatternType** constants.|


## Example

This example sets the fill pattern.

```vb
With myChart.ChartArea.Fill 
 .Patterned msoPatternDiagonalBrick 
 .Visible = True 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]