---
title: GradientDegree property (Excel Graph)
keywords: vbagr10.chm67177
f1_keywords:
- vbagr10.chm67177
ms.prod: excel
api_name:
- Excel.GradientDegree
ms.assetid: 6f325dc0-5f6c-7a55-52fa-55eeb15ccfe6
ms.date: 04/10/2019
localization_priority: Normal
---


# GradientDegree property (Excel Graph)

Returns the gradient degree of the specified one-color shaded fill as a floating-point value from 0.0 (dark) through 1.0 (light). Read-only **Single**.

## Syntax

_expression_.**GradientDegree**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Remarks

This property is read-only. Use the **[OneColorGradient](excel.onecolorgradient.md)** method to set the gradient degree for the fill.

## Example

This example sets the chart's fill format so that its gradient degree is at least 0.3.

```vb
With myChart.ChartArea.Fill 
 If .Type = msoFillGradient Then 
 If .GradientColorType = msoGradientOneColor Then 
 If .GradientDegree < 0.3 Then 
 .OneColorGradient .GradientStyle, _ 
 .GradientVariant, 0.3 
 End If 
 End If 
 End If 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]