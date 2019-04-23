---
title: GradientStyle property (Excel Graph)
keywords: vbagr10.chm3077038
f1_keywords:
- vbagr10.chm3077038
ms.prod: excel
api_name:
- Excel.GradientStyle
ms.assetid: 042a271c-165c-ba10-9702-7db744654760
ms.date: 04/10/2019
localization_priority: Normal
---


# GradientStyle property (Excel Graph)

Returns the gradient style for the specified fill. Read-only **[MsoGradientStyle](office.msogradientstyle.md)**.

## Syntax

_expression_.**GradientStyle**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Remarks

This property is read-only. Use the **[OneColorGradient](excel.onecolorgradient.md)** or **[TwoColorGradient](excel.twocolorgradient.md)** method to set the gradient style for the fill.

## Example

This example sets the chart's fill format so that its gradient style is changed to **msoGradientDiagonalUp** if it was originally **msoGradientDiagonalDown**.

```vb
With myChart.ChartArea.Fill 
 If .Type = msoFillGradient Then 
 If .GradientColorType = msoGradientOneColor Then 
 If .GradientStyle = msoGradientDiagonalDown Then 
 .OneColorGradient msoGradientDiagonalUp, _ 
 .GradientVariant, .GradientDegree 
 End If 
 End If 
 End If 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]