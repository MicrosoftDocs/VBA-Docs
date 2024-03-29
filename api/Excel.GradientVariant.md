---
title: GradientVariant property (Excel Graph)
keywords: vbagr10.chm3077040
f1_keywords:
- vbagr10.chm3077040
api_name:
- Excel.GradientVariant
ms.assetid: 7aa7c237-9dc7-8588-6b19-68b98f2a3662
ms.date: 04/10/2019
ms.localizationpriority: medium
---


# GradientVariant property (Excel Graph)

Returns the shade variant for the specified fill as an integer value from 1 through 4. The values for this property correspond to the gradient variants (numbered from left to right and from top to bottom) listed on the **Gradient** tab in the **Fill Effects** dialog box. Read-only **Long**.

## Syntax

_expression_.**GradientVariant**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Remarks

This property is read-only. Use the **[OneColorGradient](excel.onecolorgradient.md)** or **[TwoColorGradient](excel.twocolorgradient.md)** method to set the gradient variant for the fill.

## Example

This example sets the chart's fill format so that it's displayed using the second shade variant if it's currently using the first shade variant.

```vb
With myChart.ChartArea.Fill 
 If .Type = msoFillGradient Then 
 If .GradientColorType = msoGradientOneColor Then 
 If .GradientVariant = 1 Then 
 .OneColorGradient .GradientStyle, 2, _ 
 .GradientDegree 
 End If 
 End If 
 End If 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]