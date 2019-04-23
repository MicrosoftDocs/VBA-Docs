---
title: TwoColorGradient method (Excel Graph)
keywords: vbagr10.chm3077636
f1_keywords:
- vbagr10.chm3077636
ms.prod: excel
api_name:
- Excel.TwoColorGradient
ms.assetid: c42ec02c-41a2-ffc4-3d23-20a952b3de7b
ms.date: 04/09/2019
localization_priority: Normal
---


# TwoColorGradient method (Excel Graph)

Sets the specified fill to a two-color gradient.

## Syntax

_expression_.**TwoColorGradient** (_Style_, _Variant_)

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Style_ |Required |**[MsoGradientStyle](office.msogradientstyle.md)** |Specifies the gradient style. Can be one of the **MsoGradientStyle** constants.|
|_Variant_ |Required |**Long**| Specifies the gradient variant. Can be a value from 1 through 4, corresponding to the four variants on the **Gradient** tab in the **Fill Effects** dialog box. If _Style_ is **msoGradientFromCenter**, the _Variant_ argument can only be either 1 or 2.

## Example

This example sets the gradient, background color, and foreground color for the chart area fill on the chart.

```vb
With myChart.ChartArea.Fill 
 .Visible = True 
 .ForeColor.SchemeColor = 15 
 .BackColor.SchemeColor = 17 
 .TwoColorGradient msoGradientHorizontal, 1 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]