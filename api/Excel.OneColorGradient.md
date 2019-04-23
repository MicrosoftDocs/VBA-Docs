---
title: OneColorGradient method (Excel Graph)
keywords: vbagr10.chm67157
f1_keywords:
- vbagr10.chm67157
ms.prod: excel
api_name:
- Excel.OneColorGradient
ms.assetid: 7e572d28-2905-2c6b-5e62-1f763bba7f89
ms.date: 04/09/2019
localization_priority: Normal
---


# OneColorGradient method (Excel Graph)

Sets the specified fill to a one-color gradient.

## Syntax

_expression_.**OneColorGradient** (_Style_, _Variant_, _Degree_)

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Style_ |Required |**[MsoGradientStyle](office.msogradientstyle.md)** |The gradient style for the specified fill. Can be one of the **MsoGradientStyle** constants.|
|_Variant_ |Required | **Long**|The gradient variant for the specified fill. Can be a value from 1 through 4, corresponding to the four variants listed on the **Gradient** tab in the **Fill Effects** dialog box. If _Style_ is **msoGradientFromCenter**, the _Variant_ argument can only be 1 or 2.|
|_Degree_ |Required | **Single**|The gradient degree for the specified fill. Can be a value from 0.0 (dark) through 1.0 (light).|

## Example

This example sets the chart's fill format.

```vb
With myChart.ChartArea.Fill 
 If .Type = msoFillGradient Then 
 If .GradientColorType = msoGradientOneColor Then 
 .OneColorGradient Style:=msoGradientFromCorner, _ 
 Variant:=1, Degree:=0.3 
 End If 
 End If 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]