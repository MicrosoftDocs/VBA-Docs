---
title: Axis.ScaleType property (Excel)
keywords: vbaxl10.chm561097
f1_keywords:
- vbaxl10.chm561097
ms.prod: excel
api_name:
- Excel.Axis.ScaleType
ms.assetid: 6b217c08-24c4-1ce0-9b7b-96469183002f
ms.date: 04/13/2019
localization_priority: Normal
---


# Axis.ScaleType property (Excel)

Returns or sets the value axis scale type. Read/write **XlScaleType**.


## Syntax

_expression_.**ScaleType**

_expression_ A variable that represents an **[Axis](Excel.Axis(object).md)** object.


## Remarks

**XlScaleType** can be one of the **[XlScaleType](Excel.XlScaleType.md)** constants.

A logarithmic scale uses base 10 logarithms.


## Example

This example sets the value axis on Chart1 to use a logarithmic scale.

```vb
Charts("Chart1").Axes(xlValue).ScaleType = xlScaleLogarithmic
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]