---
title: Axis.MaximumScale property (Excel)
keywords: vbaxl10.chm561088
f1_keywords:
- vbaxl10.chm561088
ms.prod: excel
api_name:
- Excel.Axis.MaximumScale
ms.assetid: 384e52b5-561e-aa07-910c-67ee0fb07ba0
ms.date: 04/13/2019
localization_priority: Normal
---


# Axis.MaximumScale property (Excel)

Returns or sets the maximum value on the value axis. Read/write **Double**.


## Syntax

_expression_.**MaximumScale**

_expression_ A variable that represents an **[Axis](Excel.Axis(object).md)** object.


## Remarks

Setting this property sets the **[MaximumScaleIsAuto](Excel.Axis.MaximumScaleIsAuto.md)** property to **False**.


## Example

This example sets the minimum and maximum values for the value axis on Chart1.

```vb
With Charts("Chart1").Axes(xlValue) 
 .MinimumScale = 10 
 .MaximumScale = 120 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]