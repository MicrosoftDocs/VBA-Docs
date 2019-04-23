---
title: MajorUnit property (Excel Graph)
keywords: vbagr10.chm5207641
f1_keywords:
- vbagr10.chm5207641
ms.prod: excel
api_name:
- Excel.MajorUnit
ms.assetid: 46d4d4e0-f285-2800-f539-72e7acb98948
ms.date: 04/11/2019
localization_priority: Normal
---


# MajorUnit property (Excel Graph)

Returns or sets the major units for the axis. Read/write **Double**.

## Syntax

_expression_.**MajorUnit**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.



## Remarks

Setting this property sets the **[MajorUnitIsAuto](Excel.MajorUnitIsAuto.md)** property to **False**.

Use the **[TickMarkSpacing](Excel.TickMarkSpacing.md)** property to set tick-mark spacing on the category axis.


## Example

This example sets the major and minor units for the value axis.

```vb
With myChart.Axes(xlValue) 
 .MajorUnit = 100 
 .MinorUnit = 20 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]