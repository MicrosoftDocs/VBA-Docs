---
title: TickMarkSpacing Property
keywords: vbagr10.chm5208067
f1_keywords:
- vbagr10.chm5208067
ms.prod: excel
api_name:
- Excel.TickMarkSpacing
ms.assetid: 5c8abc42-b0bc-882d-ebdf-7125a92b121b
ms.date: 06/08/2017
localization_priority: Normal
---


# TickMarkSpacing Property

Returns or sets the number of categories or series between tick marks. Applies only to category and series axes. Read/write  **Long**.


## Remarks

Use the  **[MajorUnit](Excel.MajorUnit.md)** and  **[MinorUnit](Excel.MinorUnit.md)** properties to set tick-mark spacing on the value axis.


## Example

This example sets the number of categories between tick marks on the category axis.


```vb
myChart.Axes(xlCategory).TickMarkSpacing = 10
```


