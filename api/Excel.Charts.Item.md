---
title: Charts.Item property (Excel)
keywords: vbaxl10.chm217076
f1_keywords:
- vbaxl10.chm217076
ms.prod: excel
api_name:
- Excel.Charts.Item
ms.assetid: 792e3562-7d70-4356-7072-fa09cb40ec47
ms.date: 04/20/2019
localization_priority: Normal
---


# Charts.Item property (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[Charts](Excel.Charts.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number of the object.|


## Example

This example sets the number of units that the trendline on Chart1 extends forward and backward. The example should be run on a 2D column chart that contains a single series with a trendline.

```vb
With Charts.Item("Chart1").SeriesCollection(1).Trendlines(1) 
 .Forward = 5 
 .Backward = .5 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]