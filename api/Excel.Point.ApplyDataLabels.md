---
title: Point.ApplyDataLabels method (Excel)
keywords: vbaxl10.chm576100
f1_keywords:
- vbaxl10.chm576100
ms.prod: excel
api_name:
- Excel.Point.ApplyDataLabels
ms.assetid: f242eef7-75ed-868f-bb8d-d42838cc9ff0
ms.date: 06/08/2017
localization_priority: Normal
---


# Point.ApplyDataLabels method (Excel)

Applies data labels to a point.


## Syntax

_expression_. `ApplyDataLabels`( `_Type_` , `_LegendKey_` , `_AutoText_` , `_HasLeaderLines_` , `_ShowSeriesName_` , `_ShowCategoryName_` , `_ShowValue_` , `_ShowPercentage_` , `_ShowBubbleSize_` , `_Separator_` )

_expression_ A variable that represents a [Point](Excel.Point-graph-object.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Optional| **[xlDataLabelsType](Excel.XlDataLabelsType.md)**|The type of data label to apply.|
| _LegendKey_|Optional| **Variant**| **True** to show the legend key next to the point. The default value is **False**.|
| _AutoText_|Optional| **Variant**| **True** if the object automatically generates appropriate text based on content.|
| _HasLeaderLines_|Optional| **Variant**|For the  **[Chart](Excel.Chart(object).md)** and **[Series](Excel.Series(object).md)** objects, **True** if the series has leader lines.|
| _ShowSeriesName_|Optional| **Variant**|Pass a boolean value to enable or disable the series name for the data label.|
| _ShowCategoryName_|Optional| **Variant**|Pass a boolean value to enable or disable the category name for the data label.|
| _ShowValue_|Optional| **Variant**|Pass a boolean value to enable or disable the value for the data label.|
| _ShowPercentage_|Optional| **Variant**|Pass a boolean value to enable or disable the percentage for the data label.|
| _ShowBubbleSize_|Optional| **Variant**|Pass a boolean value to enable or disable the bubble size for the data label.|
| _Separator_|Optional| **Variant**|The separator for the data label.|

## Remarks





| **xlDataLabelsType** can be one of these **xlDataLabelsType** constants.|
| **xlDataLabelsShowBubbleSizes**. The bubble size for the data label.|
| **xlDataLabelsShowLabelAndPercent**. Percentage of the total, and category for the point. Available only for pie charts and doughnut charts.|
| **xlDataLabelsShowPercent**. Percentage of the total. Available only for pie charts and doughnut charts.|
| **xlDataLabelsShowLabel**. Category for the point.|
| **xlDataLabelsShowNone**. No data labels.|
| **xlDataLabelsShowValue**. _default_ . Value for the point (assumed if this argument isn't specified).|

## Example

This example applies category labels to series one in Chart1.


```vb
Charts("Chart1").SeriesCollection(1). _ 
 ApplyDataLabels Type:=xlDataLabelsShowLabel
```


## See also


[Point Object](Excel.Point(object).md)

