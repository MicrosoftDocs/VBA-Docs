---
title: Series.ApplyDataLabels method (Excel)
keywords: vbaxl10.chm578122
f1_keywords:
- vbaxl10.chm578122
api_name:
- Excel.Series.ApplyDataLabels
ms.assetid: 959a4d12-ed48-48fc-04cf-7a1880cd7e1f
ms.date: 05/11/2019
ms.localizationpriority: medium
---


# Series.ApplyDataLabels method (Excel)

Applies data labels to a series.


## Syntax

_expression_.**ApplyDataLabels** (_Type_, _LegendKey_, _AutoText_, _HasLeaderLines_, _ShowSeriesName_, _ShowCategoryName_, _ShowValue_, _ShowPercentage_, _ShowBubbleSize_, _Separator_)

_expression_ A variable that represents a **[Series](Excel.Series(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Optional| **[XlDataLabelsType](Excel.XlDataLabelsType.md)**|The type of data label to apply.|
| _LegendKey_|Optional| **Variant**| **True** to show the legend key next to the point. The default value is **False**.|
| _AutoText_|Optional| **Variant**| **True** if the object automatically generates appropriate text based on content.|
| _HasLeaderLines_|Optional| **Variant**|For the **[Chart](Excel.Chart(object).md)** and **Series** objects, **True** if the series has leader lines.|
| _ShowSeriesName_|Optional| **Variant**|Pass a **Boolean** value to enable or disable the series name for the data label.|
| _ShowCategoryName_|Optional| **Variant**|Pass a **Boolean** value to enable or disable the category name for the data label.|
| _ShowValue_|Optional| **Variant**|Pass a **Boolean** value to enable or disable the value for the data label.|
| _ShowPercentage_|Optional| **Variant**|Pass a **Boolean** value to enable or disable the percentage for the data label.|
| _ShowBubbleSize_|Optional| **Variant**|Pass a **Boolean** value to enable or disable the bubble size for the data label.|
| _Separator_|Optional| **Variant**|The separator for the data label.|


## Example

This example applies category labels to series one on Chart1.

```vb
Charts("Chart1").SeriesCollection(1). _ 
 ApplyDataLabels Type:=xlDataLabelsShowLabel
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]