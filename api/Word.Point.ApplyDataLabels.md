---
title: Point.ApplyDataLabels method (Word)
keywords: vbawd10.chm262145922
f1_keywords:
- vbawd10.chm262145922
ms.prod: word
api_name:
- Word.Point.ApplyDataLabels
ms.assetid: 86199dd9-f27b-c9bd-58e5-59b0873b88e3
ms.date: 06/08/2017
localization_priority: Normal
---


# Point.ApplyDataLabels method (Word)

Applies data labels to a point.


## Syntax

_expression_.**ApplyDataLabels** (_Type_, _LegendKey_, _AutoText_, _HasLeaderLines_, _ShowSeriesName_, _ShowCategoryName_, _ShowValue_, _ShowPercentage_, _ShowBubbleSize_, _Separator_)

_expression_ A variable that represents a '[Point](Word.Point.md)' object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Optional| **[XlDataLabelsType](Word.xldatalabelstype.md)**|The type of data label to apply. Can be one of the **xlDataLabelsType** constants.|
| _LegendKey_|Optional| **Variant**| **True** to show the legend key next to the point. The default is **False**.|
| _AutoText_|Optional| **Variant**| **True** if the object automatically generates appropriate text based on content.|
| _HasLeaderLines_|Optional| **Variant**|For the **[Chart](Word.Chart.md)** and **[Series](Word.Series.md)** objects, **True** if the series has leader lines.|
| _ShowSeriesName_|Optional| **Variant**| **True** to enable the series name for the data label; otherwise, **False**.|
| _ShowCategoryName_|Optional| **Variant**| **True** to enable the category name for the data label; otherwise, **False**.|
| _ShowValue_|Optional| **Variant**| **True** to enable the value for the data label; otherwise, **False**.|
| _ShowPercentage_|Optional| **Variant**| **True** to enable the percentage for the data label; otherwise, **False**.|
| _ShowBubbleSize_|Optional| **Variant**| **True** to enable the bubble size for the data label; otherwise, **False**.|
| _Separator_|Optional| **Variant**|The separator for the data label.|


## Example

The following example applies category labels to series one of the first chart in the active document.

```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1). _ 
 ApplyDataLabels Type:=xlDataLabelsShowLabel 
 End If 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]