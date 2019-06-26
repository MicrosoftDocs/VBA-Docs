---
title: Chart.ApplyDataLabels method (Word)
keywords: vbawd10.chm79366018
f1_keywords:
- vbawd10.chm79366018
ms.prod: word
api_name:
- Word.Chart.ApplyDataLabels
ms.assetid: 3d4ce40f-7194-ad96-4ae6-15434c6dd491
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.ApplyDataLabels method (Word)

Applies data labels to all the series in a chart.


## Syntax

_expression_.**ApplyDataLabels** (_Type_, _LegendKey_, _AutoText_, _HasLeaderLines_, _ShowSeriesName_, _ShowCategoryName_, _ShowValue_, _ShowPercentage_, _ShowBubbleSize_, _Separator_)

_expression_ A variable that represents a **[Chart](Word.Chart.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Optional| **[XlDataLabelsType](Word.xldatalabelstype.md)**|One of the enumeration values that specifies the type of data label to apply. Can be one of the **xlDataLabelsType** constants.|
| _LegendKey_|Optional| **Variant**| **True** to show the legend key next to the point. The default is **False**.|
| _AutoText_|Optional| **Variant**| **True** if the object automatically generates appropriate text based on content.|
| _HasLeaderLines_|Optional| **Variant**|For the  **[Chart](Word.Chart.md)** and **[Series](Word.Series.md)** objects, **True** if the series has leader lines.|
| _ShowSeriesName_|Optional| **Variant**| **True** to enable the series name for the data label; otherwise, **False**.|
| _ShowCategoryName_|Optional| **Variant**| **True** to enable the category name for the data label; otherwise, **False**.|
| _ShowValue_|Optional| **Variant**| **True** to enable the value for the data label; otherwise, **False**.|
| _ShowPercentage_|Optional| **Variant**| **True** to enable the percentage for the data label; otherwise, **False**.|
| _ShowBubbleSize_|Optional| **Variant**| **True** to enable the bubble size for the data label; otherwise, **False**.|
| _Separator_|Optional| **Variant**|The separator for the data label.|


## Example

The following example applies category labels to the first series of the first chart in the active document.

```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1). _ 
 ApplyDataLabels Type:=xlDataLabelsShowLabel 
 End If 
End With
```

> [!NOTE] 
> If you use a string for the Separator parameter, you will get a string as the separator. If you use  **xlDataLabelSeparatorDefault** (= 1), you will get the default data label separator, which is either a comma or a newline, depending on the data label.When a value of 1 is returned, it indicates that the user has not changed the default separator, which is a comma (,). You can also pass a value of 1 to change the separator back to the default separator.The chart must first be active before you can access the data labels programmatically; otherwise, a run-time error occurs.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]