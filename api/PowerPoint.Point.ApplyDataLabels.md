---
title: Point.ApplyDataLabels method (PowerPoint)
keywords: vbapp10.chm714004
f1_keywords:
- vbapp10.chm714004
ms.prod: powerpoint
api_name:
- PowerPoint.Point.ApplyDataLabels
ms.assetid: 49bd00ab-d1d1-563f-b5ce-e0a575a97a5c
ms.date: 06/08/2017
localization_priority: Normal
---


# Point.ApplyDataLabels method (PowerPoint)

Applies data labels to a point.


## Syntax

_expression_.**ApplyDataLabels** (_Type_, _LegendKey_, _AutoText_, _HasLeaderLines_, _ShowSeriesName_, _ShowCategoryName_, _ShowValue_, _ShowPercentage_, _ShowBubbleSize_, _Separator_)

_expression_ A variable that represents a '[Point](PowerPoint.Point.md)' object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Optional|**[XlDataLabelsType](PowerPoint.XlDataLabelsType.md)**|The type of data label to apply. Can be one of the **xlDataLabelsType** constants.|
| _LegendKey_|Optional|**Variant**|**True** to show the legend key next to the point. The default is **False**.|
| _AutoText_|Optional|**Variant**|**True** if the object automatically generates appropriate text based on content.|
| _HasLeaderLines_|Optional|**Variant**|For the  **[Chart](PowerPoint.Chart.md)** and **[Series](PowerPoint.Series.md)** objects, **True** if the series has leader lines.|
| _ShowSeriesName_|Optional|**Variant**|**True** to enable the series name for the data label; otherwise, **False**.|
| _ShowCategoryName_|Optional|**Variant**|**True** to enable the category name for the data label; otherwise, **False**.|
| _ShowValue_|Optional|**Variant**|**True** to enable the value for the data label; otherwise, **False**.|
| _ShowPercentage_|Optional|**Variant**|**True** to enable the percentage for the data label; otherwise, **False**.|
| _ShowBubbleSize_|Optional|**Variant**|**True** to enable the bubble size for the data label; otherwise, **False**.|
| _Separator_|Optional|**Variant**|The separator for the data label.|



## Example

> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

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