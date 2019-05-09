---
title: Series.ApplyDataLabels method (PowerPoint)
keywords: vbapp10.chm716004
f1_keywords:
- vbapp10.chm716004
ms.prod: powerpoint
api_name:
- PowerPoint.Series.ApplyDataLabels
ms.assetid: d8f4752f-1ff4-8a42-4b9f-12d81814f4f2
ms.date: 06/08/2017
localization_priority: Normal
---


# Series.ApplyDataLabels method (PowerPoint)

Applies data labels to a series.


## Syntax

_expression_.**ApplyDataLabels** (`Type`, `LegendKey`, `AutoText`, `HasLeaderLines`, `ShowSeriesName`, `ShowCategoryName`, `ShowValue`, `ShowPercentage`, `ShowBubbleSize`, `Separator`)

_expression_ A variable that represents a [Series](PowerPoint.Series.md) object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Optional|**[XlDataLabelsType](PowerPoint.XlDataLabelsType.md)**|The type of data label to apply.|
| _LegendKey_|Optional|**Variant**|**True** to show the legend key next to the point. The default is **False**.|
| _AutoText_|Optional|**Variant**|**True** if the object automatically generates appropriate text based on content.|
| _HasLeaderLines_|Optional|**Variant**|For the **[Chart](PowerPoint.Chart.md)** and **[Series](PowerPoint.Series.md)** objects, **True** if the series has leader lines.|
| _ShowSeriesName_|Optional|**Variant**|**True** to enable the series name for the data label; otherwise, **False**.|
| _ShowCategoryName_|Optional|**Variant**|**True** to enable the category name for the data label; otherwise, **False**.|
| _ShowValue_|Optional|**Variant**|**True** to enable the value for the data label; otherwise, **False**.|
| _ShowPercentage_|Optional|**Variant**|**True** to enable the percentage for the data label; otherwise, **False**.|
| _ShowBubbleSize_|Optional|**Variant**|**True** to enable the bubble size for the data label; otherwise, **False**.|
| _Separator_|Optional|**Variant**|The separator for the data label.|

## Remarks

The Type parameter can be one of the following **xlDataLabelsType** constants:

- **xlDataLabelsShowBubbleSizes** The bubble size for the data label.
    
- **xlDataLabelsShowLabelAndPercent** The percentage of the total, and the category for the point. Available only for pie charts and doughnut charts.
    
- **xlDataLabelsShowPercent** The percentage of the total. Available only for pie charts and doughnut charts.
    
- **xlDataLabelsShowLabel** The category for the point.
    
- **xlDataLabelsShowNone** No data labels.
    
- **xlDataLabelsShowValue** (Default) The value for the point (assumed if this argument is not specified).
    

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


## See also

- [Series Object](PowerPoint.Series.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]