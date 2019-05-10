---
title: Series.DataLabels method (Excel)
keywords: vbaxl10.chm578079
f1_keywords:
- vbaxl10.chm578079
ms.prod: excel
api_name:
- Excel.Series.DataLabels
ms.assetid: bde8faa1-269c-1dbe-e39e-3701a634f214
ms.date: 05/11/2019
localization_priority: Normal
---


# Series.DataLabels method (Excel)

Returns an object that represents either a single data label (a **[DataLabel](Excel.DataLabel(object).md)** object) or a collection of all the data labels for the series (a **[DataLabels](Excel.DataLabels(object).md)** collection).


## Syntax

_expression_.**DataLabels** (_Index_)

_expression_ A variable that represents a **[Series](Excel.Series(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The number of the data label.|

## Return value

Object


## Remarks

If the series has the **Show Value** option turned on for the data labels, the returned collection can contain up to one label for each point. Data labels can be turned on or off for individual points in the series.

If the series is on an area chart and has the **Show Label** option turned on for the data labels, the returned collection contains only a single label, which is the label for the area series.


## Example

This example sets the data labels for series one on Chart1 to show their key, assuming that their values are visible when the example runs.

```vb
With Charts("Chart1").SeriesCollection(1) 
 .HasDataLabels = True 
 With .DataLabels 
 .ShowLegendKey = True 
 .Type = xlValue 
 End With 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]