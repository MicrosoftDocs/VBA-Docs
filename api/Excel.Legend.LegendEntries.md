---
title: Legend.LegendEntries Method (Excel)
keywords: vbaxl10.chm622079
f1_keywords:
- vbaxl10.chm622079
ms.prod: excel
api_name:
- Excel.Legend.LegendEntries
ms.assetid: 6b20827c-7196-e1d7-485f-954b0ea90f58
ms.date: 06/08/2017
---


# Legend.LegendEntries Method (Excel)

Returns an object that represents either a single legend entry (a  **[LegendEntry](Excel.LegendEntry(object).md)** object) or a collection of legend entries (a **[LegendEntries](Excel.LegendEntries(object).md)** object) for the legend.


## Syntax

 _expression_. `LegendEntries`( `_Index_` )

 _expression_ A variable that represents a [Legend](Excel.Legend(Graph property).md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The number of the legend entry.|

### Return Value

Object


## Example

This example sets the font for legend entry one on Chart1.


```vb
Charts("Chart1").Legend.LegendEntries(1).Font.Name = "Arial"
```


## See also


[Legend Object](Excel.Legend(object).md)

