---
title: FullSeriesCollection Object (Excel)
keywords: vbaxl10.chm943072
f1_keywords:
- vbaxl10.chm943072
ms.prod: excel
ms.assetid: 5d7b7e7c-0a74-307b-84f9-56143ceba464
ms.date: 06/08/2017
---


# FullSeriesCollection Object (Excel)

Represents the full set of [Series Object (Excel)](series-object-excel.md) objects in a chart.


## Remarks

The [FullSeriesCollection Object (Excel)](fullseriescollection-object-excel.md) object enables you to get a filtered out[Series Object (Excel)](series-object-excel.md) object and filter it back in. It also enables you to iterate over the full set of Series object, filtered out or visible, programmatically. By having the existing[SeriesCollection Object (Excel)](seriescollection-object-excel.md) object contain only the visible series you can programmatically perform operations on only the visible series. It also prevents Microsoft Excel from breaking existing chart solutions on charts with filtered out data.


## Example

The following example displays a message box with the name of the second [Series Object (Excel)](series-object-excel.md) object in the second chart.


```VB.net
MsgBox Chart(1).FullSeriesCollection.Item(2).Name
```


## Methods



|**Name**|
|:-----|
|[Item](Excel.fullseriescollection.item.md)|

## Properties



|**Name**|
|:-----|
|[Application](Excel.fullseriescollection.application.md)|
|[Count](Excel.fullseriescollection.count.md)|
|[Creator](Excel.fullseriescollection.creator.md)|
|[Parent](fullseriescollection-parent-property-excel.md)|

## See also


[SeriesCollection](seriescollection-object-excel.md)



[Excel Object Model Reference](./overview/object-model-excel-vba-reference.md)
