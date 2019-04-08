---
title: Delete method (Excel Graph)
keywords: vbagr10.chm3077617
f1_keywords:
- vbagr10.chm3077617
ms.prod: excel
ms.assetid: f5bc861f-67e4-05e9-765f-d9ed34e0e936
ms.date: 04/09/2019
localization_priority: Normal
---


# Delete method (Excel Graph)

The **Delete** method as it applies to all objects in the **Applies To** list except the **Range** object, and then to the **Range** object. 

## All objects except the Range object

Applies to all objects in the **Applies To** list except the **Range** object.

Deletes the specified object.

### Syntax

_expression_.**Delete**

_expression_ Required. An expression that returns one of the above objects.



## Range object

Applies to the **Range** object.

Deletes the specified object.

### Syntax

_expression_.**Delete** (_Shift_)

_expression_ Required. An expression that returns a **[Range](excel.range-graph-object.md)** object. 


### Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Shift_ |_Optional|**[XlDeleteShiftDirection](excel.xldeleteshiftdirection.md)** |Used only with **Range** objects. Specifies how to shift cells to replace deleted cells.<br/><br/> Can be one of these **XlDeleteShiftDirection** constants: **xlShiftToLeft** or **xlShiftUp**. <br/><br/>If this argument is omitted, Graph decides how to shift cells based on the shape of the specified range.|

### Remarks

Deleting a **Point** or **LegendKey** object deletes the entire series.


## Example

This example deletes cells A1:D10 on the datasheet and shifts the remaining cells to the left.

```vb
Set mySheet = myChart.Application.DataSheet 
mySheet.Range("A1:D10").Delete Shift:=xlShiftToLeft
```

<br/>

This example deletes the chart title.

```vb
myChart.ChartTitle.Delete
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]