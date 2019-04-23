---
title: Insert method (Excel Graph)
keywords: vbagr10.chm3077620
f1_keywords:
- vbagr10.chm3077620
ms.prod: excel
api_name:
- Excel.Insert
ms.assetid: 5f6a5961-9278-a2fa-6f08-4360646a7566
ms.date: 04/09/2019
localization_priority: Normal
---


# Insert method (Excel Graph)

Inserts a cell or a range of cells into the datasheet and shifts other cells away to make space.

## Syntax

_expression_.**Insert** (_Shift_)

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Shift_ |Optional| **[XlInsertShiftDirection](excel.xlinsertshiftdirection.md)**|Specifies which way to shift the cells. Can be one of these **XlInsertShiftDirection** constants: **xlShiftToRight** or **xlShiftDown**. If this argument is omitted, Graph decides based on the shape of the range.|

## Example

This example inserts a new row before row four on the datasheet.

```vb
myChart.Application.DataSheet.Rows(4).Insert
```

<br/>

This example inserts new cells at the range A1:C5 on the datasheet and shifts cells downward.

```vb
Set mySheet = myChart.Application.DataSheet 
mySheet.Range("A1:C5").Insert Shift:=xlShiftDown
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]