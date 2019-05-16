---
title: SortFields.Add method (Excel)
keywords: vbaxl10.chm845073
f1_keywords:
- vbaxl10.chm845073
ms.prod: excel
api_name:
- Excel.SortFields.Add
ms.date: 05/16/2019
localization_priority: Normal
---

# SortFields.Add method (Excel)

Creates a new sort field and returns a **SortFields** object.

## Syntax

_expression_.**Add** (_Key_, _SortOn_, _Order_, _CustomOrder_, _DataOption_)

_expression_ A variable that represents a **[SortFields](Excel.SortFields.md)** object.

## Parameters

|Name |Required/Optional |Data type |Description|
|:-----|:-----|:-----|:-----|
| _Key_|Required| **Range**|Specifies a key value for the sort.|
| _SortOn_|Optional| **Variant**|The field to sort on.|
| _Order_|Optional| **Variant**|Specifies the sort order.|
| _CustomOrder_|Optional| **Variant**|Specifies if a custom sort order should be used.|
| _DataOption_|Optional| **Variant**|Specifies the data option.|

## Return value

**SortField**

## Remarks

This API does not include support for data types, such as Geography or Stocks. To define a sort order based off a SubField of one of these types, see the **[Add2](Excel.SortFields.Add2.md)** method.

## Example

This example sorts a table, Table1 on Sheet1, by Column1 in ascending order.

The **[Clear](Excel.SortFields.Clear.md)** method is called before to ensure that the previous sort is cleared so that a new one can be applied.

The **[Sort](Excel.Sort.md)** object is called to apply the added sort to Table1.


```vb
ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").Sort.SortFields.Add _
 Key:=Range("Table1[[#All],[Column1]]"), _
 SortOn:=xlSortOnValues, _
 Order:=xlAscending, _
 DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").Sort
 .Header = xlYes
 .MatchCase = False
 .Orientation = xlTopToBottom
 .SortMethod = xlPinYin
 .Apply
End With
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
