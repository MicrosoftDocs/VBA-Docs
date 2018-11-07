---
title: SortFields.Add2 method (Excel)
keywords: vbaxl10.chm152090
f1_keywords:
- vbaxl10.chm152090
ms.prod: excel
api_name:
- Excel.SortFields.Add2
ms.date: 09/26/2018
---

# SortFields.Add2 method (Excel)

Creates a new sort field and returns a **SortFields** object that can optionally sort data types with the SubField defined.

## Syntax

_expression_. `Add2`( `Key` , `SortOn` , `Order` , `CustomOrder` , `DataOption` , `SubField` )

_expression_ A variable that represents a [SortFields](Excel.SortFields.md) object.

### Parameters

|Name |Required/Optional |Data type |Description|
|:-----|:-----|:-----|:-----|
| _Key_|Required| **Range**|Specifies a key value for the sort.|
| _SortOn_|Optional| **Variant**|The field to sort on.|
| _Order_|Optional| **Variant**|Specifies the sort order.|
| _CustomOrder_|Optional| **Variant**|Specifies if a custom sort order should be used.|
| _DataOption_|Optional| **Variant**|Specifies the data option.|
| _SubField_|Optional| **Variant**|Specifies the field to sort on for a data type (such as "Population" for Geography or "Volume" for Stocks).|

## Return value

SortField

## Remarks

This API includes support for sorting off a SubField from data types, such as Geography or Stocks. [SortFields.Add](Excel.SortFields.Add.md) can also be used if sorting by a data type is not needed.

Unlike in formulas, SubFields do not require brackets to include spaces.

## Examples

This example sorts a Table, "Table1" on "Sheet1" by "Column1", in ascending order based off the SubField "Population" on Geography data types.

[SortFields.Clear](Excel.SortFields.Clear.md) is called before to ensure that the previous sort is cleared so that a new one can be applied.

[Sort](Excel.Sort.md) is called to apply the added sort to "Table1".

```vb
ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").Sort.SortFields.Add2 _
 Key:=Range("Table1[[#All],[Column1]]"), _
 SortOn:=xlSortOnValues, _
 Order:=xlAscending, _
 DataOption:=xlSortNormal, _
 SubField:="Population"
With ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").Sort
 .Header = xlYes
 .MatchCase = False
 .Orientation = xlTopToBottom
 .SortMethod = xlPinYin
 .Apply
End With
```

