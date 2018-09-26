---
title: SortFields.Add Method (Excel)
keywords: vbaxl10.chm845073
f1_keywords:
- vbaxl10.chm845073
ms.prod: excel
api_name:
- Excel.SortFields.Add
ms.date: 09/21/2018
---

# SortFields.Add Method (Excel)

Creates a new sort field and returns a  **SortFields** object.

## Syntax

 _expression_. `Add`( `_Key_` , `_SortOn_` , `_Order_` , `_CustomOrder_` , `_DataOption_` )

 _expression_ A variable that represents a [SortFields](./Excel.SortFields.md) object.

### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Key_|Required| **Range**|Specifies a key value for the sort.|
| _SortOn_|Optional| **Variant**|The field to sort on.|
| _Order_|Optional| **Variant**|Specifies the sort order.|
| _CustomOrder_|Optional| **Variant**|Specifies if a custom sort order should be used.|
| _DataOption_|Optional| **Variant**|Specifies the data option.|

### Return Value

SortField

## Remarks

This API does not include support for Data Types, such as Geography or Stocks. Please reference [SortFields.Add2](Excel.SortFields.Add2.md) to define a sort order based off a SubField of one of these types.

## Examples

This example sorts a Table, "Table1" on "Sheet1" by "Column1", in ascending order.

[SortFields.Clear](Excel.SortFields.Clear.md) is called before to ensure the previous sort is cleared so a new one can be applied.

[Sort](Excel.Sort.md) is called to apply the added sort to "Table1".

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

## See also

[SortFields.Add2](Excel.SortFields.Add2.md)

[SortFields.Clear](Excel.SortFields.Clear.md)

[Sort](Excel.Sort.md)

[SortFields Object](Excel.SortFields.md)