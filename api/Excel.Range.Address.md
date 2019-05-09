---
title: Range.Address property (Excel)
keywords: vbaxl10.chm144076
f1_keywords:
- vbaxl10.chm144076
ms.prod: excel
api_name:
- Excel.Range.Address
ms.assetid: aaa2432e-9bb1-4a48-3868-86455bc53938
ms.date: 05/10/2019
localization_priority: Priority
---


# Range.Address property (Excel)

Returns a **String** value that represents the range reference in the language of the macro.


## Syntax

_expression_.**Address** (_RowAbsolute_, _ColumnAbsolute_, _ReferenceStyle_, _External_, _RelativeTo_)

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _RowAbsolute_|Optional| **Variant**| **True** to return the row part of the reference as an absolute reference. The default value is **True**.|
| _ColumnAbsolute_|Optional| **Variant**| **True** to return the column part of the reference as an absolute reference. The default value is **True**.|
| _ReferenceStyle_|Optional| **[XlReferenceStyle](Excel.XlReferenceStyle.md)**|The reference style. The default value is **xlA1**.|
| _External_|Optional| **Variant**| **True** to return an external reference. **False** to return a local reference. The default value is **False**.|
| _RelativeTo_|Optional| **Variant**|If _RowAbsolute_ and _ColumnAbsolute_ are **False**, and _ReferenceStyle_ is **xlR1C1**, you must include a starting point for the relative reference. This argument is a **Range** object that defines the starting point.<br/><br/>**NOTE**: Testing with Excel VBA 7.1 shows that an explicit starting point is not mandatory. There appears to be a default reference of $A$1.|

## Remarks

If the reference contains more than one cell, _RowAbsolute_ and _ColumnAbsolute_ apply to all rows and columns.

## Example

The following example displays four different representations of the same cell address on Sheet1. The comments in the example are the addresses that will be displayed in the message boxes.

```vb
Set mc = Worksheets("Sheet1").Cells(1, 1) 
MsgBox mc.Address() ' $A$1 
MsgBox mc.Address(RowAbsolute:=False) ' $A1 
MsgBox mc.Address(ReferenceStyle:=xlR1C1) ' R1C1 
MsgBox mc.Address(ReferenceStyle:=xlR1C1, _ 
 RowAbsolute:=False, _ 
 ColumnAbsolute:=False, _ 
 RelativeTo:=Worksheets(1).Cells(3, 3)) ' R[-2]C[-2]
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
