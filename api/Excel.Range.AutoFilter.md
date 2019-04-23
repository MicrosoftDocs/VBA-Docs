---
title: Range.AutoFilter method (Excel)
keywords: vbaxl10.chm144084
f1_keywords:
- vbaxl10.chm144084
ms.prod: excel
api_name:
- Excel.Range.AutoFilter
ms.date: 09/26/2018
localization_priority: Priority
---

# Range.AutoFilter method (Excel)

Filters a list using the AutoFilter.

## Syntax

_expression_. `AutoFilter`( `Field` , `Criteria1` , `Operator` , `Criteria2` , `SubField` , `VisibleDropDown`)

_expression_ An expression that returns a **[Range](Excel.Range(object).md)** object.



## Parameters

|Name |Required/Optional |Data type |Description|
|:-----|:-----|:-----|:-----|
| _Field_|Optional| **Variant**| The integer offset of the field on which you want to base the filter (from the left of the list; the leftmost field is field one).|
| _Criteria1_|Optional| **Variant**|The criteria (a string; for example, "101"). Use `"="` to find blank fields, `"<>"` to find non-blank fields, and `"><"` to select (No Data) fields in data types. If this argument is omitted, the criteria is All. If  _Operator_ is **xlTop10Items**, _Criteria1_ specifies the number of items (for example, "10").|
| _Operator_|Optional| **[XlAutoFilterOperator](Excel.XlAutoFilterOperator.md)**|One of the constants of XlAutoFilterOperator specifying the type of filter.|
| _Criteria2_|Optional| **Variant**|The second criteria (a string). Used with  _Criteria1_ and _Operator_ to construct compound criteria.|
| _SubField_|Optional| **Variant**|The Field from a data type on which to apply the Criteria (for example, the "Population" field from Geography or "Volume" field from Stocks). Omitting this value targets the "(Display Value)".|
| _VisibleDropDown_|Optional| **Variant**| **True** to display the AutoFilter drop-down arrow for the filtered field. **False** to hide the AutoFilter drop-down arrow for the filtered field. **True** by default.|

## Return value

Variant

## Remarks

If you omit all the arguments, this method simply toggles the display of the AutoFilter drop-down arrows in the specified range.

Excel for Mac does not support this method. Similar methods on Selection and ListObject are supported.

Unlike in formulas, Subfields do not require brackets to include spaces.

## Examples

This example filters a list starting in cell A1 on Sheet1 to display only the entries in which field one is equal to the string "Otis". The drop-down arrow for field one will be hidden.

```vb
Worksheets("Sheet1").Range("A1").AutoFilter _
 Field:=1, _
 Criteria1:="Otis", _
 VisibleDropDown:=False
```

<br/>

This example filters a list starting in cell A1 on Sheet1 to display only the entries in which the values of field one contain a SubField, "Admin Division 1 (State/province/other)", where the value is "Washington".

```vb
Worksheets("Sheet1").Range("A1").AutoFilter _
 Field:=1, _
 Criteria1:="Washington", _
 SubField:="Admin Division 1 (State/province/other)"
```

<br/>

This example filters a Table, "Table1", on Sheet1 to display only the entries in which the values of field one have a "(Display Value)" that is either "1", "3", "Seattle", or "Redmond".

```vb
Worksheets("Sheet1").ListObjects("Table1").Range.AutoFilter _
 Field:=1, _
 Criteria1:=Array("1", "3", "Seattle", "Redmond"), _
 Operator:=xlFilterValues
```

<br/>

Data types can apply multiple SubField filters. This example filters a Table, "Table1", on Sheet1 to display only the entries in which the values of field one contain a SubField, "Time zone(s)", where the value is "Pacific Time Zone", and where the SubField "Date Founded" is either "1851" or there is "(No Data)".

```vb
Worksheets("Sheet1").ListObjects("Table1").Range.AutoFilter _
 Field:=1, _
 Criteria1:="Pacific Time Zone", _
 SubField:="Time Zone(s)"
Worksheets("Sheet1").ListObjects("Table1").Range.AutoFilter _
 Field:=1, _
 Criteria1:=Array("1851", "><"), _
 Operator:=xlFilterValues, _
 SubField:="Date founded"
```

<br/>

This example filters a Table, "Table1", on Sheet1 to display the Top 10 entries for field one based off the "Population" SubField.

```vb
Worksheets("Sheet1").ListObjects("Table1").Range.AutoFilter _
 Field:=1, _
 Criteria1:="10", _
 Operator:=xlTop10Items, _
 SubField:="Population"
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
