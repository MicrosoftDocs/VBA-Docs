---
title: Range.AutoFilter method (Excel)
keywords: vbaxl10.chm144084
f1_keywords:
- vbaxl10.chm144084
ms.prod: excel
api_name:
- Excel.Range.AutoFilter
ms.date: 05/10/2019
localization_priority: Priority
---

# Range.AutoFilter method (Excel)

Filters a list by using the AutoFilter.

## Syntax

_expression_.**AutoFilter** (_Field_, _Criteria1_, _Operator_, _Criteria2_, _SubField_, _VisibleDropDown_)

_expression_ An expression that returns a **[Range](Excel.Range(object).md)** object.


## Parameters

|Name |Required/Optional |Data type |Description|
|:-----|:-----|:-----|:-----|
| _Field_|Optional| **Variant**| The integer offset of the field on which you want to base the filter (from the left of the list; the leftmost field is field one).|
| _Criteria1_|Optional| **Variant**|The criteria (a string; for example, "101"). Use `"="` to find blank fields, `"<>"` to find non-blank fields, and `"><"` to select (No Data) fields in data types.<br/><br/>If this argument is omitted, the criteria is All. If _Operator_ is **xlTop10Items**, _Criteria1_ specifies the number of items (for example, "10").|
| _Operator_|Optional| **[XlAutoFilterOperator](Excel.XlAutoFilterOperator.md)**|An **XlAutoFilterOperator** constant specifying the type of filter.|
| _Criteria2_|Optional| **Variant**|The second criteria (a string). Used with _Criteria1_ and _Operator_ to construct compound criteria. Also used as single criteria on date fields filtering by date, month or year. Followed by an Array detailing the filtering **Array(Level, Date)**. Where Level is 0-2 (year,month,date) and Date is one valid Date inside the filtering period.|
| _SubField_|Optional| **Variant**|The field from a data type on which to apply the criteria (for example, the "Population" field from Geography or "Volume" field from Stocks). Omitting this value targets the "(Display Value)".|
| _VisibleDropDown_|Optional| **Variant**| **True** to display the AutoFilter drop-down arrow for the filtered field. **False** to hide the AutoFilter drop-down arrow for the filtered field. **True** by default.|

## Return value

Variant

## Remarks

If you omit all the arguments, this method simply toggles the display of the AutoFilter drop-down arrows in the specified range.

Excel for Mac does not support this method. Similar methods on **Selection** and **ListObject** are supported.

Unlike in formulas, subfields do not require brackets to include spaces.


## Example

This example filters a list starting in cell A1 on Sheet1 to display only the entries in which field one is equal to the string Otis. The drop-down arrow for field one will be hidden.

```vb
Worksheets("Sheet1").Range("A1").AutoFilter _
 Field:=1, _
 Criteria1:="Otis", _
 VisibleDropDown:=False
```

<br/>

This example filters a list starting in cell A1 on Sheet1 to display only the entries in which the values of field one contain a SubField, Admin Division 1 (State/province/other), where the value is Washington.

```vb
Worksheets("Sheet1").Range("A1").AutoFilter _
 Field:=1, _
 Criteria1:="Washington", _
 SubField:="Admin Division 1 (State/province/other)"
```

<br/>

This example filters a table, Table1, on Sheet1 to display only the entries in which the values of field one have a "(Display Value)" that is either 1, 3, Seattle, or Redmond.

```vb
Worksheets("Sheet1").ListObjects("Table1").Range.AutoFilter _
 Field:=1, _
 Criteria1:=Array("1", "3", "Seattle", "Redmond"), _
 Operator:=xlFilterValues
```

<br/>

Data types can apply multiple SubField filters. This example filters a table, Table1, on Sheet1 to display only the entries in which the values of field one contain a SubField, Time Zone(s), where the value is Pacific Time Zone, and where the SubField named Date Founded is either 1851 or there is "(No Data)".

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

This example filters a table, Table1, on Sheet1 to display the Top 10 entries for field one based off the Population SubField.

```vb
Worksheets("Sheet1").ListObjects("Table1").Range.AutoFilter _
 Field:=1, _
 Criteria1:="10", _
 Operator:=xlTop10Items, _
 SubField:="Population"
```

This example filters a table, Table1, on Sheet1 to display the all entries for January 2019 and February 2019 for field one. There does not have to be a row containing January the 31.

```vb
Worksheets("Sheet1").ListObjects("Table1").Range.AutoFilter _
 Field:=1, _
 Criteria2:=Array(1, "1/31/2019", 1, "2/28/2019") 
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
