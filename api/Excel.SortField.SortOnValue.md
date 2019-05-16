---
title: SortField.SortOnValue property (Excel)
keywords: vbaxl10.chm843074
f1_keywords:
- vbaxl10.chm843074
ms.prod: excel
api_name:
- Excel.SortField.SortOnValue
ms.assetid: eeaaf959-71d2-99a3-7e66-61744ad4709e
ms.date: 05/16/2019
localization_priority: Normal
---


# SortField.SortOnValue property (Excel)

Returns the value on which the sort is performed for the specified **SortField** object. Read-only.


## Syntax

_expression_.**SortOnValue**

_expression_ A variable that represents a **[SortField](Excel.SortField.md)** object.


## Example

This example sorts the data in column B on Sheet1 by font color in ascending order.

```vb
ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear 
ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add(Range("B1:B25"), _ 
 xlSortOnFontColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(0, 0, 0) 
 
With ActiveWorkbook.Worksheets("Sheet1").Sort 
 .SetRange Range("A1:B25") 
 .Header = xlGuess 
 .MatchCase = False 
 .Orientation = xlTopToBottom 
 .SortMethod = xlPinYin 
 .Apply 
End With
```

<br/>

This example sorts the data by cell color.

```vb
SortOn = xlSortOnCellColor 
SortOnValue.Color = RGB(255, 255, 0)
```

<br/>

This example sorts the data by font color.

```vb
SortOn = xlSortOnFontColor 
SortOnValue.Color = RGB(255, 255, 0)
```

<br/>

This example sorts the data by icons.

```vb
SortOn = xlSortOnIcon 
SortOnValue.Color = RGB(255, 255, 0) 
SortField.SetIcon ActiveWorkbook.IconSets(1).Item(3)
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]