---
title: Application.Rows property (Excel)
keywords: vbaxl10.chm132103
f1_keywords:
- vbaxl10.chm132103
ms.prod: excel
api_name:
- Excel.Application.Rows
ms.assetid: 499f6045-1334-a8f8-9a04-f1aef7908312
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.Rows property (Excel)

Returns a **[Range](Excel.Range(object).md)** object that represents all the rows on the active worksheet. If the active document isn't a worksheet, the **Rows** property fails. Read-only **Range** object.


## Syntax

_expression_.**Rows**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

Using this property without an object qualifier is equivalent to using ActiveSheet.Rows.

When applied to a **Range** object that's a multiple selection, this property returns rows from only the first area of the range. For example, if the **Range** object has two areas—A1:B2 and C3:D4, Selection.Rows.Count returns 2, not 4. 

To use this property on a range that may contain a multiple selection, test Areas.Count to determine whether the range is a multiple selection. If it is, loop over each area in the range, as shown in the third example.


## Example

This example deletes row three on Sheet1.

```vb
Worksheets("Sheet1").Rows(3).Delete
```

<br/>

This example deletes rows in the current region on worksheet one where the value of cell one in the row is the same as the value in cell one in the previous row.

```vb
For Each rw In Worksheets(1).Cells(1, 1).CurrentRegion.Rows 
 this = rw.Cells(1, 1).Value 
 If this = last Then rw.Delete 
 last = this 
Next
```

<br/>

This example displays the number of rows in the selection on Sheet1. If more than one area is selected, the example loops through each area.

```vb
Worksheets("Sheet1").Activate 
areaCount = Selection.Areas.Count 
If areaCount <= 1 Then 
 MsgBox "The selection contains " & _ 
 Selection.Rows.Count & " rows." 
Else 
 i = 1 
 For Each a In Selection.Areas 
 MsgBox "Area " & i & " of the selection contains " & _ 
 a.Rows.Count & " rows." 
 i = i + 1 
 Next a 
End If
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
