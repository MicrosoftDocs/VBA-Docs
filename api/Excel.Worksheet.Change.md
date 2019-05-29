---
title: Worksheet.Change event (Excel)
keywords: vbaxl10.chm502079
f1_keywords:
- vbaxl10.chm502079
ms.prod: excel
api_name:
- Excel.Worksheet.Change
ms.assetid: d9e11d08-41ba-f0a8-dc55-6c6cd4e76dd0
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.Change event (Excel)

Occurs when cells on the worksheet are changed by the user or by an external link.


## Syntax

_expression_.**Change** (_Target_)

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Target_|Required| **[Range](Excel.Range(object).md)**|The changed range. Can be more than one cell.|


## Return value

**Nothing**


## Remarks

This event does not occur when cells change during a recalculation. Use the **[Calculate](excel.worksheet.calculate(even).md)** event to trap a sheet recalculation.


## Example

The following code example changes the color of changed cells to blue.

```vb
Private Sub Worksheet_Change(ByVal Target as Range) 
    Target.Font.ColorIndex = 5 
End Sub
```

<br/>

The following code example verifies that, when a cell value changes, the changed cell is in column A, and if the changed value of the cell is greater than 100. If the value is greater than 100, the adjacent cell in column B is changed to the color red.

```vb
Private Sub Worksheet_Change(ByVal Target As Excel.Range) 
    If Target.Column = 1 Then 
        ThisRow = Target.Row 
        If Target.Value > 100 Then 
            Range("B" & ThisRow).Interior.ColorIndex = 3 
        Else 
            Range("B" & ThisRow).Interior.ColorIndex = xlColorIndexNone 
        End If 
    End If 
End Sub
```

<br/>


The following code example sets the values in the range A1:A10 to be uppercase as the data is entered into the cell.

```vb
Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Range("A1:A10")) Is Nothing Or Target.Cells.Count > 1 Then Exit Sub
    Application.EnableEvents = False
    'Set the values to be uppercase
    Target.Value = UCase(Target.Value)
    Application.EnableEvents = True
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
