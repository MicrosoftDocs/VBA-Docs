---
title: Validation.Value property (Excel)
keywords: vbaxl10.chm532090
f1_keywords:
- vbaxl10.chm532090
api_name:
- Excel.Validation.Value
ms.assetid: 8c1e3946-ea57-4aa7-5f1d-be9e6a2c8f77
ms.date: 10/06/2026
ms.localizationpriority: medium
---


# Validation.Value property (Excel)

Returns a **Boolean** value that indicates whether the value of the cell meets the data validation criteria. For a multiple-cell range, only the upper-left cell of the range is evaluated. Read-only.


## Syntax

_expression_.**Value**

_expression_ A variable that represents a **[Validation](Excel.Validation.md)** object.


## Remarks

When the **Validation** object is returned from a range of more than one cell, the **Value** property reports the result for the upper-left cell of the range only. It returns **True** if the upper-left cell is valid, even if other cells in the range contain invalid data. This is also the case when the cells in the range have different validation rules. To check every cell in a range, read the property for each cell individually, as shown in the following example.

The **Value** property returns **True** for a cell that doesn't have data validation.

For a blank cell, the result depends on the **[IgnoreBlank](Excel.Validation.IgnoreBlank.md)** property. If **IgnoreBlank** is **True** (the default), a blank cell is reported as valid. If it's **False**, a blank cell is checked against the validation criteria like any other value. For example, a whole number rule that allows 0 to 100 treats a blank cell as 0 and reports it as valid, while a rule that allows 1 to 100 reports it as invalid. A list rule reports a blank cell as invalid.

Data validation is evaluated when a user enters a value in the cell. Values that are assigned by code (for example, through the **[Range.Value](Excel.Range.Value.md)** property), pasted as values only (for example, by using **Paste Special** > **Values**), or pasted as text from another application aren't validated when they're written, so a cell can contain a value that doesn't meet its own validation criteria. You can use the **Value** property to find such cells afterward. Note that a regular paste from another cell replaces the data validation of the destination cell with that of the source cell, and removes it if the source cell has none.


## Example

This example checks each cell in the range A1:A20 on Sheet1 and reports the cells whose values don't meet the data validation criteria.

```vb
Sub ListInvalidCells()
    Dim cell As Range

    For Each cell In Worksheets("Sheet1").Range("A1:A20").Cells
        If Not cell.Validation.Value Then
            Debug.Print cell.Address & " is invalid: " & cell.Formula
        End If
    Next cell
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]