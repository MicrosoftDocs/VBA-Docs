---
title: Range.ResetContents method (Excel)
api_name:
- Excel.Range.ResetContents
ms.assetid: 375545fc-688f-47db-9185-65e9a60ed9de
ms.date: 07/18/2024
ms.localizationpriority: medium
---


# Range.ResetContents method (Excel)

Clears the cell values of the cells in the range, with special consideration given to cell controls. 

## Remarks

If the range contains only blanks and cells with cell controls set to their defaults (e.g. checkboxes with a cell value of FALSE), then the values and cell control formatting are removed. Otherwise, the cells with cell controls are set to their defaults (e.g. checkboxes are set to cell value of FALSE), and all other cells have their values cleared. 

## Syntax

_expression_.**ResetControls**

-_or_-

_expression_.**RemoveControls**


_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Parameters

Optionally, you can use method **RemoveControls** to clear the cell values of the cells in the range. If ommitted, it gets set to **False**. When set to **True**, it clears all cell control formatting from the range selection, in addition to clearing the cell values. 

## Examples

This example removes contents from the range A1:C4 on Sheet1.

```vb
Range("A1:C4").ResetContents
Selection.ResetContents 
```

This example clears the values from the cells in the range A1:C4 on Sheet1. 

```vb
Range("A1:C4").ClearContents 
Selection.ClearContents 
Selection.ClearContents RemoveControls := True 
Selection.ClearContents RemoveControls := False 
```

This example clears all cell control formatting from the given range, leaving cell values in place. 

```vb
Range("A1:C4").RemoveControls
Selection.RemoveControls 
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
