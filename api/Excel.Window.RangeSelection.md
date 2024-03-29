---
title: Window.RangeSelection property (Excel)
keywords: vbaxl10.chm356104
f1_keywords:
- vbaxl10.chm356104
api_name:
- Excel.Window.RangeSelection
ms.assetid: 1290970f-4a7a-ce68-da5a-d1a90dacf19f
ms.date: 05/21/2019
ms.localizationpriority: medium
---


# Window.RangeSelection property (Excel)

Returns a **[Range](Excel.Range(object).md)** object that represents the selected cells on the worksheet in the specified window even if a graphic object is active or selected on the worksheet. Read-only.


## Syntax

_expression_.**RangeSelection**

_expression_ A variable that represents a **[Window](Excel.Window.md)** object.


## Remarks

When a graphic object is selected on a worksheet, the **Selection** property returns the graphic object instead of a **Range** object; the **RangeSelection** property returns the range of cells that was selected before the graphic object was selected.

This property and the **Selection** property return identical values when a range (not a graphic object) is selected on the worksheet.

If the active sheet in the specified window isn't a worksheet, this property fails.


## Example

This example displays the address of the selected cells on the worksheet in the active window.

```vb
MsgBox ActiveWindow.RangeSelection.Address
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]