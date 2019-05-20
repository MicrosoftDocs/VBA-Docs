---
title: Window.Selection property (Excel)
keywords: vbaxl10.chm356109
f1_keywords:
- vbaxl10.chm356109
ms.prod: excel
api_name:
- Excel.Window.Selection
ms.assetid: 852ca473-28c6-6315-f793-1a12e7f239a4
ms.date: 05/21/2019
localization_priority: Normal
---


# Window.Selection property (Excel)

Returns the specified window, for a **[Windows](Excel.Windows.md)** object.


## Syntax

_expression_.**Selection**

_expression_ A variable that represents a **[Window](Excel.Window.md)** object.


## Remarks

The returned object type depends on the current selection (for example, if a cell is selected, this property returns a **[Range](Excel.Range(object).md)** object). The **Selection** property returns **Nothing** if nothing is selected.

Using this property with no object qualifier is equivalent to using **[Application.Selection](excel.application.selection.md)**.


## Example

This example clears the selection on Sheet1 (assuming that the selection is a range of cells).

```vb
Worksheets("Sheet1").Activate 
Selection.Clear
```

<br/>

This example displays the Visual Basic object type of the selection.

```vb
Worksheets("Sheet1").Activate 
MsgBox "The selection object type is " & TypeName(Selection)
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]