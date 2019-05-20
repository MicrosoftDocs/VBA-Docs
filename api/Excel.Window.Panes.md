---
title: Window.Panes property (Excel)
keywords: vbaxl10.chm356101
f1_keywords:
- vbaxl10.chm356101
ms.prod: excel
api_name:
- Excel.Window.Panes
ms.assetid: ba89f562-66f8-990d-e034-c996557b3687
ms.date: 05/21/2019
localization_priority: Normal
---


# Window.Panes property (Excel)

Returns a **[Panes](Excel.Panes.md)** collection that represents all the panes in the specified window. Read-only.


## Syntax

_expression_.**Panes**

_expression_ A variable that represents a **[Window](Excel.Window.md)** object.


## Remarks

This property is available for a window only if the window's **Split** property can be set to **True**.


## Example

This example displays the number of panes in the active window in Book1.xls.

```vb
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
MsgBox "There are " & ActiveWindow.Panes.Count & _ 
 " panes in the active window"
```

<br/>

This example activates the pane in the upper-left corner of the active window in Book1.xls.

```vb
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
ActiveWindow.Panes(1).Activate
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]