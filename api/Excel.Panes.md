---
title: Panes object (Excel)
keywords: vbaxl10.chm357072
f1_keywords:
- vbaxl10.chm357072
ms.prod: excel
api_name:
- Excel.Panes
ms.assetid: ce27ae27-52d9-9e51-a068-b9c082a0a692
ms.date: 03/30/2019
localization_priority: Normal
---


# Panes object (Excel)

A collection of all the **[Pane](Excel.Pane.md)** objects shown in the specified window.


## Remarks

**Pane** objects exist only for worksheets and Microsoft Excel 4.0 macro sheets.


## Example

Use the **[Panes](excel.window.panes.md)** property of the **Windows** object to return the **Panes** collection. 

The following example freezes panes in the active window if the window contains more than one pane.

```vb
If ActiveWindow.Panes.Count > 1 Then _ 
 ActiveWindow.FreezePanes = True
```

<br/>

Use **Panes** (_index_), where _index_ is the pane index number, to return a single **Pane** object. The following example scrolls through the upper-left pane of the window in which Sheet1 is displayed.

```vb
Worksheets("sheet1").Activate 
Windows(1).Panes(1).LargeScroll down:=1
```

## Properties

- [Application](Excel.Panes.Application.md)
- [Count](Excel.Panes.Count.md)
- [Creator](Excel.Panes.Creator.md)
- [Item](Excel.Panes.Item.md)
- [Parent](Excel.Panes.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]