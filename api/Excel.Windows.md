---
title: Windows object (Excel)
keywords: vbaxl10.chm353072
f1_keywords:
- vbaxl10.chm353072
ms.prod: excel
api_name:
- Excel.Windows
ms.assetid: d5d0e3c9-9132-469c-d033-d29397dacd77
ms.date: 04/03/2019
localization_priority: Normal
---


# Windows object (Excel)

A collection of all the **[Window](Excel.Window.md)** objects in Microsoft Excel.


## Remarks

The **Windows** collection for the **[Application](Excel.Application(object).md)** object contains all the windows in the application, whereas the **Windows** collection for the **[Workbook](Excel.Workbook.md)** object contains only the windows in the specified workbook.


## Example

Use the **[Windows](excel.application.windows.md)** property to return the **Windows** collection. The following example cascades all the windows that are currently displayed in Microsoft Excel.

```vb
Windows.Arrange arrangeStyle:=xlCascade
```

<br/>

Use the **[NewWindow](Excel.Window.NewWindow.md)** method to create a new window and add it to the collection. The following example creates a new window for the active workbook.

```vb
ActiveWorkbook.NewWindow
```

<br/>

Use **Windows** (_index_), where _index_ is the window name or index number, to return a single **Window** object. The following example maximizes the active window.

Note that the active window is always `Windows(1)`.

```vb
Windows(1).WindowState = xlMaximized
```


## Methods

- [Arrange](Excel.Windows.Arrange.md)
- [BreakSideBySide](Excel.Windows.BreakSideBySide.md)
- [CompareSideBySideWith](Excel.Windows.CompareSideBySideWith.md)
- [ResetPositionsSideBySide](Excel.Windows.ResetPositionsSideBySide.md)

## Properties

- [Application](Excel.Windows.Application.md)
- [Count](Excel.Windows.Count.md)
- [Creator](Excel.Windows.Creator.md)
- [Item](Excel.Windows.Item.md)
- [Parent](Excel.Windows.Parent.md)
- [SyncScrollingSideBySide](Excel.Windows.SyncScrollingSideBySide.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
