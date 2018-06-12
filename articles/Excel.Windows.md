---
title: Windows Object (Excel)
keywords: vbaxl10.chm353072
f1_keywords:
- vbaxl10.chm353072
ms.prod: excel
api_name:
- Excel.Windows
ms.assetid: d5d0e3c9-9132-469c-d033-d29397dacd77
ms.date: 06/08/2017
---


# Windows Object (Excel)

A collection of all the  **[Window](Excel.Window.md)** objects in Microsoft Excel.


## Remarks

The  **Windows** collection for the **[Application](Excel.Application(objec).md)** object contains all the windows in the application, whereas the **Windows** collection for the **[Workbook](Excel.Workbook.md)** object contains only the windows in the specified workbook.


## Example

Use the  **Windows** property to return the **Windows** collection. The following example cascades all the windows that are currently displayed in Microsoft Excel.


```
Windows.Arrange arrangeStyle:=xlCascade
```

Use the  **[NewWindow](Excel.Window.NewWindow.md)** method to create a new window and add it to the collection. The following example creates a new window for the active workbook.




```
ActiveWorkbook.NewWindow
```

Use  **Windows** ( _index_ ), where _index_ is the window name or index number, to return a single **Window** object. The following example maximizes the active window.

Note that the active window is always  `Windows(1)`.




```
Windows(1).WindowState = xlMaximized
```


## Methods



|**Name**|
|:-----|
|[Arrange](Excel.Windows.Arrange.md)|
|[BreakSideBySide](Excel.Windows.BreakSideBySide.md)|
|[CompareSideBySideWith](Excel.Windows.CompareSideBySideWith.md)|
|[ResetPositionsSideBySide](Excel.Windows.ResetPositionsSideBySide.md)|

## Properties



|**Name**|
|:-----|
|[Application](Excel.Windows.Application.md)|
|[Count](Excel.Windows.Count.md)|
|[Creator](Excel.Windows.Creator.md)|
|[Item](Excel.Windows.Item.md)|
|[Parent](Excel.Windows.Parent.md)|
|[SyncScrollingSideBySide](Excel.Windows.SyncScrollingSideBySide.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
