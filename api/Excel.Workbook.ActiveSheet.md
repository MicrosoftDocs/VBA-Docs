---
title: Workbook.ActiveSheet Property (Excel)
keywords: vbaxl10.chm199076
f1_keywords:
- vbaxl10.chm199076
ms.prod: excel
api_name:
- Excel.Workbook.ActiveSheet
ms.assetid: fb5578c3-64a7-edb7-4004-e608739d4c1e
ms.date: 06/08/2017
---


# Workbook.ActiveSheet Property (Excel)

<<<<<<< HEAD
Returns an object that represents the active sheet (the sheet on top) in the active workbook or in the specified window or workbook. Returns  **Nothing** if no sheet is active.
=======
Returns a **[Worksheet](excel.worksheet.md)** object that represents the active sheet (the sheet on top) in the active workbook or specified workbook. Returns **Nothing** if no sheet is active.
>>>>>>> master


## Syntax

 _expression_. `ActiveSheet`

 _expression_ A variable that represents a [Workbook](./Excel.Workbook.md) object.


## Remarks

<<<<<<< HEAD
If you don't specify an object qualifier, this property returns the active sheet in the active workbook.

If a workbook appears in more than one window, the  **ActiveSheet** property may be different in different windows.

=======
Using the `ActiveSheet` property without an object qualifier returns the active sheet in the active workbook in the active window.

If a workbook appears in more than one window, the active sheet might be different in different windows.
>>>>>>> master

## Example

This example displays the name of the active sheet.


```vb
<<<<<<< HEAD
MsgBox "The name of the active sheet is " &; ActiveSheet.Name
=======
MsgBox "The name of the active sheet is " & ActiveSheet.Name
>>>>>>> master
```


## See also

<<<<<<< HEAD

[Workbook Object](Excel.Workbook.md)

=======
[Workbook Object](Excel.Workbook.md)

[Window Object](excel.window.md)

[Worksheet.Activate Method](excel.worksheet.activate(method).md)

[Window.ActiveSheet Property](excel.window.activesheet.md)


>>>>>>> master
