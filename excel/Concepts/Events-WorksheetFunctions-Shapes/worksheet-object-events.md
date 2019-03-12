---
title: Worksheet object events
keywords: vbaxl10.chm5206017
f1_keywords:
- vbaxl10.chm5206017
ms.prod: excel
ms.assetid: 512e329c-92f6-a8e0-8564-b3ba57e8c296
ms.date: 11/13/2018
localization_priority: Normal
---


# Worksheet object events

Events on sheets are enabled by default. To view the event procedures for a sheet, right-click the sheet tab and click **View Code** on the shortcut menu. Select one of the following events from the **Procedure** list box.

- [Activate](../../../api/Excel.Worksheet.Activate(even).md) 
- [BeforeDoubleClick](../../../api/Excel.Worksheet.BeforeDoubleClick.md) 
- [BeforeRightClick](../../../api/Excel.Worksheet.BeforeRightClick.md) 
- [Calculate](../../../api/Excel.Worksheet.Calculate(even).md) 
- [Change](../../../api/Excel.Worksheet.Change.md) 
- [Deactivate](../../../api/Excel.Worksheet.Deactivate.md) 
- [FollowHyperlink](../../../api/Excel.Worksheet.FollowHyperlink.md) 
- [PivotTableUpdate](../../../api/Excel.Worksheet.PivotTableUpdate.md) 
- [SelectionChange](../../../api/Excel.Worksheet.SelectionChange.md)

Worksheet-level events occur when a worksheet is activated, when the user changes a worksheet cell, or when the PivotTable changes. The following example adjusts the size of columns A through F whenever the worksheet is recalculated.

```vb
Private Sub Worksheet_Calculate() 
    Columns("A:F").AutoFit 
End Sub
```

<br/>

Some events can be used to substitute an action for the default application behavior, or to make a small change to the default behavior. The following example traps the right-click event and adds a new menu item to the shortcut menu for cells B1:B10.

```vb
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, _ 
        Cancel As Boolean) 
    For Each icbc In Application.CommandBars("cell").Controls 
        If icbc.Tag = "brccm" Then icbc.Delete 
    Next icbc 
    If Not Application.Intersect(Target, Range("b1:b10")) _ 
            Is Nothing Then 
        With Application.CommandBars("cell").Controls _ 
            .Add(Type:=msoControlButton, before:=6, _ 
                temporary:=True) 
           .Caption = "New Context Menu Item" 
           .OnAction = "MyMacro" 
           .Tag = "brccm" 
        End With 
    End If 
End Sub
```

## See also

- [Excel functions (by category)](https://support.office.com/article/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
