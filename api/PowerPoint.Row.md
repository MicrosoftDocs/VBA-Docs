---
title: Row object (PowerPoint)
keywords: vbapp10.chm626000
f1_keywords:
- vbapp10.chm626000
ms.prod: powerpoint
api_name:
- PowerPoint.Row
ms.assetid: df5ca5df-8119-1af8-b698-d96669ed0a02
ms.date: 06/08/2017
localization_priority: Normal
---


# Row object (PowerPoint)

Represents a row in a table. The **Row** object is a member of the **[Rows](PowerPoint.Rows.md)** collection. The **Rows** collection includes all the rows in the specified table.


## Example

Use  **Rows** (_index_), where _index_ is a number that represents the position of the row in the table, to return a single **Row** object. This example deletes the first row from the table in shape five on slide two of the active presentation.


```vb
ActivePresentation.Slides(2).Shapes(5).Table.Rows(1).Delete
```

Use the [Select](PowerPoint.Row.Select.md)method to select a row in a table. This example selects row one of the specified table.




```vb
ActivePresentation.Slides(2).Shapes(5).Table.Rows(1).Select
```

Use the [Cells](PowerPoint.Row.Cells.md)property to modify the individual cells in a **Row** object. This example selects the second row in the table and applies a dashed line style to the bottom border.




```vb
ActiveWindow.Selection.ShapeRange.Table.Rows(2) _
    .Cells.Borders(ppBorderBottom).DashStyle = msoLineDash
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]