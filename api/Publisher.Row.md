---
title: Row object (Publisher)
keywords: vbapb10.chm4915199
f1_keywords:
- vbapb10.chm4915199
ms.prod: publisher
api_name:
- Publisher.Row
ms.assetid: 11f4688b-b94e-fa09-7c1b-4cbcca330936
ms.date: 06/01/2019
localization_priority: Normal
---


# Row object (Publisher)

Represents a row in a table. The **Row** object is a member of the **[Rows](Publisher.Rows.md)** collection. The **Rows** collection includes all the rows in a specified table.
 
## Remarks

Use **Rows** (_index_), where _index_ is the row number, to return a single **Row** object. The index number represents the position of the row in the **Rows** collection (counting from left to right). 

Use the **[Item](Publisher.Rows.Item.md)** method of a **Rows** collection to return a **Row** object. 

Use the **[Add](Publisher.Rows.Add.md)** method to add a row to a table. 

Use the **Delete** method to delete a row from a table. 


## Example

This example selects the first row in the first shape on the second page of the active publication. This example assumes that the specified shape is a table and not another type of shape.

```vb
Sub SelectRow() 
 ActiveDocument.Pages(2).Shapes(1).Table.Rows(1).Cells.Select 
End Sub
```

<br/>

This example sets the fill for all even-numbered rows, and clears the fill for all odd-numbered rows in the specified table. This example assumes that the specified shape is a table and not another type of shape.

```vb
Sub FillCellsByRow() 
 Dim shpTable As Shape 
 Dim rowTable As Row 
 Dim celTable As Cell 
 
 Set shpTable = ActiveDocument.Pages(2).Shapes(1) 
 For Each rowTable In shpTable.Table.Rows 
 For Each celTable In rowTable.Cells 
 If celTable.Row Mod 2 = 0 Then 
 celTable.Fill.ForeColor.RGB = RGB _ 
 (Red:=180, Green:=180, Blue:=180) 
 Else 
 celTable.Fill.ForeColor.RGB = RGB _ 
 (Red:=255, Green:=255, Blue:=255) 
 End If 
 Next celTable 
 Next rowTable 
 
End Sub
```

<br/>

This example adds a row to the specified table on the second page of the active publication, and then adjusts the width, merges the cells, and sets the fill color. This example assumes that the first shape is a table and not another type of shape.

```vb
Sub NewRow() 
 Dim rowNew As Row 
 
 Set rowNew = ActiveDocument.Pages(2).Shapes(1).Table.Rows _ 
 .Add(BeforeRow:=3) 
 With rowNew 
 .Height = 2 
 .Cells.Merge 
 .Cells(1).Fill.ForeColor.RGB = RGB(Red:=0, Green:=0, Blue:=0) 
 End With 
End Sub
```

<br/>

This example deletes the row added in the previous example.

```vb
Sub DeleteRow() 
 ActiveDocument.Pages(2).Shapes(1).Table.Rows(3).Delete 
End Sub
```


## Methods

- [Delete](Publisher.Row.Delete.md)

## Properties

- [Application](Publisher.Row.Application.md)
- [Cells](Publisher.Row.Cells.md)
- [Height](Publisher.Row.Height.md)
- [Parent](Publisher.Row.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]