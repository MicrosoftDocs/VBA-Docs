---
title: Column object (Publisher)
keywords: vbapb10.chm5046271
f1_keywords:
- vbapb10.chm5046271
ms.prod: publisher
api_name:
- Publisher.Column
ms.assetid: 7f14fd4f-3919-8dd9-ed1e-988269b4b0c9
ms.date: 05/31/2019
localization_priority: Normal
---


# Column object (Publisher)

Represents a single table column. The **Column** object is a member of the **[Columns](Publisher.Columns.md)** collection. The **Columns** collection includes all the columns in a table, selection, or range.

## Remarks

Use **Columns** (_index_), where _index_ is the column number, to return a single **Column** object. The index number represents the position of the column in the **Columns** collection (counting from left to right). 

Use the **[Item](Publisher.Columns.Item.md)** method of a **Columns** collection to return a **Column** object. 

Use the **Delete** method to delete a column from a table. 
 
## Example

This example selects column three in the first shape in the active publication. It assumes that the first shape is a table and not another type of shape.

```vb
Sub SelectColumn() 
 ActiveDocument.Pages(2).Shapes(1).Table.Columns(3).Cells.Select 
End Sub
```

<br/>

This example enters text into the first cell of the third column of the specified table and formats the text with a bold, 15-point font. It assumes that the first shape is a table and not another type of shape.
 
```vb
Sub ColumnHeading() 
 With ActiveDocument.Pages(2).Shapes(1).Table.Columns(3) _ 
 .Cells(1).Text 
 .Text = "Sales" 
 .Font.Bold = msoTrue 
 .Font.Size = 15 
 End With 
End Sub
```

<br/>

This example deletes the column added in the previous example.
 
```vb
Sub DeleteColumn() 
 ActiveDocument.Pages(2).Shapes(1).Table.Columns(3).Delete 
End Sub
```


## Methods

- [Delete](Publisher.Column.Delete.md)

## Properties

- [Application](Publisher.Column.Application.md)
- [Cells](Publisher.Column.Cells.md)
- [Parent](Publisher.Column.Parent.md)
- [Width](Publisher.Column.Width.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]