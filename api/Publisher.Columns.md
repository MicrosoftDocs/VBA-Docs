---
title: Columns object (Publisher)
keywords: vbapb10.chm5111807
f1_keywords:
- vbapb10.chm5111807
ms.prod: publisher
api_name:
- Publisher.Columns
ms.assetid: 3fe6ddce-a598-a967-fc89-7296c18a6a55
ms.date: 05/31/2019
localization_priority: Normal
---


# Columns object (Publisher)

A collection of **[Column](Publisher.Column.md)** objects that represent the columns in a table.
 
## Remarks

Use the **[Columns](Publisher.Table.Columns.md)** property of the **Table** object to return the **Columns** collection.

Use **Columns** (_index_), where _index_ is the index number, to return a single **Column** object. The index number represents the position of the column in the **Columns** collection (counting from left to right). 

Use the **Add** method to add a column to a table. 

## Example

The following example displays the number of **Column** objects in the **Columns** collection for the first table in the active document.

```vb
Sub CountColumns() 
 MsgBox "The number of columns in the table is " & _ 
 ActiveDocument.Pages(2).Shapes(1).Table.Columns.Count 
End Sub
```

<br/>

This example enters a bold number into each cell in the specified table. It assumes that the specified shape is a table and not another type of shape.

```vb
Sub CountCellsByColumn() 
 Dim shpTable As Shape 
 Dim colTable As Column 
 Dim celTable As Cell 
 Dim intCount As Integer 
 
 intCount = 1 
 
 Set shpTable = ActiveDocument.Pages(2).Shapes(1) 
 For Each colTable In shpTable.Table.Columns 
 For Each celTable In colTable.Cells 
 With celTable.Text 
 .Text = intCount 
 .ParagraphFormat.Alignment = _ 
 pbParagraphAlignmentCenter 
 .Font.Bold = msoTrue 
 intCount = intCount + 1 
 End With 
 Next celTable 
 Next colTable 
 
End Sub
```

<br/>

The following example selects the third column in the specified table.

```vb
Sub SelectColumns() 
 ActiveDocument.Pages(2).Shapes(1).Table.Columns(3).Cells.Select 
End Sub
```

<br/>

This example adds a column to the specified table on the second page of the active publication, and then adjusts the width, merges the cells, and sets the fill color. This example assumes that the first shape is a table and not another type of shape.

```vb
Sub NewColumn() 
 Dim colNew As Column 
 
 Set colNew = ActiveDocument.Pages(2).Shapes(1).Table.Columns _ 
 .Add(BeforeColumn:=3) 
 With colNew 
 .Width = 2 
 .Cells.Merge 
 .Cells(1).Fill.ForeColor.RGB = RGB(Red:=202, Green:=202, Blue:=202) 
 End With 
End Sub
```


## Methods

- [Add](Publisher.Columns.Add.md)
- [Item](Publisher.Columns.Item.md)

## Properties

- [Application](Publisher.Columns.Application.md)
- [Count](Publisher.Columns.Count.md)
- [Parent](Publisher.Columns.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]