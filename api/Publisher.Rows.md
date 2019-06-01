---
title: Rows object (Publisher)
keywords: vbapb10.chm4980735
f1_keywords:
- vbapb10.chm4980735
ms.prod: publisher
api_name:
- Publisher.Rows
ms.assetid: 31b04a41-9005-8f51-87ab-426af0e901ed
ms.date: 06/01/2019
localization_priority: Normal
---


# Rows object (Publisher)

A collection of **[Row](Publisher.Row.md)** objects that represent the rows in a table.
 
## Remarks

Use the **[Rows](Publisher.Table.Rows.md)** property of the **Table** object to return the **Rows** collection. 

Use **Rows** (_index_), where _index_ is the index number, to return a single **Row** object. The index number represents the position of the row in the **Rows** collection (counting from left to right). 

## Example

The following example displays the number of **Row** objects in the **Rows** collection for the first table in the active document.

```vb
Sub CountRows() 
 MsgBox ActiveDocument.Pages(2).Shapes(1).Table.Rows.Count 
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

The following example selects the third row in the specified table.

```vb
Sub SelectRows() 
 ActiveDocument.Pages(2).Shapes(1).Table.Rows(3).Cells.Select 
End Sub
```


## Methods

- [Add](Publisher.Rows.Add.md)
- [Item](Publisher.Rows.Item.md)

## Properties

- [Application](Publisher.Rows.Application.md)
- [Count](Publisher.Rows.Count.md)
- [Parent](Publisher.Rows.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]