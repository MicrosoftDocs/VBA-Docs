---
title: Table object (Publisher)
keywords: vbapb10.chm4849663
f1_keywords:
- vbapb10.chm4849663
ms.prod: publisher
api_name:
- Publisher.Table
ms.assetid: 09da4a0a-2230-067e-1cac-55321ea044c5
ms.date: 06/01/2019
localization_priority: Normal
---


# Table object (Publisher)

Represents a single table.

## Remarks

Use the **[Shape.Table](Publisher.Shape.Table.md)** property to return a **Table** object. 

Use the **[Shapes.AddTable](Publisher.Shapes.AddTable.md)** method to add a **Shape** object representing a table at the specified range. 


## Example

The following example selects the specified table in the active publication.

```vb
Sub SelectTable() 
 With ActiveDocument.Pages(1).Shapes(1) 
 If .Type = pbTable Then _ 
 .Table.Cells.Select 
 End With 
End Sub
```

<br/>

The following example adds a 5x5 table on the first page of the active publication, and then selects the first column of the new table.

```vb
Sub NewTable() 
 With ActiveDocument.Pages(1).Shapes.AddTable(NumRows:=5, NumColumns:=5, _ 
 Left:=72, Top:=300, Width:=400, Height:=100) 
 .Table.Columns(1).Cells.Select 
 End With 
End Sub
```


## Methods

- [ApplyAutoFormat](Publisher.Table.ApplyAutoFormat.md)

## Properties

- [Application](Publisher.Table.Application.md)
- [Cells](Publisher.Table.Cells.md)
- [Columns](Publisher.Table.Columns.md)
- [GrowToFitText](Publisher.Table.GrowToFitText.md)
- [Parent](Publisher.Table.Parent.md)
- [Rows](Publisher.Table.Rows.md)
- [TableDirection](Publisher.Table.TableDirection.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]