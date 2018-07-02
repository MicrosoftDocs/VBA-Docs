---
title: Cell.Select Method (Publisher)
keywords: vbapb10.chm5111843
f1_keywords:
- vbapb10.chm5111843
ms.prod: publisher
api_name:
- Publisher.Cell.Select
ms.assetid: 37a026a3-67ec-7a13-5eb4-66e14918579d
ms.date: 06/08/2017
---


# Cell.Select Method (Publisher)

Selects the specified object.


## Syntax

 _expression_. **Select**

 _expression_ A variable that represents a  **Cell** object.


## Example

This example selects the upper-left cell from a table that has been added to the first page in the active publication.


```vb
Dim shpTable As Shape 
Dim cllTemp As Cell 
 
With ActiveDocument.Pages(1).Shapes 
 Set shpTable = .AddTable(NumRows:=3, NumColumns:=3, _ 
 Left:=100, Top:=100, Width:=150, Height:=150) 
 Set cllTemp = shpTable.Table.Cells.Item(1) 
 cllTemp.Select 
End With
```

This example selects the first column from a table that has been added to the first page in the active publication.




```vb
Dim shpTable As Shape 
Dim crTemp As CellRange 
 
With ActiveDocument.Pages(1).Shapes 
 Set shpTable = .AddTable(NumRows:=3, NumColumns:=3, _ 
 Left:=100, Top:=100, Width:=150, Height:=150) 
 Set crTemp = shpTable.Table.Cells(StartRow:=1, _ 
 StartColumn:=1, EndRow:=3, EndColumn:=1) 
 crTemp.Select 
End With
```

This example selects the first five characters in shape one on page one of the active publication.




```vb
ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.Characters(1, 5).Select
```


