---
title: Rows object (PowerPoint)
keywords: vbapp10.chm625000
f1_keywords:
- vbapp10.chm625000
ms.prod: powerpoint
api_name:
- PowerPoint.Rows
ms.assetid: 9a72b6bb-2aec-e37b-f1a2-005f910e1335
ms.date: 06/08/2017
localization_priority: Normal
---


# Rows object (PowerPoint)

A collection of  **[Row](PowerPoint.Row.md)** objects that represent the rows in a table.


## Example

Use the [Rows](PowerPoint.Table.Rows.md)property to return the  **Rows** collection. This example changes the height of all rows in the specified table to 160 points.


```vb
Dim i As Integer

With ActivePresentation.Slides(2).Shapes(4).Table

    For i = 1 To .Rows.Count

        .Rows.Height = 160

    Next i

End With
```

Use the [Add](PowerPoint.Rows.Add.md)method to add a row to a table. This example inserts a row before the second row in the referenced table.




```vb
ActivePresentation.Slides(2).Shapes(5).Table.Rows.Add (2)
```

Use  **Rows** (_index_), where _index_ is a number that represents the position of the row in the table, to return a single **Row** object. This example deletes the first row from the table in shape five on slide two.




```vb
ActivePresentation.Slides(2).Shapes(5).Table.Rows(1).Delete
```


## Methods



|Name|
|:-----|
|[Add](PowerPoint.Rows.Add.md)|
|[Item](PowerPoint.Rows.Item.md)|

## Properties



|Name|
|:-----|
|[Application](PowerPoint.Rows.Application.md)|
|[Count](PowerPoint.Rows.Count.md)|
|[Parent](PowerPoint.Rows.Parent.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]