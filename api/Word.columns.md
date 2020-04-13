---
title: Columns object (Word)
ms.prod: word
ms.assetid: 7c2d1353-cbc4-a162-83a1-6cac1300266f
ms.date: 06/08/2017
localization_priority: Normal
---


# Columns object (Word)

A collection of  **[Column](Word.Column.md)** objects that represent the columns in a table.


## Remarks

Use the **Columns** property of a **[Range](Word.Range.md)**, **[Selection](Word.Selection.md)**, or **[Table](Word.Table.md)** object to return a **Columns** collection. The following example displays the number of **Column** objects in the **Columns** collection for the first table in the active document.


```vb
MsgBox ActiveDocument.Tables(1).Columns.Count
```

The following example creates a table with six columns and three rows and then formats each column with a progressively larger (darker) shading percentage.




```vb
Set myTable = ActiveDocument.Tables.Add(Range:=Selection.Range, _ 
 NumRows:=3, NumColumns:=6) 
For Each col In myTable.Columns 
 col.Shading.Texture = 2 + i 
 i = i + 1 
Next col
```

Use the **[Add](Word.Columns.Add.md)** method to add a column to a table. The following example adds a column to the first table in the active document, and then it makes the column widths equal.




```vb
If ActiveDocument.Tables.Count >= 1 Then 
 Set myTable = ActiveDocument.Tables(1) 
 myTable.Columns.Add BeforeColumn:=myTable.Columns(1) 
 myTable.Columns.DistributeWidth 
End If
```

Use  **Columns** (Index), where Index is the index number, to return a single **Column** object. The index number represents the position of the column in the **Columns** collection (counting from left to right). The following example selects the first column in the first table.




```vb
ActiveDocument.Tables(1).Columns(1).Select
```


## Methods



|Name|
|:-----|
|[Add](Word.Columns.Add.md)|
|[AutoFit](Word.Columns.AutoFit.md)|
|[Delete](Word.Columns.Delete.md)|
|[DistributeWidth](Word.Columns.DistributeWidth.md)|
|[Item](Word.Columns.Item.md)|
|[Select](Word.Columns.Select.md)|
|[SetWidth](Word.Columns.SetWidth.md)|

## Properties



|Name|
|:-----|
|[Application](Word.Columns.Application.md)|
|[Borders](Word.Columns.Borders.md)|
|[Count](Word.Columns.Count.md)|
|[Creator](Word.Columns.Creator.md)|
|[First](Word.Columns.First.md)|
|[Last](Word.Columns.Last.md)|
|[NestingLevel](Word.Columns.NestingLevel.md)|
|[Parent](Word.Columns.Parent.md)|
|[PreferredWidth](Word.Columns.PreferredWidth.md)|
|[PreferredWidthType](Word.Columns.PreferredWidthType.md)|
|[Shading](Word.Columns.Shading.md)|
|[Width](Word.Columns.Width.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]