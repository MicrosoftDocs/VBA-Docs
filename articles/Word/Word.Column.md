---
title: Column Object (Word)
keywords: vbawd10.chm2383
f1_keywords:
- vbawd10.chm2383
ms.prod: word
api_name:
- Word.Column
ms.assetid: 49d68571-2a57-6795-34b9-eb09aeb43043
ms.date: 06/08/2017
---


# Column Object (Word)

Represents a single table column. The  **Column** object is a member of the **[Columns](Word.columns.md)** collection. The **Columns** collection includes all the columns in a table, selection, or range.


## Remarks

Use  **Columns** (Index), where Index is the index number, to return a single **Column** object. The index number represents the position of the column in the **[Columns](Word.columns.md)** collection (counting from left to right).

The following example selects column one in table one in the active document.




```
ActiveDocument.Tables(1).Columns(1).Select
```

Use the  **[Column](Word.Cell.Column.md)** property with a **[Cell](Word.Cell.md)** object to return a **Column** object. The following example deletes the text in cell one, inserts new text, and then sorts the entire column.




```
With ActiveDocument.Tables(1).Cell(1, 1) 
 .Range.Delete 
 .Range.InsertBefore "Sales" 
 .Column.Sort 
End With
```

Use the  **[Add](Word.Columns.Add.md)** method to add a column to a table. The following example adds a column to the first table in the active document, and then it makes the column widths equal.




```
If ActiveDocument.Tables.Count >= 1 Then 
 Set myTable = ActiveDocument.Tables(1) 
 myTable.Columns.Add BeforeColumn:=myTable.Columns(1) 
 myTable.Columns.DistributeWidth 
End If
```

Remarks

Use the  **[Information](Word.Selection.Information.md)** property with a **[Selection](Word.Selection.md)** object to return the current column number. The following example selects the current column and then displays the column number in a message box.




```
If Selection.Information(wdWithInTable) = True Then 
 Selection.Columns(1).Select 
 MsgBox "Column " _ 
 &amp; Selection.Information(wdStartOfRangeColumnNumber) 
End If
```


## Methods



|**Name**|
|:-----|
|[AutoFit](Word.Column.AutoFit.md)|
|[Delete](Word.Column.Delete.md)|
|[Select](Word.Column.Select.md)|
|[SetWidth](Word.Column.SetWidth.md)|
|[Sort](Word.Column.Sort.md)|

## Properties



|**Name**|
|:-----|
|[Application](Word.Column.Application.md)|
|[Borders](Word.Column.Borders.md)|
|[Cells](Word.Column.Cells.md)|
|[Creator](Word.Column.Creator.md)|
|[Index](Word.Column.Index.md)|
|[IsFirst](Word.Column.IsFirst.md)|
|[IsLast](Word.Column.IsLast.md)|
|[NestingLevel](Word.Column.NestingLevel.md)|
|[Next](Word.Column.Next.md)|
|[Parent](Word.Column.Parent.md)|
|[PreferredWidth](Word.Column.PreferredWidth.md)|
|[PreferredWidthType](Word.Column.PreferredWidthType.md)|
|[Previous](Word.Column.Previous.md)|
|[Shading](Word.Column.Shading.md)|
|[Width](column-width-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
