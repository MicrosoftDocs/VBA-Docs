---
title: Selection.TopLevelTables property (Word)
keywords: vbawd10.chm158663662
f1_keywords:
- vbawd10.chm158663662
ms.prod: word
api_name:
- Word.Selection.TopLevelTables
ms.assetid: 7ab1b2a3-85a8-8892-53b9-dc85ff747078
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.TopLevelTables property (Word)

Returns a  **[Tables](Word.tables.md)** collection that represents the tables at the outermost nesting level in the current selection. Read-only.


## Syntax

_expression_. `TopLevelTables`

_expression_ A variable that represents a **[Selection](Word.Selection.md)** object.


## Remarks

This method returns a collection containing only those tables at the outermost nesting level within the context of the current selection. These tables may not be at the outermost nesting level within the entire set of nested tables.

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example creates a new document, creates a nested table with three levels, and then fills the first cell of each table with its nesting level. The example selects the second column of the second-level table and then selects the first of the top-level tables in this selection. The innermost table is selected, even though it isn't a top-level table within the context of the entire set of nested tables.


```vb
Documents.Add 
ActiveDocument.Tables.Add Selection.Range, _ 
 3, 3, wdWord9TableBehavior, wdAutoFitContent 
With ActiveDocument.Tables(1).Range 
 .Copy 
 .Cells(1).Range.Text = .Cells(1).NestingLevel 
 .Cells(5).Range.PasteAsNestedTable 
 With .Cells(5).Tables(1).Range 
 .Cells(1).Range.Text = .Cells(1).NestingLevel 
 .Cells(5).Range.PasteAsNestedTable 
 With .Cells(5).Tables(1).Range 
 .Cells(1).Range.Text = _ 
 .Cells(1).NestingLevel 
 End With 
 .Columns(2).Select 
 Selection.TopLevelTables(1).Select 
 End With 
End With
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]