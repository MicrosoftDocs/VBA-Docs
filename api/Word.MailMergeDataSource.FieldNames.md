---
title: MailMergeDataSource.FieldNames property (Word)
keywords: vbawd10.chm152895498
f1_keywords:
- vbawd10.chm152895498
ms.prod: word
api_name:
- Word.MailMergeDataSource.FieldNames
ms.assetid: 3e88ee90-c44e-1dbb-dcfd-6ea99cbb1c2c
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMergeDataSource.FieldNames property (Word)

Returns a  **[MailMergeFieldNames](Word.MailMergeFieldNames.md)** collection that represents the names of all the fields in the specified mail merge data source. Read-only.


## Syntax

_expression_. `FieldNames`

_expression_ A variable that represents a '[MailMergeDataSource](Word.MailMergeDataSource.md)' object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example displays the name of the first field in the data source attached to the active mail merge main document.


```vb
MsgBox ActiveDocument.MailMerge.DataSource.FieldNames(1).Name
```

This example uses the mNames() array to store the names of each merge field contained in the data source attached to the active document.




```vb
Dim mNames As Variant 
Dim mmTemp As MailMerge 
Dim intCount As Integer 
Dim intIncrement As Integer 
Dim mmfnLoop As MailMergeFieldName 
 
Set mmTemp = ActiveDocument.MailMerge 
intCount = _ 
 ActiveDocument.MailMerge.DataSource.FieldNames.Count - 1 
 
ReDim mNames(intCount) 
intIncrement = 0 
 
For Each mmfnLoop In mmTemp.DataSource.FieldNames 
 mNames(intIncrement) = mmfnLoop.Name 
 intIncrement = intIncrement + 1 
Next mmfnLoop
```


## See also


[MailMergeDataSource Object](Word.MailMergeDataSource.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]