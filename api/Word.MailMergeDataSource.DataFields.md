---
title: MailMergeDataSource.DataFields property (Word)
keywords: vbawd10.chm152895499
f1_keywords:
- vbawd10.chm152895499
ms.prod: word
api_name:
- Word.MailMergeDataSource.DataFields
ms.assetid: 613c4bc6-bd87-fbdc-2170-8a1daf2cfd2c
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMergeDataSource.DataFields property (Word)

Returns a  **[MailMergeDataFields](Word.mailmergedatafields.md)** collection that represents the fields in the specified mail merge data source. Read-only.


## Syntax

_expression_. `DataFields`

_expression_ A variable that represents a '[MailMergeDataSource](Word.MailMergeDataSource.md)' object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example displays the name of each field in the data source attached to the active mail merge main document.


```vb
Dim mmdfTemp As MailMergeDataField 
 
For Each mmdfTemp In _ 
 ActiveDocument.MailMerge.DataSource.DataFields 
 MsgBox mmdfTemp.Name 
Next mmdfTemp
```

This example displays the value of the LastName field from the first record in the data source attached to "Main.doc."




```vb
With Documents("Main.doc").MailMerge.DataSource 
 .ActiveRecord = wdFirstRecord 
 MsgBox .DataFields("LastName").Value 
End With
```


## See also


[MailMergeDataSource Object](Word.MailMergeDataSource.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]