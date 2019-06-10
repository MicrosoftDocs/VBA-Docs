---
title: MailMergeDataSource.DataFields property (Publisher)
keywords: vbapb10.chm6291461
f1_keywords:
- vbapb10.chm6291461
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.DataFields
ms.assetid: 820af882-d54c-a205-2925-e7110fc0c02b
ms.date: 06/11/2019
localization_priority: Normal
---


# MailMergeDataSource.DataFields property (Publisher)

Returns a **[MailMergeDataFields](Publisher.MailMergeDataFields.md)** collection that represents the fields in the specified data source.


## Syntax

_expression_.**DataFields**

_expression_ A variable that represents a **[MailMergeDataSource](Publisher.MailMergeDataSource.md)** object.


## Return value

MailMergeDataFields


## Example

This example displays the value of the FirstName and LastName fields from the active record in the data source attached to the active publication.

```vb
Sub ShowNameForActiveRecord() 
 Dim mdfFirst As MailMergeDataField 
 Dim mdfLast As MailMergeDataField 
 
 With ActiveDocument.MailMerge.DataSource 
 Set mdfFirst = .DataFields.Item("FirstName") 
 Set mdfLast = .DataFields.Item("LastName") 
 MsgBox "The active record in the attached " & _ 
 vbLf & "data source is : " & _ 
 mdfFirst.Value & " " & _ 
 mdfLast.Value 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]