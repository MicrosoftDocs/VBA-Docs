---
title: MailMergeDataFields object (Publisher)
keywords: vbapb10.chm6422527
f1_keywords:
- vbapb10.chm6422527
ms.prod: publisher
api_name:
- Publisher.MailMergeDataFields
ms.assetid: 44ae8a3c-b8a8-fc57-9d02-d71dcffc21ef
ms.date: 05/31/2019
localization_priority: Normal
---


# MailMergeDataFields object (Publisher)

A collection of **[MailMergeDataField](Publisher.MailMergeDataField.md)** objects that represent the data fields in a mail merge or catalog merge data source.
 

## Remarks

You cannot add fields to the **MailMergeDataFields** collection. When a data field is added to a data source, the field is automatically included in the **MailMergeDataFields** collection.
 
Use the **[MailMergeDataSource.DataFields](Publisher.MailMergeDataSource.DataFields.md)** property to return the **MailMergeDataFields** collection.

Use **DataFields** (_index_), where _index_ is the data field name or the index number, to return a single **MailMergeDataField** object. The index number represents the position of the data field in the mail merge data source. 

## Example

The following example displays the field names in the data source attached to the active publication.

```vb
Sub ShowFieldNames() 
 Dim intCount As Integer 
 With ActiveDocument.MailMerge.DataSource.DataFields 
 For intCount = 1 To .Count 
 MsgBox .Item(intCount).Name 
 Next 
 End With 
End Sub
```

<br/>

This example retrieves the name of the first field and value of the first record of the FirstName field in the data source attached to the active publication.

```vb
Sub GetDataFromSource() 
 With ActiveDocument.MailMerge.DataSource.DataFields 
 MsgBox "First field name: " & .Item(1).Name & vbLf & _ 
 "Value of the first record of the FirstName field: " & _ 
 .Item("FirstName").Value 
 End With 
End Sub
```


## Methods

- [Item](Publisher.MailMergeDataFields.Item.md)

## Properties

- [Application](Publisher.MailMergeDataFields.Application.md)
- [Count](Publisher.MailMergeDataFields.Count.md)
- [Creator](Publisher.MailMergeDataFields.Creator.md)
- [Parent](Publisher.MailMergeDataFields.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]