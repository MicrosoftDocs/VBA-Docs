---
title: MailMergeDataField object (Publisher)
keywords: vbapb10.chm6488063
f1_keywords:
- vbapb10.chm6488063
ms.prod: publisher
api_name:
- Publisher.MailMergeDataField
ms.assetid: 46768b72-482c-06c5-5e77-27a95109f610
ms.date: 05/31/2019
localization_priority: Normal
---


# MailMergeDataField object (Publisher)

Represents a single merge field in a data source. The **MailMergeDataField** object is a member of the **[MailMergeDataFields](Publisher.MailMergeDataFields.md)** collection. The **MailMergeDataFields** collection includes all the data fields in a mail merge or catalog merge data source (for example, Name, Address, and City).
 

## Remarks

You cannot add fields to the **MailMergeDataFields** collection. All data fields in a data source are automatically included in the **MailMergeDataFields** collection.
 
Use **[MailMergeDataSource.DataFields](Publisher.MailMergeDataSource.DataFields.md)** (_index_), where _index_ is the data field name or index number, to return a single **MailMergeDataField** object. The index number represents the position of the data field in the mail merge data source. 
 

## Example

This example retrieves the name of the first field and value of the first record of the FirstName field in the data source attached to the active publication.
 
```vb
Sub GetDataFromSource() 
 With ActiveDocument.MailMerge.DataSource 
 MsgBox "Field Name: " & .DataFields.Item(1).Name & _ 
 "Value: " & .DataFields.Item("FirstName").Value 
 End With 
End Sub
```


## Methods

- [AddToRecipientFields](Publisher.MailMergeDataField.AddToRecipientFields.md)
- [Insert](Publisher.MailMergeDataField.Insert.md)
- [MapToRecipientField](Publisher.MailMergeDataField.MapToRecipientField.md)
- [UnMapRecipientField](Publisher.MailMergeDataField.UnMapRecipientField.md)

## Properties

- [Application](Publisher.MailMergeDataField.Application.md)
- [Creator](Publisher.MailMergeDataField.Creator.md)
- [FieldType](Publisher.MailMergeDataField.FieldType.md)
- [Index](Publisher.MailMergeDataField.Index.md)
- [IsMapped](Publisher.MailMergeDataField.IsMapped.md)
- [MappedTo](Publisher.MailMergeDataField.MappedTo.md)
- [Name](Publisher.MailMergeDataField.Name.md)
- [Parent](Publisher.MailMergeDataField.Parent.md)
- [Value](Publisher.MailMergeDataField.Value.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]