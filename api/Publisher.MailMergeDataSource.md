---
title: MailMergeDataSource object (Publisher)
keywords: vbapb10.chm6356991
f1_keywords:
- vbapb10.chm6356991
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource
ms.assetid: a02eb4fb-7db7-e533-c3ca-95bc4ca68e82
ms.date: 05/31/2019
localization_priority: Normal
---


# MailMergeDataSource object (Publisher)

Represents the data source in a mail merge or catalog merge operation.
 
## Remarks

Use the **[MailMerge.DataSource](Publisher.MailMerge.DataSource.md)** property to return the **MailMergeDataSource** object. 

## Example

The following example displays the name of the data source associated with the active publication.

```vb
Sub ShowDataSourceName() 
 If ActiveDocument.MailMerge.DataSource.Name <> "" Then _ 
 MsgBox ActiveDocument.MailMerge.DataSource.Name 
End Sub
```

<br/>

The following example tests the open data source associated with the active publication to determine whether the LastName field includes the name Fuller.

```vb
Sub FindSelectedRecord() 
 With ActiveDocument.MailMerge 
 If .DataSource.FindRecord(FindText:="Fuller", _ 
 Field:="LastName") = True Then 
 MsgBox "Data was found" 
 End If 
 End With 
End Sub
```


## Methods

- [ApplyFilter](Publisher.MailMergeDataSource.ApplyFilter.md)
- [Close](Publisher.MailMergeDataSource.Close.md)
- [EditRecord](Publisher.MailMergeDataSource.EditRecord.md)
- [FindRecord](Publisher.MailMergeDataSource.FindRecord.md)
- [OpenRecipientsDialog](Publisher.MailMergeDataSource.OpenRecipientsDialog.md)
- [SetAllErrorFlags](Publisher.MailMergeDataSource.SetAllErrorFlags.md)
- [SetAllIncludedFlags](Publisher.MailMergeDataSource.SetAllIncludedFlags.md)
- [SetSortOrder](Publisher.MailMergeDataSource.SetSortOrder.md)

## Properties

- [ActiveRecord](Publisher.MailMergeDataSource.ActiveRecord.md)
- [Application](Publisher.MailMergeDataSource.Application.md)
- [ConnectString](Publisher.MailMergeDataSource.ConnectString.md)
- [DataFields](Publisher.MailMergeDataSource.DataFields.md)
- [DataSources](Publisher.MailMergeDataSource.DataSources.md)
- [EverValidated](Publisher.MailMergeDataSource.EverValidated.md)
- [Filters](Publisher.MailMergeDataSource.Filters.md)
- [FirstRecord](Publisher.MailMergeDataSource.FirstRecord.md)
- [Included](Publisher.MailMergeDataSource.Included.md)
- [InvalidAddress](Publisher.MailMergeDataSource.InvalidAddress.md)
- [InvalidComments](Publisher.MailMergeDataSource.InvalidComments.md)
- [IsMaster](Publisher.MailMergeDataSource.IsMaster.md)
- [LastRecord](Publisher.MailMergeDataSource.LastRecord.md)
- [MappedDataFields](Publisher.MailMergeDataSource.MappedDataFields.md)
- [Name](Publisher.MailMergeDataSource.Name.md)
- [Parent](Publisher.MailMergeDataSource.Parent.md)
- [RecordCount](Publisher.MailMergeDataSource.RecordCount.md)
- [TableName](Publisher.MailMergeDataSource.TableName.md)
- [Type](Publisher.MailMergeDataSource.Type.md)
- [ValidatedClean](Publisher.MailMergeDataSource.ValidatedClean.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]