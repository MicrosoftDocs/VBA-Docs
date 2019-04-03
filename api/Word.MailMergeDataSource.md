---
title: MailMergeDataSource object (Word)
keywords: vbawd10.chm2333
f1_keywords:
- vbawd10.chm2333
ms.prod: word
api_name:
- Word.MailMergeDataSource
ms.assetid: f86f7d3c-d7ab-45e8-21e7-fd5a426e0391
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMergeDataSource object (Word)

Represents the mail merge data source in a mail merge operation.


## Remarks

Use the  **DataSource** property to return the **MailMergeDataSource** object. The following example displays the name of the data source associated with the active document.


```vb
If ActiveDocument.MailMerge.DataSource.Name <> "" Then _ 
 MsgBox ActiveDocument.MailMerge.DataSource.Name
```

The following example displays the field names in the data source associated with the active document.




```vb
For Each aField In ActiveDocument.MailMerge.DataSource.FieldNames 
 MsgBox aField.Name 
Next aField
```

The following example opens the data source associated with Form letter.doc and determines whether the FirstName field includes the name "Kate."




```vb
With Documents("Form letter.doc").MailMerge 
 .EditDataSource 
 If .DataSource.FindRecord(FindText:="Kate", _ 
 Field:="FirstName") = True Then 
 MsgBox "Data was found" 
 End If 
End With
```


## Methods



|Name|
|:-----|
|[Close](Word.MailMergeDataSource.Close.md)|
|[FindRecord](Word.MailMergeDataSource.FindRecord.md)|
|[SetAllErrorFlags](Word.MailMergeDataSource.SetAllErrorFlags.md)|
|[SetAllIncludedFlags](Word.MailMergeDataSource.SetAllIncludedFlags.md)|

## Properties



|Name|
|:-----|
|[ActiveRecord](Word.MailMergeDataSource.ActiveRecord.md)|
|[Application](Word.MailMergeDataSource.Application.md)|
|[ConnectString](Word.MailMergeDataSource.ConnectString.md)|
|[Creator](Word.MailMergeDataSource.Creator.md)|
|[DataFields](Word.MailMergeDataSource.DataFields.md)|
|[FieldNames](Word.MailMergeDataSource.FieldNames.md)|
|[FirstRecord](Word.MailMergeDataSource.FirstRecord.md)|
|[HeaderSourceName](Word.MailMergeDataSource.HeaderSourceName.md)|
|[HeaderSourceType](Word.MailMergeDataSource.HeaderSourceType.md)|
|[Included](Word.MailMergeDataSource.Included.md)|
|[InvalidAddress](Word.MailMergeDataSource.InvalidAddress.md)|
|[InvalidComments](Word.MailMergeDataSource.InvalidComments.md)|
|[LastRecord](Word.MailMergeDataSource.LastRecord.md)|
|[MappedDataFields](Word.MailMergeDataSource.MappedDataFields.md)|
|[Name](Word.MailMergeDataSource.Name.md)|
|[Parent](Word.MailMergeDataSource.Parent.md)|
|[QueryString](Word.MailMergeDataSource.QueryString.md)|
|[RecordCount](Word.MailMergeDataSource.RecordCount.md)|
|[TableName](Word.MailMergeDataSource.TableName.md)|
|[Type](Word.MailMergeDataSource.Type.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]