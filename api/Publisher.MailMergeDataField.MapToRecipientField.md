---
title: MailMergeDataField.MapToRecipientField method (Publisher)
keywords: vbapb10.chm6422563
f1_keywords:
- vbapb10.chm6422563
ms.prod: publisher
api_name:
- Publisher.MailMergeDataField.MapToRecipientField
ms.assetid: d3da8a00-e2ca-b07b-cc8f-02d729cb149c
ms.date: 06/11/2019
localization_priority: Normal
---


# MailMergeDataField.MapToRecipientField method (Publisher)

Maps a field (column) in a particular data source represented by the parent **MailMergeDataField** object to a recipient field (column) in the master data source (combined mail-merge recipient list).


## Syntax

_expression_.**MapToRecipientField** (_bstrValue_)

_expression_ A variable that represents a **[MailMergeDataField](Publisher.MailMergeDataField.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_bstrValue_|Optional| **String**|The name of the recipient field that the data source column is to be mapped to.|

## Remarks

This method works only if the parent **MailMergeDataField** object has not already been mapped to a recipient field. You can use the **[IsMapped](Publisher.MailMergeDataField.IsMapped.md)** property to determine if the object has already been mapped.

If you do not pass a value for the optional _bstrValue_ parameter, Microsoft Publisher assumes that the field to be mapped has the same name as the recipient field in the master data source to which it is mapped.

If you pass the name of a field that does not exist, Publisher returns an error. 

> [!NOTE] 
> To add a field, use the **[AddToRecipientFields](Publisher.MailMergeDataField.AddToRecipientFields.md)** method.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **MapToRecipientField** method to map a data field (column) in a particular data source to a field in the master data source (combined recipient list) for the publication.

Before running this macro, replace `datasourceindex` with the index number of a valid data source in the data source collection of the active document, replace `fieldname` with the name of the field in the data source that you want to map to a recipient field, and replace `recipientfieldname` with the name of the recipient field.

> [!NOTE] 
> For an example of how you can use the **Name** property of the **DataSource** object to determine the index number of the data source that you want, see the **[MailMergeDataSources.Item](Publisher.MailMergeDataSources.Item.md)** method. 

```vb
Public Sub Map() 
 
 Dim pubMailMergeDataSources As Publisher.MailMergeDataSources 
 Dim pubMailMergeDataField As Publisher.MailMergeDataField 
 
 Set pubMailMergeDataSources = ThisDocument.MailMerge.DataSource.DataSources 
 Set pubMailMergeDataField = pubMailMergeDataSources.Item(datasourceindex).DataFields.Item("fieldname") 
 
 If pubMailMergeDataField.IsMapped Then 
 
 Debug.Print "This field is already mapped" 
 
 Else 
 
 pubMailMergeDataField.MapToRecipientField ("recipientfieldname") 
 Debug.Print "Field mapped successfully." 
 
 End If 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]