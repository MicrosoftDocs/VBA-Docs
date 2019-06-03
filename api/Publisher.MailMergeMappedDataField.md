---
title: MailMergeMappedDataField object (Publisher)
keywords: vbapb10.chm6619135
f1_keywords:
- vbapb10.chm6619135
ms.prod: publisher
api_name:
- Publisher.MailMergeMappedDataField
ms.assetid: 3711d28e-f005-27fb-88b5-8674d4ece887
ms.date: 05/31/2019
localization_priority: Normal
---


# MailMergeMappedDataField object (Publisher)

Represents a single mapped data field. The **MailMergeMappedDataField** object is a member of the **[MailMergeMappedDataFields](Publisher.MailMergeMappedDataFields.md)** collection. 

A mapped data field is a field contained within Microsoft Publisher that represents commonly used name or address information, such as First Name. If a data source contains a First Name field or a variation (such as First_Name, FirstName, First, or FName), the field in the data source automatically maps to the corresponding mapped data field. If a publication is to be merged with more than one data source, mapped data fields make it unnecessary to reenter the fields into the publication to agree with the field names in the database.
 
## Remarks

Use **[MailMergeDataSource.MappedDataFields](Publisher.MailMergeDataSource.MappedDataFields.md)** (_index_) to return a **MailMergeMappedDataField** object. 

## Example

This example returns the data source field name for the **pbFirstName** mapped data field. This example assumes that the current publication is a mail merge publication. A blank string value returned for the **DataFieldName** property indicates that the mapped data field is not mapped to a field in the data source.

```vb
Sub MappedFieldName() 
 Dim strMappedDataField As String 
 With ActiveDocument.MailMerge.DataSource 
 strMappedDataField = .MappedDataFields(pbFirstName).DataFieldName 
 If strMappedDataField <> "" Then 
 MsgBox "The mapped data field 'FirstName' is mapped to " _ 
 & .MappedDataFields(pbFirstName).DataFieldName & "." 
 Else 
 MsgBox "The mapped data field 'FirstName' is not " & _ 
 "mapped to any of the data fields in your " & _ 
 "data source." 
 End If 
 End With 
End Sub
```


## Properties

- [Application](Publisher.MailMergeMappedDataField.Application.md)
- [DataFieldName](Publisher.MailMergeMappedDataField.DataFieldName.md)
- [Index](Publisher.MailMergeMappedDataField.Index.md)
- [Name](Publisher.MailMergeMappedDataField.Name.md)
- [Parent](Publisher.MailMergeMappedDataField.Parent.md)
- [Value](Publisher.MailMergeMappedDataField.Value.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]