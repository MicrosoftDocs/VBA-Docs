---
title: MailMergeDataField.UnMapRecipientField method (Publisher)
keywords: vbapb10.chm6422564
f1_keywords:
- vbapb10.chm6422564
ms.prod: publisher
api_name:
- Publisher.MailMergeDataField.UnMapRecipientField
ms.assetid: 0063dfa7-1168-3701-56a3-f1908cf0d23a
ms.date: 06/11/2019
localization_priority: Normal
---


# MailMergeDataField.UnMapRecipientField method (Publisher)

Undoes the mapping between the parent **MailMergeDataField** object in a particular data source and the recipient field in the master data source (combined mail-merge recipient list) to which it is currently mapped.


## Syntax

_expression_.**UnMapRecipientField**

_expression_ A variable that represents a **[MailMergeDataField](Publisher.MailMergeDataField.md)** object.


## Remarks

This method works only if the parent **MailMergeDataField** object is mapped to a recipient field. You can use the **[IsMapped](Publisher.MailMergeDataField.IsMapped.md)** property to determine if the object is mapped.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **UnmapRecipientField** method to undo the mapping between a data field (column) in a particular data source and a field in the master data source (combined recipient list) for the publication.

Before running this macro, replace `datasourceindex` with the index number of a valid data source in the data source collection of the active document, and replace `fieldname` with the name of the field in the data source that you want to remove from the combined list of recipient fields.

> [!NOTE] 
> For an example of how you can use the **Name** property of the **DataSource** object to determine the index number of the data source that you want, see the **[MailMergeDataSources.Item](Publisher.MailMergeDataSources.Item.md)** method.

```vb
Public Sub UnmapRecipientField_Example() 
 
 Dim pubMailMergeDataSources As Publisher.MailMergeDataSources 
 Dim pubMailMergeDataField As Publisher.MailMergeDataField 
 
 Set pubMailMergeDataSources = ThisDocument.MailMerge.DataSource.DataSources 
 Set pubMailMergeDataField = pubMailMergeDataSources.Item(datasourceindex).DataFields.Item("fieldname") 
 
 If pubMailMergeDataField.IsMapped Then 
 
 pubMailMergeDataField.UnMapRecipientField 
 Debug.Print "Field unmapped successfully." 
 
 Else 
 
 Debug.Print "This field is not mapped." 
 
 End If 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]