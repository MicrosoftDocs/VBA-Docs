---
title: MailMergeDataField.FieldType property (Publisher)
ms.prod: publisher
api_name:
- Publisher.Field.FieldType
ms.assetid: 9574f59b-a03f-ab0b-a2ac-085f31473f78
ms.date: 06/11/2019
localization_priority: Normal
---


# MailMergeDataField.FieldType property (Publisher)

Returns a **[PbMailMergeDataFieldType](Publisher.PbMailMergeDataFieldType.md)** constant that represents the type of data contained in the data field.


## Syntax

_expression_.**FieldType**

_expression_ A variable that represents a **[MailMergeDataField](Publisher.MailMergeDataField.md)** object.


## Return value

PbMailMergeDataFieldType


## Remarks

Use the **[Insert](Publisher.MailMergeDataField.Insert.md)** method to add a picture data field to a publication's catalog merge area.

Use the **[InsertMailMergeField](Publisher.TextRange.InsertMailMergeField.md)** method of the **TextRange** object to add a text data field to a text box in the publication's catalog merge area.

The **FieldType** property value can be one of the **PbMailMergeDataFieldType** constants declared in the Microsoft Publisher type library.


## Example

This example defines a data field as a picture data field, inserts it into the catalog merge area of the specified publication, and sizes and positions the picture data field. This example assumes that the publication has been connected to a data source, and that a catalog merge area has been added to the publication.

```vb
Dim pbPictureField1 As Shape 
 
 'Define the Photo field as a picture data type 
 With ThisDocument.MailMerge.DataSource.DataFields 
 .Item("Photo:").FieldType = pbMailMergeDataFieldPicture 
 End With 
 
 'Insert a picture field, then size and position it 
 Set pbPictureField1 = ThisDocument.MailMerge.DataSource.DataFields.Item("Photo:").Insert 
 With pbPictureField1 
 .Height = 100 
 .Width = 100 
 .Top = 85 
 .Left = 375 
 End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]