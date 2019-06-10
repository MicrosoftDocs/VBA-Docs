---
title: MailMergeDataField.Insert method (Publisher)
keywords: vbapb10.chm6422561
f1_keywords:
- vbapb10.chm6422561
ms.prod: publisher
api_name:
- Publisher.MailMergeDataField.Insert
ms.assetid: 54482cda-d0d3-c799-7e7f-b25835a8bd6f
ms.date: 06/11/2019
localization_priority: Normal
---


# MailMergeDataField.Insert method (Publisher)

Returns a **[Shape](Publisher.Shape.md)** object that represents a data field inserted into a publication.


## Syntax

_expression_.**Insert** (_Range_)

_expression_ A variable that represents a **[MailMergeDataField](Publisher.MailMergeDataField.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Range_|Optional| **[TextRange](Publisher.TextRange.md)** |The text range to insert.|

## Return value

Shape


## Remarks

The **Insert** method works for both picture and string (text) fields.

> [!NOTE] 
> You can also use the **[InsertMailMergeField](Publisher.TextRange.InsertMailMergeField.md)** method of the **TextRange** object to add a text data field to a text box in the publication's catalog merge area.


## Example

This example defines a data field as a picture data field, inserts it into the catalog merge area of the specified publication, and sizes and positions the picture data field. This example assumes that the publication has been connected to a data source, and a catalog merge area has been added to the publication.

```vb
Dim pbPictureField1 As Shape 
 
 'Define the field as a picture data type 
 With ThisDocument.MailMerge.DataSource.DataFields 
 .Item("Photo:").FieldType = pbMailMergeDataFieldPicture 
 End With 
 
 'Insert a picture field, and then size and position it 
 Set pbPictureField1 = ThisDocument.MailMerge.DataSource.DataFields.Item("Photo:").Insert 
 With pbPictureField1 
 .Height = 100 
 .Width = 100 
 .Top = 85 
 .Left = 375 
 End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]