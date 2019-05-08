---
title: MappedDataField.DataFieldIndex property (Word)
keywords: vbawd10.chm107544581
f1_keywords:
- vbawd10.chm107544581
ms.prod: word
api_name:
- Word.MappedDataField.DataFieldIndex
ms.assetid: ba10017b-5ac4-483d-2c37-6e41286aaf65
ms.date: 06/08/2017
localization_priority: Normal
---


# MappedDataField.DataFieldIndex property (Word)

Returns or sets a  **Long** that represents the corresponding field number in the mail merge data source to which a mapped data field maps. Read/write.


## Syntax

_expression_. `DataFieldIndex`

_expression_ A variable that represents a '[MappedDataField](Word.MappedDataField.md)' object.


## Remarks

This property returns zero if the specified data field is not mapped to a mapped data field.


## Example

This example maps the PostalAddress1 field in the data source to the wdAddress1 mapped data field. This example assumes that the current document is a mail merge document.


```vb
Sub MapField() 
 With ActiveDocument.MailMerge.DataSource 
 .MappedDataFields(wdAddress1).DataFieldIndex = _ 
 .FieldNames("PostalAddress1").Index 
 End With 
End Sub
```


## See also


[MappedDataField Object](Word.MappedDataField.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]