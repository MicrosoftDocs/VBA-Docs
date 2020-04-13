---
title: MappedDataField.Value property (Word)
keywords: vbawd10.chm107544580
f1_keywords:
- vbawd10.chm107544580
ms.prod: word
api_name:
- Word.MappedDataField.Value
ms.assetid: 08567167-2aa7-ccd0-0eea-30bae7439b6b
ms.date: 06/08/2017
localization_priority: Normal
---


# MappedDataField.Value property (Word)

Returns the contents of the mail merge data field or mapped data field for the current record. Read-only  **String**.


## Syntax

_expression_.**Value**

_expression_ Required. A variable that represents a '[MappedDataField](Word.MappedDataField.md)' object.


## Remarks

Use the **ActiveRecord** property to set the active record in a mail merge data source.


## Example

This example displays the contents of the active record in the data source attached to Main.doc.


```vb
For Each dataF In _ 
 Documents("Main.doc").MailMerge.DataSource.DataFields 
 If dataF.Value <> "" Then dRecord = dRecord & _ 
 dataF.Value & vbCr 
Next dataF 
MsgBox dRecord
```


## See also


[MappedDataField Object](Word.MappedDataField.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]