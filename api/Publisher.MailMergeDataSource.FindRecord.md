---
title: MailMergeDataSource.FindRecord method (Publisher)
keywords: vbapb10.chm6291480
f1_keywords:
- vbapb10.chm6291480
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.FindRecord
ms.assetid: a4b37255-bdff-ac61-6d18-05a4fe008beb
ms.date: 06/11/2019
localization_priority: Normal
---


# MailMergeDataSource.FindRecord method (Publisher)

Searches the contents of the specified mail merge data source for text in a particular field. Returns a **Boolean** indicating whether the search text is found; **True** if the search text is found.


## Syntax

_expression_.**FindRecord** (_FindText_, _Field_)

_expression_ A variable that represents a **[MailMergeDataSource](Publisher.MailMergeDataSource.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_FindText_|Required| **String**|The text to look for.|
|_Field_|Optional| **String**|The name of the field to be searched.|

## Return value

Boolean


## Example

This example displays a merge publication for the first record in which the FirstName field contains Joe. If the record is found, the record number is stored in a variable.

```vb
Sub FindDataSourceRecord() 
 Dim dsMain As MailMergeDataSource 
 Dim intRecord As Integer 
 
 'Makes the data in the data source records instead of the field codes 
 ActiveDocument.MailMerge.ViewMailMergeFieldCodes = False 
 
 Set dsMain = ActiveDocument.MailMerge.DataSource 
 
 If dsMain.FindRecord(FindText:="Joe", _ 
 Field:="FirstName") = True Then 
 intRecord = dsMain.ActiveRecord 
 End If 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]