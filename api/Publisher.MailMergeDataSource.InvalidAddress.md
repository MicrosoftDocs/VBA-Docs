---
title: MailMergeDataSource.InvalidAddress property (Publisher)
keywords: vbapb10.chm6291472
f1_keywords:
- vbapb10.chm6291472
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.InvalidAddress
ms.assetid: c1857edc-260b-c9c2-8624-d6628e0733c4
ms.date: 06/11/2019
localization_priority: Normal
---


# MailMergeDataSource.InvalidAddress property (Publisher)

**True** to mark a record in a mail merge data source if it contains invalid data. Read/write **Boolean**.


## Syntax

_expression_.**InvalidAddress**

_expression_ A variable that represents a **[MailMergeDataSource](Publisher.MailMergeDataSource.md)** object.


## Return value

Boolean


## Remarks

Use the **[SetAllErrorFlags](Publisher.MailMergeDataSource.SetAllErrorFlags.md)** method to set both the **InvalidAddress** and **[InvalidComments](Publisher.MailMergeDataSource.InvalidComments.md)** properties for all records in a data source.


## Example

This example searches the records to verify that the length of the PostalCode field for each record is at least five digits long. If it is not, the record is excluded from the mail merge and flagged as invalid.

```vb
Sub ExcludeRecords() 
 Dim intRecord As Integer 
 With ActiveDocument.MailMerge 
 For intRecord = 1 To .DataSource.RecordCount 
 .DataSource.ActiveRecord = intRecord 
 If Len(.DataSource.DataFields("PostalCode").Value) < 5 Then 
 With .DataSource 
 .Included = False 
 .InvalidAddress = True 
 .InvalidComments = "This record is removed " & _ 
 "from the mail merge because its postal code" & _ 
 "has less than five digits." 
 End With 
 End If 
 Next 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]