---
title: MailMergeDataSource.InvalidComments property (Publisher)
keywords: vbapb10.chm6291473
f1_keywords:
- vbapb10.chm6291473
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.InvalidComments
ms.assetid: ee08b03a-57e2-d79c-ee9f-a6f9231c8d6b
ms.date: 06/11/2019
localization_priority: Normal
---


# MailMergeDataSource.InvalidComments property (Publisher)

If the **[InvalidAddress](Publisher.MailMergeDataSource.InvalidAddress.md)** property is **True**, this property returns or sets a **String** that describes invalid data in a mail merge record. Read/write.


## Syntax

_expression_.**InvalidComments**

_expression_ A variable that represents a **[MailMergeDataSource](Publisher.MailMergeDataSource.md)** object.


## Return value

String


## Remarks

Use the **[SetAllErrorFlags](Publisher.MailMergeDataSource.SetAllErrorFlags.md)** method to set both the **InvalidAddress** and **InvalidComments** properties for all records in a data source.


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