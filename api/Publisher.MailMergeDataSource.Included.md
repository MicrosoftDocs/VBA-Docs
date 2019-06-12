---
title: MailMergeDataSource.Included property (Publisher)
keywords: vbapb10.chm6291465
f1_keywords:
- vbapb10.chm6291465
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.Included
ms.assetid: 1cdac925-5fd6-e1d0-4612-0641e6057a7e
ms.date: 06/11/2019
localization_priority: Normal
---


# MailMergeDataSource.Included property (Publisher)

**True** if a record is included in a mail merge. Read/write **Boolean**.


## Syntax

_expression_.**Included**

_expression_ A variable that represents a **[MailMergeDataSource](Publisher.MailMergeDataSource.md)** object.


## Return value

Boolean


## Remarks

Use the **[SetAllIncludedFlags](Publisher.MailMergeDataSource.SetAllIncludedFlags.md)** method to set the included status for all mail merge records.


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