---
title: MailMergeDataSource.SetAllIncludedFlags method (Publisher)
keywords: vbapb10.chm6291481
f1_keywords:
- vbapb10.chm6291481
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.SetAllIncludedFlags
ms.assetid: ab668e95-55ac-fcbd-19c9-3c13fe3aa995
ms.date: 06/11/2019
localization_priority: Normal
---


# MailMergeDataSource.SetAllIncludedFlags method (Publisher)

**True** to include all data source records in a mail merge.


## Syntax

_expression_.**SetAllIncludedFlags** (_Included_)

_expression_ A variable that represents a **[MailMergeDataSource](Publisher.MailMergeDataSource.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Included_|Required| **Boolean**| **True** to include all data source records in a mail merge. **False** to exclude all data source records from a mail merge.|

## Remarks

You can set individual records in a data source to be included in or excluded from a mail merge by using the **[Included](Publisher.MailMergeDataSource.Included.md)** property.


## Example

This example marks all records in the data source as containing an invalid address field, sets a comment as to why it is invalid, and excludes all records from the mail merge.

```vb
Sub FlagAllRecords() 
 With ActiveDocument.MailMerge.DataSource 
 .SetAllErrorFlags Invalid:=True, InvalidComment:= _ 
 "All records in the data source have only 5-" _ 
 & "digit ZIP Codes. Need 5+4 digit ZIP Codes." 
 .SetAllIncludedFlags Included:=False 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]