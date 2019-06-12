---
title: MailMergeDataSource.SetAllErrorFlags method (Publisher)
keywords: vbapb10.chm6291488
f1_keywords:
- vbapb10.chm6291488
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.SetAllErrorFlags
ms.assetid: 17c41fbb-3b21-c31a-63cd-ed26065bfa79
ms.date: 06/11/2019
localization_priority: Normal
---


# MailMergeDataSource.SetAllErrorFlags method (Publisher)

Marks all records in a mail merge data source as containing invalid data in an address field.


## Syntax

_expression_.**SetAllErrorFlags** (_Invalid_, _InvalidComment_)

_expression_ A variable that represents a **[MailMergeDataSource](Publisher.MailMergeDataSource.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Invalid_|Required| **Boolean**| **True** marks all records in the data source of a mail merge as invalid.|
|_InvalidComment_|Optional| **String**|Text describing the invalid setting.|

## Remarks

You can individually mark records in a data source that contain invalid data in an address field by using the **[InvalidAddress](Publisher.MailMergeDataSource.InvalidAddress.md)** and **[InvalidComments](Publisher.MailMergeDataSource.InvalidComments.md)** properties.


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