---
title: MailMergeDataSource.SetSortOrder method (Publisher)
keywords: vbapb10.chm6291489
f1_keywords:
- vbapb10.chm6291489
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.SetSortOrder
ms.assetid: 0ecb5f77-2cd1-92c6-b7f2-bf709f015ba5
ms.date: 06/11/2019
localization_priority: Normal
---


# MailMergeDataSource.SetSortOrder method (Publisher)

Sets the sort order for mail merge data.


## Syntax

_expression_.**SetSortOrder** (_SortField1_, _SortAscending1_, _SortField2_, _SortAscending2_, _SortField3_, _SortAscending3_)

_expression_ A variable that represents a **[MailMergeDataSource](Publisher.MailMergeDataSource.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_SortField1_|Optional| **String**|The first field on which to sort the mail merge data. Default is an empty string.|
|_SortAscending1_|Optional| **Boolean**| **True** (default) to perform an ascending sort on SortField1; **False** to perform a descending sort.|
|_SortField2_|Optional| **String**|The second field on which to sort the mail merge data. Default is an empty string.|
|_SortAscending2_|Optional| **Boolean**| **True** (default) to perform an ascending sort on SortField2; **False** to perform a descending sort.|
|_SortField3_|Optional| **String**|The third field on which to sort the mail merge data. Default is an empty string.|
|_SortAscending3_|Optional| **Boolean**| **True** (default) to perform an ascending sort on SortField3; **False** to perform a descending sort.|

## Example

The following example sorts mail merge data first on postal code in descending order, and then on last name and first name in ascending order.

```vb
ActiveDocument.MailMerge.DataSource.SetSortOrder _ 
 SortField1:="ZIPCode", SortAscending1:=False, _ 
 SortField2:="LastName", SortField3:="FirstName"
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]