---
title: MailMergeFields.AddSkipIf method (Word)
keywords: vbawd10.chm153026670
f1_keywords:
- vbawd10.chm153026670
ms.prod: word
api_name:
- Word.MailMergeFields.AddSkipIf
ms.assetid: feaa8b59-292c-0e6f-661a-af501b395cf9
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMergeFields.AddSkipIf method (Word)

Adds a SKIPIF field to a mail merge main document. Returns a  **MailMergeField** object. .


## Syntax

_expression_. `AddSkipIf`( `_Range_` , `_MergeField_` , `_Comparison_` , `_CompareTo_` )

_expression_ Required. A variable that represents a '[MailMergeFields](Word.mailmergefields.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range object**|The location for the SKIPIF field.|
| _MergeField_|Required| **String**|The merge field name.|
| _Comparison_|Required| **WdMailMergeComparison**|The operator used in the comparison.|
| _CompareTo_|Optional| **Variant**|The text to compare with the contents of MergeField.|

## Return value

MailMergeField


## Remarks

A SKIPIF field compares two expressions, and if the comparison is true, SKIPIF moves to the next record in the data source and starts a new merge document.


## Example

This example adds a SKIPIF field before the first MERGEFIELD field in Main.doc. If the next postal code equals 98040, the next record is skipped.


```vb
Documents("Main.doc").MailMerge.Fields(1).Select 
Selection.Collapse Direction:=wdCollapseStart 
Documents("Main.doc").MailMerge.Fields.AddSkipIf _ 
 Range:=Selection.Range, MergeField:="PostalCode", _ 
 Comparison:=wdMergeIfEqual, CompareTo:="98040"
```


## See also


[MailMergeFields Collection Object](Word.mailmergefields.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]