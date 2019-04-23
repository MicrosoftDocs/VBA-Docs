---
title: MailMergeFields.AddIf method (Word)
keywords: vbawd10.chm153026664
f1_keywords:
- vbawd10.chm153026664
ms.prod: word
api_name:
- Word.MailMergeFields.AddIf
ms.assetid: 13c9338a-b70e-1132-0390-121d4daa15fb
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMergeFields.AddIf method (Word)

Adds an IF field to a mail merge main document. Returns a  **MailMergeField** object.


## Syntax

_expression_. `AddIf`( `_Range_` , `_MergeField_` , `_Comparison_` , `_CompareTo_` , `_TrueAutoText_` , `_TrueText_` , `_FalseAutoText_` , `_FalseText_` )

_expression_ Required. A variable that represents a '[MailMergeFields](Word.mailmergefields.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range object**|The location for the IF field.|
| _MergeField_|Required| **String**|The merge field name.|
| _Comparison_|Required| **WdMailMergeComparison**|The operator used in the comparison.|
| _CompareTo_|Optional| **Variant**|The text to compare with the contents of MergeField.|
| _TrueAutoText_|Optional| **Variant**|The AutoText entry that's inserted if the comparison is true. If this argument is specified, TrueText is ignored.|
| _TrueText_|Optional| **Variant**|The text that's inserted if the comparison is true.|
| _FalseAutoText_|Optional| **Variant**|The AutoText entry that's inserted if the comparison is false. If this argument is specified, FalseText is ignored.|
| _FalseText_|Optional| **Variant**|The text that's inserted if the comparison is false.|

## Return value

MailMergeField


## Remarks

When updated, an IF field compares a field in a record with a specified value, and then it inserts the appropriate text according to the result of the comparison.


## Example

This example inserts "for your personal use" if the Company merge field is blank and "for your business" if the Company merge field is not blank.


```vb
ActiveDocument.MailMerge.Fields.AddIf Range:=Selection.Range, _ 
 MergeField:="Company", Comparison:=wdMergeIfIsBlank, _ 
 TrueText:="for your personal use", _ 
 FalseText:="for your business"
```


## See also


[MailMergeFields Collection Object](Word.mailmergefields.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]