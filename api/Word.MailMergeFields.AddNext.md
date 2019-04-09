---
title: MailMergeFields.AddNext method (Word)
keywords: vbawd10.chm153026667
f1_keywords:
- vbawd10.chm153026667
ms.prod: word
api_name:
- Word.MailMergeFields.AddNext
ms.assetid: c267f484-b9b0-44a0-c519-ca6624057223
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMergeFields.AddNext method (Word)

Adds a NEXT field to a mail merge main document. Returns a  **MailMergeField** object.


## Syntax

_expression_. `AddNext`( `_Range_` )

_expression_ Required. A variable that represents a '[MailMergeFields](Word.mailmergefields.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range object**|The location for the NEXT field.|

## Return value

MailMergeField


## Remarks

A NEXT field advances to the next record so that data from more than one record can be merged into the same merge document (for example, a sheet of mailing labels).


## Example

This example adds a NEXT field after the third MERGEFIELD field in Main.doc.


```vb
Documents("Main.doc").MailMerge.Fields(3).Select 
Selection.Collapse Direction:=wdCollapseEnd 
Documents("Main.doc").MailMerge.Fields.AddNext _ 
 Range:=Selection.Range
```


## See also


[MailMergeFields Collection Object](Word.mailmergefields.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]