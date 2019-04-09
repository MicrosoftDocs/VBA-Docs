---
title: MailMergeFields.AddMergeRec method (Word)
keywords: vbawd10.chm153026665
f1_keywords:
- vbawd10.chm153026665
ms.prod: word
api_name:
- Word.MailMergeFields.AddMergeRec
ms.assetid: 50146076-696e-9a78-5e58-4ecb0f32607f
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMergeFields.AddMergeRec method (Word)

Adds a MERGEREC field to a mail merge main document. Returns a  **MailMergeField** object.


## Syntax

_expression_. `AddMergeRec`( `_Range_` )

_expression_ Required. A variable that represents a '[MailMergeFields](Word.mailmergefields.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range object**|The location for the MERGEREC field.|

## Return value

MailMergeField


## Remarks

A MERGEREC field inserts the number of the current record (the position of the record in the current query result) during a mail merge.


## Example

This example inserts text and a MERGEREC field at the beginning of the active document.


```vb
Dim rngTemp As Range 
 
Set rngTemp = ActiveDocument.Range(Start:=0, End:=0) 
 
ActiveDocument.MailMerge.Fields.AddMergeRec Range:=rngTemp 
rngTemp.InsertAfter "Record Number: "
```


## See also


[MailMergeFields Collection Object](Word.mailmergefields.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]