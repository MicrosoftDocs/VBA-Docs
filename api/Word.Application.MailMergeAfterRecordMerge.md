---
title: Application.MailMergeAfterRecordMerge event (Word)
keywords: vbawd10.chm4000017
f1_keywords:
- vbawd10.chm4000017
ms.prod: word
api_name:
- Word.Application.MailMergeAfterRecordMerge
ms.assetid: 6f483874-3999-815d-28b3-69fef89ed2be
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.MailMergeAfterRecordMerge event (Word)

Occurs after each record in the data source successfully merges in a mail merge.


## Syntax

_expression_.**MailMergeAfterRecordMerge** (_Doc_)

_expression_ A variable that represents an '[Application](Word.Application.md)' object that has been declared with events in a class module. For information about using events with the **Application** object, see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md).


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **Document**|The mail merge main document.|

## Example

This example displays a message with the value of the first and second fields in the record that has just finished merging. This example assumes that you have declared an application variable called MailMergeApp in your general declarations and have set the variable equal to the Word Application object.


```vb
Private Sub MailMergeApp_MailMergeAfterRecordMerge(ByVal Doc As Document) 
 
 With Doc.MailMerge.DataSource 
 MsgBox .DataFields(1).Value & " " & _ 
 .DataFields(2).Value & " is finished merging." 
 End With 
 
End Sub
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]