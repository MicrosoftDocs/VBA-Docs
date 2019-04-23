---
title: MailMerge.MailFormat property (Word)
keywords: vbawd10.chm153092108
f1_keywords:
- vbawd10.chm153092108
ms.prod: word
api_name:
- Word.MailMerge.MailFormat
ms.assetid: 2bfe3efa-3aee-c451-3ccc-828f64636f33
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMerge.MailFormat property (Word)

Returns a  **WdMailMergeMailFormat** constant that represents the format to use when the mail merge destination is an email message. Read/write.


## Syntax

_expression_. `MailFormat`

_expression_ Required. A variable that represents a '[MailMerge](Word.MailMerge.md)' object.


## Remarks

The  **MailFormat** property is ignored if the **MailAsAttachment** property is set to **True**. Conversely, when **MailFormat** is set, **MailAsAttachment** is automatically set to **False**.


## Example

This example merges the active document to an email message and formats it using HTML.


```vb
Sub MergeDestination() 
    With ActiveDocument.MailMerge 
        .Destination = wdSendToEmail 
        .MailAsAttachment = False 
        .MailFormat = wdMailFormatHTML 
        .Execute 
    End With 
End Sub
```


## See also


[MailMerge Object](Word.MailMerge.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]