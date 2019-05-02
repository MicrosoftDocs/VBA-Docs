---
title: MailMerge.Destination property (Word)
keywords: vbawd10.chm153092099
f1_keywords:
- vbawd10.chm153092099
ms.prod: word
api_name:
- Word.MailMerge.Destination
ms.assetid: 05c6ac16-afd9-f611-abc4-d115ad01bce3
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMerge.Destination property (Word)

Returns or sets the destination of the mail merge results. Read/write  **WdMailMergeDestination**.


## Syntax

_expression_.**Destination**

_expression_ Required. A variable that represents a '[MailMerge](Word.MailMerge.md)' object.


## Example

This example sends the results of a mail merge operation to a new document.


```vb
Dim mmTemp As MailMerge 
 
Set mmTemp = ActiveDocument.MailMerge 
 
If mmTemp.State = wdMainAndDataSource Then 
 mmTemp.Destination = wdSendToNewDocument 
 mmTemp.Execute 
End If
```


## See also


[MailMerge Object](Word.MailMerge.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]