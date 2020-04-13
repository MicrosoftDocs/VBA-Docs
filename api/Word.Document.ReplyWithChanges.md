---
title: Document.ReplyWithChanges method (Word)
keywords: vbawd10.chm158007650
f1_keywords:
- vbawd10.chm158007650
ms.prod: word
api_name:
- Word.Document.ReplyWithChanges
ms.assetid: ad476bde-0240-ab4b-b246-d5b143207fa5
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.ReplyWithChanges method (Word)

Sends an email message to the author of a document that has been sent out for review, notifying them that a reviewer has completed review of the document.


## Syntax

_expression_. `ReplyWithChanges`( `_ShowMessage_` )

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ShowMessage_|Optional| **Variant**| **True** to display the message prior to sending. **False** to automatically send the message without displaying it first. The default value is **True**.|

## Remarks

Use the **SendForReview** method to start a collaborative review of a document. If the **ReplyWithChanges** method is executed on a document that is not part of a collaborative review cycle, Microsoft Word displays an error message.


## Example

This example sends a message notifying the author that a reviewer has completed a review, without first displaying the email message to the reviewer. This example assumes that the current document is part of a collaborative review cycle.


```vb
Sub ReplyMsg() 
 ActiveDocument.ReplyWithChanges ShowMessage:=False 
End Sub
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]