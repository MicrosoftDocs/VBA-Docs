---
title: MailingLabel.LabelOptions method (Word)
keywords: vbawd10.chm152502375
f1_keywords:
- vbawd10.chm152502375
ms.prod: word
api_name:
- Word.MailingLabel.LabelOptions
ms.assetid: b49c8ade-59ae-f315-76f0-0a73d62e1ea7
ms.date: 06/08/2017
localization_priority: Normal
---


# MailingLabel.LabelOptions method (Word)

Displays the  **Label Options** dialog box.


## Syntax

_expression_. `LabelOptions`

_expression_ Required. A variable that represents a '[MailingLabel](Word.MailingLabel.md)' object.


## Remarks

The  **LabelOptions** method works only if the document is the main document of a mailing labels mail merge.


## Example

This example determines if the current document is a Mailing Label document, and if it is, displays the Label Options dialog box.


```vb
Sub LabelOps() 
 If ActiveDocument.MailMerge _ 
 .MainDocumentType = wdMailingLabels Then 
 Application.MailingLabel.LabelOptions 
 End If 
End Sub
```


## See also


[MailingLabel Object](Word.MailingLabel.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]