---
title: MailMergeDataSource.OpenRecipientsDialog method (Publisher)
keywords: vbapb10.chm6291490
f1_keywords:
- vbapb10.chm6291490
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.OpenRecipientsDialog
ms.assetid: 5a0a2b4a-ce23-435c-6e18-f778d6e14fd6
ms.date: 06/11/2019
localization_priority: Normal
---


# MailMergeDataSource.OpenRecipientsDialog method (Publisher)

Displays the **Recipients** dialog box for a mail merge publication.


## Syntax

_expression_.**OpenRecipientsDialog**

_expression_ A variable that represents a **[MailMergeDataSource](Publisher.MailMergeDataSource.md)** object.


## Example

This example displays the **Mail Merge Recipients** dialog box.

```vb
Sub ShowRecipientsDialog() 
 ActiveDocument.MailMerge.DataSource.OpenRecipientsDialog 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]