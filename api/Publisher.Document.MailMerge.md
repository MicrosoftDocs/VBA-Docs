---
title: Document.MailMerge property (Publisher)
keywords: vbapb10.chm196628
f1_keywords:
- vbapb10.chm196628
ms.prod: publisher
api_name:
- Publisher.Document.MailMerge
ms.assetid: 15b1a8aa-3472-c67d-1d99-92617b05c157
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.MailMerge property (Publisher)

Returns a **[MailMerge](Publisher.MailMerge.md)** object that represents the mail merge functionality for the specified publication.


## Syntax

_expression_.**MailMerge**

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Return value

MailMerge


## Example

This example displays the information from the current record in the data source.

```vb
Sub ViewMergeData() 
 ActiveDocument.MailMerge.ViewMailMergeFieldCodes = False 
End Sub
```

<br/>

This example displays the **Mail Merge Recipients** dialog box, which contains the records from the data source.

```vb
Sub ExecuteMergeField() 
 ActiveDocument.MailMerge.DataSource.OpenRecipientsDialog 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]