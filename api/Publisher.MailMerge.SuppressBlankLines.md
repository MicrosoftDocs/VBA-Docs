---
title: MailMerge.SuppressBlankLines property (Publisher)
keywords: vbapb10.chm6225927
f1_keywords:
- vbapb10.chm6225927
ms.prod: publisher
api_name:
- Publisher.MailMerge.SuppressBlankLines
ms.assetid: 3b41e0c0-8588-e86a-77ed-90c4692c03dc
ms.date: 06/08/2019
localization_priority: Normal
---


# MailMerge.SuppressBlankLines property (Publisher)

**True** to suppress blank lines when mail merge fields in a mail merge main document are empty. Read/write **Boolean**.


## Syntax

_expression_.**SuppressBlankLines**

_expression_ A variable that represents a **[MailMerge](Publisher.MailMerge.md)** object.


## Return value

Boolean


## Example

This example suppresses blank lines in the active publication when mail merge fields are blank. This example assumes that a mail merge data source is attached to the active publication.

```vb
Sub SuppressBlankLines() 
 ActiveDocument.MailMerge.SuppressBlankLines = True 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]