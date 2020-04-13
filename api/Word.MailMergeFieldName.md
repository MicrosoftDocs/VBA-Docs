---
title: MailMergeFieldName object (Word)
keywords: vbawd10.chm2331
f1_keywords:
- vbawd10.chm2331
ms.prod: word
api_name:
- Word.MailMergeFieldName
ms.assetid: f4e09d1e-0da2-2f0f-1747-566a4ae443b6
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMergeFieldName object (Word)

Represents a mail merge field name in a data source. The **MailMergeFieldName** object is a member of the **[MailMergeFieldNames](Word.MailMergeFieldNames.md)** collection. The **MailMergeFieldNames** collection includes all the data field names in a mail merge data source.


## Remarks

Use  **FieldNames** (Index), where Index is the index number, to return a single **MailMergeFieldName** object. The index number represents the position of the field in the mail merge data source. The following example retrieves the name of the last field in the data source attached to the active document.


```vb
alast = ActiveDocument.MailMerge.DataSource.FieldNames.Count 
afirst = ActiveDocument.MailMerge.DataSource.FieldNames(alast).Name 
MsgBox afirst
```

You cannot add fields to the **MailMergeFieldNames** collection. Field names in a data source are automatically included in the **MailMergeFieldNames** collection.


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]