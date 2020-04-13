---
title: MailMergeField object (Word)
keywords: vbawd10.chm2334
f1_keywords:
- vbawd10.chm2334
ms.prod: word
api_name:
- Word.MailMergeField
ms.assetid: 8beb6228-079c-008c-10aa-3f8f711fcf5c
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMergeField object (Word)

Represents a single mail merge field in a document. The **MailMergeDataField** object is a member of the **[MailMergeDataFields](Word.mailmergedatafields.md)** collection. The **MailMergeDataFields** collection includes all the mail merge related fields in a document.


## Remarks

Use  **Fields** (Index), where Index is the index number, to return a single **MailMergeField** object. The following example displays the field code of the first mail merge field in the active document.


```vb
MsgBox ActiveDocument.MailMerge.Fields(1).Code
```

Use the **Add** method to add a merge field to the **MailMergeFields** collection. The following example replaces the selection with a MiddleInitial merge field.




```vb
ActiveDocument.MailMerge.Fields.Add Range:=Selection.Range, _ 
 Name:="MiddleInitial"
```

The **MailMergeFields** collection has additional methods, such as **AddAsk** and **AddFillIn**, for adding fields related to a mail merge operation.


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]