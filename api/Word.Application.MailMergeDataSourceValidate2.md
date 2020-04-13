---
title: Application.MailMergeDataSourceValidate2 event (Word)
keywords: vbawd10.chm4000029
f1_keywords:
- vbawd10.chm4000029
ms.prod: word
api_name:
- Word.Application.MailMergeDataSourceValidate2
ms.assetid: dba0dc60-a8c7-7e0c-ac02-4f5311534c89
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.MailMergeDataSourceValidate2 event (Word)

Occurs when a user validates mail merge recipients by clicking the **Validate addresses** link button in the **Mail Merge Recipients** dialog box.


## Syntax

_expression_. `MailMergeDataSourceValidate2`( `_Doc_` , `_Handled_` )

_expression_ A variable that represents an '[Application](Word.Application.md)' object declared with events in a class module.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **Document**|The mail merge main document.|
| _Handled_|Required| **Boolean**| **True** if the add-in has handled the validation event.|

## Remarks

If you do not have address verification software installed on your computer, the **MailMergeDataSourceValidate2** event allows you to create simple filtering routines, such as looping through records to check the postal codes and removing any that are non-U.S.


> [!NOTE] 
> You cannot raise this event from within a Microsoft Visual Basic for Applications (VBA) project. This event functions correctly only in managed add-ins and external applications. For COM add-ins, use the **[MailMergeDataSourceValidate](Word.Application.MailMergeDataSourceValidate.md)** event.


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]