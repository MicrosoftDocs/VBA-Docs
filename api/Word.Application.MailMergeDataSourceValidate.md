---
title: Application.MailMergeDataSourceValidate event (Word)
keywords: vbawd10.chm4000021
f1_keywords:
- vbawd10.chm4000021
ms.prod: word
api_name:
- Word.Application.MailMergeDataSourceValidate
ms.assetid: 31e03b87-b76c-9cfe-afb0-c9ee5cbcd13b
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.MailMergeDataSourceValidate event (Word)

Occurs when a user validates mail merge recipients by clicking  **Validate** in the **Mail Merge Recipients** dialog box.


## Syntax

_expression_.**MailMergeDataSourceValidate** (_Doc As Document_**, **_Handled As Boolean_**)

_expression_ A variable that represents an '[Application](Word.Application.md)' object that has been declared with events in a class module.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **Document**|The mail merge main document.|
| _Handled_|Required| **Boolean**| **True** if the add-in has handled the validation event. This is a forward-only parameter and cannot be set in code. To set this value, you must use the **[MailMergeDataSourceValidate2](Word.Application.MailMergeDataSourceValidate2.md)** event.|

## Remarks

If you do not have address verification software installed on your computer, the **MailMergeDataSourceValidate** event allows you to create simple filtering routines, such as looping through records to check the postal codes and removing any that are non-U.S.


> [!NOTE] 
> The Handled parameter does not function correctly in this version of the event; use the **[MailMergeDataSourceValidate2](Word.Application.MailMergeDataSourceValidate2.md)** event. In addition, you cannot raise this event from within a Microsoft Visual Basic for Applications (VBA) project. This event functions correctly only in COM add-ins. For managed add-ins and external applications, use the **MailMergeDataSourceValidate2** event.

For information about using events with the **Application** object, see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md).


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]