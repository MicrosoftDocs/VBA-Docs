---
title: JournalItem.Categories property (Outlook)
keywords: vbaol11.chm1235
f1_keywords:
- vbaol11.chm1235
ms.prod: outlook
api_name:
- Outlook.JournalItem.Categories
ms.assetid: 640caa61-a29f-e6d4-8833-a3d263b2d00e
ms.date: 06/08/2017
localization_priority: Normal
---


# JournalItem.Categories property (Outlook)

Returns or sets a  **String** representing the categories assigned to the Outlook item. Read/write.


## Syntax

_expression_. `Categories`

_expression_ A variable that represents a [JournalItem](Outlook.JournalItem.md) object.


## Remarks

 **Categories** is a delimited string of category names that have been assigned to an Outlook item. This property uses the character specified in the value name, **sList**, under **HKEY_CURRENT_USER\Control Panel\International** in the Windows registry, as the delimiter for multiple categories. To convert the string of category names to an array of category names, use the Microsoft Visual Basic function **Split**.


## See also


[JournalItem Object](Outlook.JournalItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]