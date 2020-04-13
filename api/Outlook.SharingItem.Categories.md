---
title: SharingItem.Categories property (Outlook)
keywords: vbaol11.chm601
f1_keywords:
- vbaol11.chm601
ms.prod: outlook
api_name:
- Outlook.SharingItem.Categories
ms.assetid: c81a3075-8934-c28a-4018-f66454aefcc5
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.Categories property (Outlook)

Returns or sets a **String** representing the categories assigned to the **[SharingItem](Outlook.SharingItem.md)**. Read/write.


## Syntax

_expression_. `Categories`

_expression_ A variable that represents a [SharingItem](Outlook.SharingItem.md) object.


## Remarks

 **Categories** is a delimited string of category names that have been assigned to an Outlook item. This property uses the character specified in the value name, **sList**, under **HKEY_CURRENT_USER\Control Panel\International** in the Windows registry, as the delimiter for multiple categories. To convert the string of category names to an array of category names, use the Microsoft Visual Basic function **Split**.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]