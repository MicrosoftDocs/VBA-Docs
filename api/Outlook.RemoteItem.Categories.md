---
title: RemoteItem.Categories property (Outlook)
keywords: vbaol11.chm1590
f1_keywords:
- vbaol11.chm1590
ms.prod: outlook
api_name:
- Outlook.RemoteItem.Categories
ms.assetid: 7e4639b6-4fa5-ff9b-640e-d96702dc09e1
ms.date: 06/08/2017
localization_priority: Normal
---


# RemoteItem.Categories property (Outlook)

Returns or sets a  **String** representing the categories assigned to the Outlook item. Read/write.


## Syntax

_expression_. `Categories`

_expression_ A variable that represents a [RemoteItem](Outlook.RemoteItem.md) object.


## Remarks

 **Categories** is a delimited string of category names that have been assigned to an Outlook item. This property uses the character specified in the value name, **sList**, under **HKEY_CURRENT_USER\Control Panel\International** in the Windows registry, as the delimiter for multiple categories. To convert the string of category names to an array of category names, use the Microsoft Visual Basic function **Split**.


## See also


[RemoteItem Object](Outlook.RemoteItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]