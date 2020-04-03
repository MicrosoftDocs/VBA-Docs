---
title: ContactItem.Categories property (Outlook)
keywords: vbaol11.chm934
f1_keywords:
- vbaol11.chm934
ms.prod: outlook
api_name:
- Outlook.ContactItem.Categories
ms.assetid: c2ac3005-caa9-cc91-766e-a341ed0d0e9e
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.Categories property (Outlook)

Returns or sets a **String** representing the categories assigned to the Outlook item. Read/write.


## Syntax

_expression_. `Categories`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

 **Categories** is a delimited string of category names that have been assigned to an Outlook item. This property uses the character specified in the value name, **sList**, under HKEY_CURRENT_USER\Control Panel\International in the Windows registry, as the delimiter for multiple categories. To convert the string of category names to an array of category names, use the Microsoft Visual Basic function **Split**.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]