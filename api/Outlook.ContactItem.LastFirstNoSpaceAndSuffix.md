---
title: ContactItem.LastFirstNoSpaceAndSuffix property (Outlook)
keywords: vbaol11.chm1082
f1_keywords:
- vbaol11.chm1082
ms.prod: outlook
api_name:
- Outlook.ContactItem.LastFirstNoSpaceAndSuffix
ms.assetid: 15c9527b-3837-d4a0-0249-2cd751e4379f
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.LastFirstNoSpaceAndSuffix property (Outlook)

Returns a  **String** that contains the last name, first name, and suffix of the user without a space. Read-only


## Syntax

_expression_. `LastFirstNoSpaceAndSuffix`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property is used only when the  **[FirstName](Outlook.ContactItem.FirstName.md)**, **[LastName](Outlook.ContactItem.LastName.md)**, and **[Suffix](Outlook.ContactItem.Suffix.md)** properties (the fields that define this property) contain Asian (DBCS) characters. Note that any such changes or entries to the **FirstName**, **LastName**, or **Suffix** properties will be overwritten by any subsequent changes or entries to FullName.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]