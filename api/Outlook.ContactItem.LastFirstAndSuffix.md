---
title: ContactItem.LastFirstAndSuffix property (Outlook)
keywords: vbaol11.chm1027
f1_keywords:
- vbaol11.chm1027
ms.prod: outlook
api_name:
- Outlook.ContactItem.LastFirstAndSuffix
ms.assetid: b234614c-e2c0-cba2-6ec8-69be1a31caf1
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.LastFirstAndSuffix property (Outlook)

Returns a  **String** representing the last name, first name, middle name, and suffix of the contact. Read-only.


## Syntax

_expression_. `LastFirstAndSuffix`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

There is a comma between the last and first names and spaces between all the names and the suffix. This property is parsed from the  **[LastName](Outlook.ContactItem.LastName.md)**, **[FirstName](Outlook.ContactItem.FirstName.md)**, **[MiddleName](Outlook.ContactItem.MiddleName.md)** and **[Suffix](Outlook.ContactItem.Suffix.md)** properties. The **LastName**, **FirstName**, and **Suffix** properties are themselves parsed from the **[FullName](Outlook.ContactItem.FullName.md)** property. The value of this property is only filled when its associated property (**FirstName**, **LastName**, **MiddleName**, **CompanyName**, and **Suffix**) contain Asian (DBCS) characters. If the corresponding field does not contain Asian characters, the property will be empty.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]