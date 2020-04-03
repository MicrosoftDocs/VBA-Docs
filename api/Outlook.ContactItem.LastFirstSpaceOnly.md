---
title: ContactItem.LastFirstSpaceOnly property (Outlook)
keywords: vbaol11.chm1030
f1_keywords:
- vbaol11.chm1030
ms.prod: outlook
api_name:
- Outlook.ContactItem.LastFirstSpaceOnly
ms.assetid: ab1e1edc-23af-ceaf-64e7-d8604c689752
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.LastFirstSpaceOnly property (Outlook)

Returns a  **String** representing the concatenated last name, first name, and middle name of the contact with spaces between them. Read-only.


## Syntax

_expression_. `LastFirstSpaceOnly`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

 This property is parsed from the **[LastName](Outlook.ContactItem.LastName.md)**, **[FirstName](Outlook.ContactItem.FirstName.md)**, and **[MiddleName](Outlook.ContactItem.MiddleName.md)** properties. The **LastName**, **FirstName**, and **MiddleName** properties are themselves parsed from the **[FullName](Outlook.ContactItem.FullName.md)** property. The value of this property is only filled when its associated property (**FirstName**, **LastName**, **MiddleName**, **CompanyName**, and **Suffix**) contain Asian (DBCS) characters. If the corresponding field does not contain Asian characters, the property will be empty.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]