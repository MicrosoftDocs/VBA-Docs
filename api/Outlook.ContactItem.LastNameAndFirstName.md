---
title: ContactItem.LastNameAndFirstName property (Outlook)
keywords: vbaol11.chm1033
f1_keywords:
- vbaol11.chm1033
ms.prod: outlook
api_name:
- Outlook.ContactItem.LastNameAndFirstName
ms.assetid: 7667650d-3da9-8a30-63d5-2d6b0d55ccb7
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.LastNameAndFirstName property (Outlook)

Returns a  **String** representing the concatenated last name and first name for the contact. Read-only.


## Syntax

_expression_. `LastNameAndFirstName`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property is parsed from the  **[FirstName](Outlook.ContactItem.FirstName.md)** and **[LastName](Outlook.ContactItem.LastName.md)** properties for the contact, which are themselves parsed from the **[FullName](Outlook.ContactItem.FullName.md)** property.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]