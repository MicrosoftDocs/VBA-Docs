---
title: ContactItem.FirstName property (Outlook)
keywords: vbaol11.chm1004
f1_keywords:
- vbaol11.chm1004
ms.prod: outlook
api_name:
- Outlook.ContactItem.FirstName
ms.assetid: 403b5e5a-037b-cf21-efc2-2bd2a80c3789
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.FirstName property (Outlook)

Returns or sets a  **String** representing the first name for the contact. Read/write.


## Syntax

_expression_. `FirstName`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property is parsed from the  **[FullName](Outlook.ContactItem.FullName.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes of entries to **FullName**.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]