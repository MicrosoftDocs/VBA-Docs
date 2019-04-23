---
title: ContactItem.LastName property (Outlook)
keywords: vbaol11.chm1032
f1_keywords:
- vbaol11.chm1032
ms.prod: outlook
api_name:
- Outlook.ContactItem.LastName
ms.assetid: 430682f6-a230-887b-404b-a71989121fa2
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.LastName property (Outlook)

Returns or sets a  **String** representing the last name for the contact. Read/write.


## Syntax

_expression_. `LastName`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property is parsed from the  **[FullName](Outlook.ContactItem.FullName.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes of entries to **FullName**.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]