---
title: ContactItem.MiddleName property (Outlook)
keywords: vbaol11.chm1042
f1_keywords:
- vbaol11.chm1042
ms.prod: outlook
api_name:
- Outlook.ContactItem.MiddleName
ms.assetid: 07e0c9b1-1093-2f8a-3b89-ba8570b2bdf5
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.MiddleName property (Outlook)

Returns or sets a  **String** representing the middle name for the contact. Read/write.


## Syntax

_expression_. `MiddleName`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property is parsed from the  **[FullName](Outlook.ContactItem.FullName.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes of entries to **FullName**.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]