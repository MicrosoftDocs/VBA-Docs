---
title: ContactItem.MailingAddressStreet property (Outlook)
keywords: vbaol11.chm1040
f1_keywords:
- vbaol11.chm1040
ms.prod: outlook
api_name:
- Outlook.ContactItem.MailingAddressStreet
ms.assetid: 8487bbf4-0d48-4224-9370-e4e78f100d09
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.MailingAddressStreet property (Outlook)

Returns or sets a  **String** representing the street address portion of the selected mailing address of the contact. Read/write.


## Syntax

_expression_. `MailingAddressStreet`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property replicates the property indicated by the  **[SelectedMailingAddress](Outlook.ContactItem.SelectedMailingAddress.md)** property, which is one of the following **OlMailingAddress** constants: **olBusiness**, **olHome**, **olNone**, or **olOther**. While it can be changed or entered independently, any such changes or entries to this property will be overwritten by any subsequent changes or entries to the property indicated by **SelectedMailingAddress**.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]