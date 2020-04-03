---
title: ContactItem.MailingAddressCity property (Outlook)
keywords: vbaol11.chm1035
f1_keywords:
- vbaol11.chm1035
ms.prod: outlook
api_name:
- Outlook.ContactItem.MailingAddressCity
ms.assetid: f9b8510a-998a-bf7e-9fa5-f567f9d784bc
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.MailingAddressCity property (Outlook)

Returns or sets a **String** representing the city name portion of the selected mailing address of the contact. Read/write.


## Syntax

_expression_. `MailingAddressCity`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property replicates the property indicated by the  **[SelectedMailingAddress](Outlook.ContactItem.SelectedMailingAddress.md)** property, which is one of the following **OlMailingAddress** constants: **olBusiness**, **olHome**, **olNone**, or **olOther**. While it can be changed or entered independently, any such changes or entries to this property will be overwritten by any subsequent changes or entries to the property indicated by **SelectedMailingAddress**.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]