---
title: ContactItem.MailingAddress property (Outlook)
keywords: vbaol11.chm1034
f1_keywords:
- vbaol11.chm1034
ms.prod: outlook
api_name:
- Outlook.ContactItem.MailingAddress
ms.assetid: 7af2770c-1f8b-510b-4e6f-3ef919082088
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.MailingAddress property (Outlook)

Returns or sets a **String** representing the full, unparsed selected mailing address for the contact. Read/write.


## Syntax

_expression_. `MailingAddress`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property replicates the property indicated by the  **[SelectedMailingAddress](Outlook.ContactItem.SelectedMailingAddress.md)** property, which is one of the following **OlMailingAddress** constants: **olBusiness**, **olHome**, **olNone**, or **olOther**. While it can be changed or entered independently, any such changes or entries to this property will be overwritten by any subsequent changes or entries to the property indicated by **SelectedMailingAddress**.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]