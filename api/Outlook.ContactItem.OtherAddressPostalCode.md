---
title: ContactItem.OtherAddressPostalCode property (Outlook)
keywords: vbaol11.chm1052
f1_keywords:
- vbaol11.chm1052
ms.prod: outlook
api_name:
- Outlook.ContactItem.OtherAddressPostalCode
ms.assetid: a9cecb5e-d6c3-9496-8537-fab14520321f
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.OtherAddressPostalCode property (Outlook)

Returns or sets a **String** representing the postal code portion of the other address for the contact. Read/write.


## Syntax

_expression_. `OtherAddressPostalCode`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property is parsed from the  **[OtherAddress](Outlook.ContactItem.OtherAddress.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to **OtherAddress**.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]