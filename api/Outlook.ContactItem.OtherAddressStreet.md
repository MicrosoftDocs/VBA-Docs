---
title: ContactItem.OtherAddressStreet property (Outlook)
keywords: vbaol11.chm1055
f1_keywords:
- vbaol11.chm1055
ms.prod: outlook
api_name:
- Outlook.ContactItem.OtherAddressStreet
ms.assetid: dd82de5e-63fc-18bb-5211-f8218e08354b
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.OtherAddressStreet property (Outlook)

Returns or sets a  **String** representing the street portion of the other address for the contact. Read/write.


## Syntax

_expression_. `OtherAddressStreet`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property is parsed from the  **[OtherAddress](Outlook.ContactItem.OtherAddress.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to **OtherAddress**.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]