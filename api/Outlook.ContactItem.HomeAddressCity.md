---
title: ContactItem.HomeAddressCity property (Outlook)
keywords: vbaol11.chm1013
f1_keywords:
- vbaol11.chm1013
ms.prod: outlook
api_name:
- Outlook.ContactItem.HomeAddressCity
ms.assetid: 1d2334f2-0401-3bcc-53bf-fa55e1664d9c
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.HomeAddressCity property (Outlook)

Returns or sets a  **String** representing the city portion of the home address for the contact. Read/write.


## Syntax

_expression_. `HomeAddressCity`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property is parsed from the  **[HomeAddress](Outlook.ContactItem.HomeAddress.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to **HomeAddress**.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]