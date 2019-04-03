---
title: ContactItem.HomeAddressCountry property (Outlook)
keywords: vbaol11.chm1014
f1_keywords:
- vbaol11.chm1014
ms.prod: outlook
api_name:
- Outlook.ContactItem.HomeAddressCountry
ms.assetid: a3e1f178-c01c-e7df-ee4e-fc82f89915f0
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.HomeAddressCountry property (Outlook)

Returns or sets a  **String** representing the country/region portion of the home address for the contact. Read/write.


## Syntax

_expression_. `HomeAddressCountry`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property is parsed from the  **[HomeAddress](Outlook.ContactItem.HomeAddress.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to **HomeAddress**.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]