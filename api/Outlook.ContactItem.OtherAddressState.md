---
title: ContactItem.OtherAddressState property (Outlook)
keywords: vbaol11.chm1054
f1_keywords:
- vbaol11.chm1054
ms.prod: outlook
api_name:
- Outlook.ContactItem.OtherAddressState
ms.assetid: a8073ae6-eb63-5674-16c1-ceb83babda25
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.OtherAddressState property (Outlook)

Returns or sets a  **String** representing the state portion of the other address for the contact. Read/write.


## Syntax

_expression_. `OtherAddressState`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property is parsed from the  **[OtherAddress](Outlook.ContactItem.OtherAddress.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to **OtherAddress**.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]