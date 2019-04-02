---
title: ContactItem.OtherAddressPostOfficeBox property (Outlook)
keywords: vbaol11.chm1053
f1_keywords:
- vbaol11.chm1053
ms.prod: outlook
api_name:
- Outlook.ContactItem.OtherAddressPostOfficeBox
ms.assetid: 905500a2-475a-ed2a-79b5-e46a3d8c117c
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.OtherAddressPostOfficeBox property (Outlook)

Returns or sets a  **String** representing the post office box portion of the other address for the contact. Read/write.


## Syntax

_expression_. `OtherAddressPostOfficeBox`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property is parsed from the  **[OtherAddress](Outlook.ContactItem.OtherAddress.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to **OtherAddress**.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]