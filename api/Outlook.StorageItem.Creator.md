---
title: StorageItem.Creator property (Outlook)
keywords: vbaol11.chm2152
f1_keywords:
- vbaol11.chm2152
ms.prod: outlook
api_name:
- Outlook.StorageItem.Creator
ms.assetid: c89c777c-5f4b-f672-ff74-d34db3bcd790
ms.date: 06/08/2017
localization_priority: Normal
---


# StorageItem.Creator property (Outlook)

Returns and sets the solution that created the  **[StorageItem](Outlook.StorageItem.md)** object. Read/write.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a [StorageItem](Outlook.StorageItem.md) object.


## Remarks

Outlook does not set the  **Creator** property. Use the **Creator** property to identify the **StorageItem** objects you have created for your add-in. One recommended value for this property is the programmatic identifier (ProgID) of the add-in.


## See also


[StorageItem Object](Outlook.StorageItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]