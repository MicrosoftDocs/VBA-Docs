---
title: StorageItem.Size property (Outlook)
keywords: vbaol11.chm2148
f1_keywords:
- vbaol11.chm2148
ms.prod: outlook
api_name:
- Outlook.StorageItem.Size
ms.assetid: 7bf2fd39-8705-aa1b-af76-a3a21073d152
ms.date: 06/08/2017
localization_priority: Normal
---


# StorageItem.Size property (Outlook)

Returns a **Long** indicating the size (in bytes) of the **[StorageItem](Outlook.StorageItem.md)**. Read-only.


## Syntax

_expression_.**Size**

_expression_ A variable that represents a [StorageItem](Outlook.StorageItem.md) object.


## Remarks

The **Size** of a **StorageItem** that is newly created is zero (0) until you make an explicit call on the **[Save](Outlook.StorageItem.Save.md)** method of the item.


## See also


[StorageItem Object](Outlook.StorageItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]