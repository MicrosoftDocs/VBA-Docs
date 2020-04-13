---
title: Store.DisplayName property (Outlook)
keywords: vbaol11.chm800
f1_keywords:
- vbaol11.chm800
ms.prod: outlook
api_name:
- Outlook.Store.DisplayName
ms.assetid: 785ec583-3553-6002-41b6-d0c6d0028b5a
ms.date: 06/08/2017
localization_priority: Normal
---


# Store.DisplayName property (Outlook)

Returns a **String** representing the display name of the **[Store](Outlook.Store.md)** object. Read-only.


## Syntax

_expression_.**DisplayName**

_expression_ A variable that represents a [Store](Outlook.Store.md) object.


## Remarks

 **DisplayName** is the default property of the **Store** object. This property corresponds to the MAPI property, **PidTagDisplayName**.

 **DisplayName** is read-only. To change the **DisplayName** of a Personal Folders File (.pst), use the **[PropertyAccessor](Outlook.PropertyAccessor.md)** object and the **[PropertyAccessor.SetProperty](Outlook.PropertyAccessor.SetProperty.md)** method.


## See also


[Store Object](Outlook.Store.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]