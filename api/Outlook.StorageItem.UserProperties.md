---
title: StorageItem.UserProperties property (Outlook)
keywords: vbaol11.chm2149
f1_keywords:
- vbaol11.chm2149
ms.prod: outlook
api_name:
- Outlook.StorageItem.UserProperties
ms.assetid: 0a08e77c-1665-a612-2f47-ef1c3fc331d2
ms.date: 06/08/2017
localization_priority: Normal
---


# StorageItem.UserProperties property (Outlook)

Returns the  **[UserProperties](Outlook.UserProperties.md)** collection that represents all the user properties for the Outlook item. Read-only.


## Syntax

_expression_. `UserProperties`

_expression_ A variable that represents a [StorageItem](Outlook.StorageItem.md) object.


## Remarks

If you use the  **[UserProperties.Add](Outlook.UserProperties.Add.md)** method on the **[UserProperties](Outlook.UserProperties.md)** object associated with a **[StorageItem](Outlook.StorageItem.md)**, the optional _AddToFolderFields_ and _DisplayFormat_ arguments of the **UserProperties.Add** method will be ignored. Any custom properties of the **StorageItem** object will not be exposed as custom properties in the **Field Chooser**.


## See also


[StorageItem Object](Outlook.StorageItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]