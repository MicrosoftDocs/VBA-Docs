---
title: NameSpace.OpenSharedItem method (Outlook)
keywords: vbaol11.chm789
f1_keywords:
- vbaol11.chm789
ms.prod: outlook
api_name:
- Outlook.NameSpace.OpenSharedItem
ms.assetid: ebfed85c-0af5-eb72-7a58-ae9e8b655347
ms.date: 06/08/2017
localization_priority: Normal
---


# NameSpace.OpenSharedItem method (Outlook)

Opens a shared item from a specified path or URL.


## Syntax

_expression_. `OpenSharedItem`( `_Path_` )

 _expression_ An expression that returns a '[NameSpace](Outlook.NameSpace.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Path_|Required| **String**|The path or URL of the shared item to be opened.|

## Return value

An **Object** representing the appropriate Outlook item for the shared item.


## Remarks

This method is used to open iCalendar appointment (.ics) files, vCard (.vcf) files, and Outlook message (.msg) files. The type of object returned by this method depends on the type of shared item opened, as described in the following table.



| **Shared item type**| **Outlook item**|
|:-----|:-----|
|iCalendar appointment (.ics) file| **[AppointmentItem](Outlook.AppointmentItem.md)**|
|vCard (.vcf) file| **[ContactItem](Outlook.ContactItem.md)**|
|Outlook message (.msg) file|Type corresponds to the type of the item that was saved as the .msg file|

> [!NOTE] 
> This method does not support iCalendar calendar (.ics) files. To open iCalendar calendar files, you can use the  **[OpenSharedFolder](Outlook.NameSpace.OpenSharedFolder.md)** method of the **NameSpace** object.


## See also


[NameSpace Object](Outlook.NameSpace.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
