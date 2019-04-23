---
title: AttachmentSelection.Parent property (Outlook)
keywords: vbaol11.chm2942
f1_keywords:
- vbaol11.chm2942
ms.prod: outlook
api_name:
- Outlook.AttachmentSelection.Parent
ms.assetid: 1c80c1fd-b7bd-288c-d017-8159ddcbd037
ms.date: 06/08/2017
localization_priority: Normal
---


# AttachmentSelection.Parent property (Outlook)

Returns the parent  **Object** of the specified object. Read-only.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents an '[AttachmentSelection](Outlook.AttachmentSelection.md)' object.


## Remarks

The  **Parent** property of an **AttachmentSelection** object represents the Microsoft Outlook item that contains the selected attachments.

If the item is in an explorer, the value of the  **Parent** property is the same as the first item in the selection that is returned by the **[Explorer.Selection](Outlook.Explorer.Selection.md)** property, which is `Explorer.Selection.Item(1)`. 

If the item is in an inspector, the value of the  **Parent** property is the same as the value of the **[Inspector.CurrentItem](Outlook.Inspector.CurrentItem.md)** property.


## See also


[AttachmentSelection Object](Outlook.AttachmentSelection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]