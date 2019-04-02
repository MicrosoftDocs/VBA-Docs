---
title: DocumentItem.Write event (Outlook)
ms.prod: outlook
api_name:
- Outlook.DocumentItem.Write
ms.assetid: ae8c445f-cf46-9544-7073-bf08638b9247
ms.date: 06/08/2017
localization_priority: Normal
---


# DocumentItem.Write event (Outlook)

Occurs when an instance of the parent object is saved, either explicitly (for example, using the  **[Save](Outlook.DocumentItem.Save.md)** or **[SaveAs](Outlook.DocumentItem.SaveAs.md)** methods) or implicitly (for example, in response to a prompt when closing the item's inspector).


## Syntax

_expression_. `Write`( `_Cancel_` )

_expression_ A variable that represents a [DocumentItem](Outlook.DocumentItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| (Not used in VBScript). **False** when the event occurs. If the event procedure sets this argument to **True**, the save operation is not completed.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False**, the save operation is not completed.


## See also


[DocumentItem Object](Outlook.DocumentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]