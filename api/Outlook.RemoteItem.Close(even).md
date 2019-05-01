---
title: RemoteItem.Close event (Outlook)
ms.prod: outlook
api_name:
- Outlook.RemoteItem.Close
ms.assetid: 77276903-9e9e-713a-5844-c4efd36a020d
ms.date: 06/08/2017
localization_priority: Normal
---


# RemoteItem.Close event (Outlook)

Occurs when the inspector associated with an item (which is an instance of the parent object) is being closed.


## Syntax

_expression_.**Close** (_Cancel_)

_expression_ A variable that represents a [RemoteItem](Outlook.RemoteItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the close operation is not completed and the inspector is left open.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False**, the close operation isn't completed and the inspector is left open.

If you use the  **[Close](Outlook.RemoteItem.Close(method).md)** method to fire this event, it can only be canceled if the **Close** method uses the **olPromptForSave** argument.


## See also


[RemoteItem Object](Outlook.RemoteItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]