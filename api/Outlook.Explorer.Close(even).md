---
title: Explorer.Close event (Outlook)
keywords: vbaol11.chm456
f1_keywords:
- vbaol11.chm456
ms.prod: outlook
api_name:
- Outlook.Explorer.Close
ms.assetid: 20586ee0-35b5-02f9-327b-8431f6083cca
ms.date: 06/08/2017
localization_priority: Normal
---


# Explorer.Close event (Outlook)

Occurs when an explorer is being closed.


## Syntax

_expression_.**Close** (_Cancel_)

_expression_ A variable that represents an **[Explorer](Outlook.Explorer.md)** object.


## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False**, the close operation isn't completed and the inspector is left open. This event cannot be cancelled.

If you use the  **[Close](Outlook.Explorer.Close(method).md)** method to fire this event, it can only be canceled if the **Close** method uses the **olPromptForSave** argument.


## See also


[Explorer Object](Outlook.Explorer.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]