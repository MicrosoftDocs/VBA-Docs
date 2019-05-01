---
title: ReportItem.Close event (Outlook)
ms.prod: outlook
api_name:
- Outlook.ReportItem.Close
ms.assetid: d20e50a8-c73d-d866-0cd0-d6085a3b6eb6
ms.date: 06/08/2017
localization_priority: Normal
---


# ReportItem.Close event (Outlook)

Occurs when the inspector associated with an item (which is an instance of the parent object) is being closed.


## Syntax

_expression_.**Close** (_Cancel_)

_expression_ A variable that represents a [ReportItem](Outlook.ReportItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the close operation is not completed and the inspector is left open.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False**, the close operation isn't completed and the inspector is left open.

If you use the  **[Close](Outlook.ReportItem.Close(method).md)** method to fire this event, it can only be canceled if the **Close** method uses the **olPromptForSave** argument.


## See also


[ReportItem Object](Outlook.ReportItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]