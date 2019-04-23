---
title: OlkOptionButton.BeforeUpdate event (Outlook)
keywords: vbaol11.chm1000191
f1_keywords:
- vbaol11.chm1000191
ms.prod: outlook
api_name:
- Outlook.OlkOptionButton.BeforeUpdate
ms.assetid: a6f40320-1cbb-08bd-b9b0-7e70b25d4529
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkOptionButton.BeforeUpdate event (Outlook)

Occurs when the data in the control is changed through the user interface and is about to be saved to the item. 


## Syntax

_expression_.**BeforeUpdate** (_Cancel_)

_expression_ A variable that represents an [OlkOptionButton](Outlook.OlkOptionButton.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the operation will not be completed and the property bound to the control will not be updated.|

## Remarks

Canceling this property will revert the control to the current value of the property and return the focus to the control.

 **BeforeUpdate** and **AfterUpdate** can occur any time the data in the control is being saved to the item. The typical sequence of events involving **BeforeUpdate** for this control is as follows:


1. User focuses on the control
    
2.  **BeforeUpdate**
    
3. Control data is updated
    
4.  ** AfterUpdate**
    
5.  **Exit** : User moves focus away from control
    



## See also


[OlkOptionButton Object](Outlook.OlkOptionButton.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]