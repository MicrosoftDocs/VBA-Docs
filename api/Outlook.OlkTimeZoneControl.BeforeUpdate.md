---
title: OlkTimeZoneControl.BeforeUpdate event (Outlook)
keywords: vbaol11.chm1000529
f1_keywords:
- vbaol11.chm1000529
ms.prod: outlook
api_name:
- Outlook.OlkTimeZoneControl.BeforeUpdate
ms.assetid: f30947cb-39ae-5b5b-ffb5-a5b3281e837a
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkTimeZoneControl.BeforeUpdate event (Outlook)

Occurs when the data in the control is changed through the user interface and is about to be saved to the item. 


## Syntax

_expression_.**BeforeUpdate** (_Cancel_)

_expression_ A variable that represents an [OlkTimeZoneControl](Outlook.OlkTimeZoneControl.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the operation will not be completed and the property bound to the control will not be updated.|

## Remarks

Canceling this property will revert the control to the current value of the property and return the focus to the control.

 **BeforeUpdate** and **[AfterUpdate](Outlook.OlkTimeZoneControl.AfterUpdate.md)** can occur any time the data in the control is being saved to the item. The typical sequence of events involving **AfterUpdate** for this control is as follows:


1. User focuses on the control
    
2.  **BeforeUpdate** occurs
    
3. Control data is updated
    
4.  **AfterUpdate** occurs
    
5.  **[Exit](Outlook.OlkTimeZoneControl.Exit.md)** occurs: User moves focus away from control
    



## See also


[OlkTimeZoneControl Object](Outlook.OlkTimeZoneControl.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]