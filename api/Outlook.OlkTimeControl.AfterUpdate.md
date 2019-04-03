---
title: OlkTimeControl.AfterUpdate event (Outlook)
keywords: vbaol11.chm1000413
f1_keywords:
- vbaol11.chm1000413
ms.prod: outlook
api_name:
- Outlook.OlkTimeControl.AfterUpdate
ms.assetid: 5454d296-9508-a4c4-37b7-9c119e29d20e
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkTimeControl.AfterUpdate event (Outlook)

Occurs after the data in the control has been changed through the user interface.


## Syntax

_expression_.**AfterUpdate**

_expression_ A variable that represents an [OlkTimeControl](Outlook.OlkTimeControl.md) object.


## Remarks

 **BeforeUpdate** and **AfterUpdate** can occur any time the data in the control is being saved to the item. The typical sequence of events involving **AfterUpdate** for this control is as follows:


1. User focuses on the control
    
2.  **BeforeUpdate**
    
3. Control data is updated
    
4.  ** AfterUpdate**
    
5.  **Exit** : User moves focus away from control
    



## See also


[OlkTimeControl Object](Outlook.OlkTimeControl.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]