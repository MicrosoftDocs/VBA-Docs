---
title: NavigationModule.Position property (Outlook)
keywords: vbaol11.chm2809
f1_keywords:
- vbaol11.chm2809
ms.prod: outlook
api_name:
- Outlook.NavigationModule.Position
ms.assetid: cdf7eedb-18a4-028c-8663-eae70e466617
ms.date: 06/08/2017
localization_priority: Normal
---


# NavigationModule.Position property (Outlook)

Returns or sets a  **Long** value that represents the ordinal position of the **[NavigationModule](Outlook.NavigationModule.md)** object when displayed in the navigation pane. Read/write.


## Syntax

_expression_.**Position**

 _expression_ An expression that returns a [NavigationModule](Outlook.NavigationModule.md) object.


## Remarks

This property can only be set to a value between 1 and 8. An error occurs if you attempt to set this property to a value outside that range.

Changing the value of this property for a  **NavigationModule** object changes the **Position** values of other navigation modules contained by a **[NavigationModules](Outlook.NavigationModules.md)** collection, depending on the relative change between the new value and the original value of the **Position** property for that **NavigationModule** object:


- If the new value is less than the original value, then the specified  **NavigationModule** object moves up to the new position and pushes the other navigation modules already at or below that new position down.
    
- If the new value is greater than the original value, then the specified  **NavigationModule** object moves down to the new position and pushes the other navigation modules between the old position and the new position up, filling the old position.
    

## See also


[NavigationModule Object](Outlook.NavigationModule.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]