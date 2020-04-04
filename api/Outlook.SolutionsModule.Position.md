---
title: SolutionsModule.Position property (Outlook)
keywords: vbaol11.chm3366
f1_keywords:
- vbaol11.chm3366
ms.prod: outlook
api_name:
- Outlook.SolutionsModule.Position
ms.assetid: e2e0c1d7-f08a-e291-f3de-1454d6a239a8
ms.date: 06/08/2017
localization_priority: Normal
---


# SolutionsModule.Position property (Outlook)

Returns or sets a **Long** value that represents the ordinal position of the **[SolutionsModule](Outlook.SolutionsModule.md)** object when it is displayed in the **Navigation Pane**. Read/write.


## Syntax

_expression_.**Position**

_expression_ A variable that represents a [SolutionsModule](Outlook.SolutionsModule.md) object.


## Remarks

This property can only be set to a value between 1 and 9. An error occurs if you attempt to set this property to a value outside that range. If no solutions exist in the  **Solutions** module, setting or getting the **Position** property also raises an error.

Changing the value of this property for a given  **SolutionsModule** object changes the **Position** values of other navigation modules in a **[NavigationModules](Outlook.NavigationModules.md)** collection, depending on the relative change between the new value and the original value.


- If the new value is less than the original value, the specified  **SolutionsModule** object moves up to the new position and the other navigation modules that are already at or below that new position move down.
    
- If the new value is greater than the original value, the specified  **SolutionsModule** object moves down to the new position and the other navigation modules that are between the old position and the new position move up, filling the old position.
    



## See also


[SolutionsModule Object](Outlook.SolutionsModule.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]