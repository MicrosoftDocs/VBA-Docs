---
title: TasksModule.Position property (Outlook)
keywords: vbaol11.chm2848
f1_keywords:
- vbaol11.chm2848
ms.prod: outlook
api_name:
- Outlook.TasksModule.Position
ms.assetid: a4166776-dc7a-07d3-b9a2-b187c7a3d821
ms.date: 06/08/2017
localization_priority: Normal
---


# TasksModule.Position property (Outlook)

Returns or sets a  **Long** value that represents the ordinal position of the **[TasksModule](Outlook.TasksModule.md)** object when it is displayed in the navigation pane. Read/write.


## Syntax

_expression_.**Position**

_expression_ A variable that represents a [TasksModule](Outlook.TasksModule.md) object.


## Remarks

This property can only be set to a value between 1 and 9. An error occurs if you attempt to set this property to a value outside that range.

Changing the value of this property for a given  **TasksModule** object changes the **Position** values of other navigation modules in a **[NavigationModules](Outlook.NavigationModules.md)** collection, depending on the relative change between the new value and the original value.


- If the new value is less than the original value, the specified  **TasksModule** object moves up to the new position and the other navigation modules that are already at or below that new position move down.
    
- If the new value is greater than the original value, the specified  **TasksModule** object moves down to the new position and the other navigation modules that are between the old position and the new position move up, filling the old position.
    

## See also


[TasksModule Object](Outlook.TasksModule.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]