---
title: JournalModule.Position property (Outlook)
keywords: vbaol11.chm2868
f1_keywords:
- vbaol11.chm2868
ms.prod: outlook
api_name:
- Outlook.JournalModule.Position
ms.assetid: 87cd12a7-b414-4f47-a204-7997f6d25989
ms.date: 06/08/2017
localization_priority: Normal
---


# JournalModule.Position property (Outlook)

Returns or sets a **Long** value that represents the ordinal position of the **[JournalModule](Outlook.JournalModule.md)** object when it is displayed in the navigation pane. Read/write.


## Syntax

_expression_.**Position**

_expression_ A variable that represents a [JournalModule](Outlook.JournalModule.md) object.


## Remarks

This property can only be set to a value between 1 and 9. An error occurs if you attempt to set this property to a value outside that range.

Changing the value of this property for a given  **JournalModule** object changes the **Position** values of other navigation modules in a **[NavigationModules](Outlook.NavigationModules.md)** collection, depending on the relative change between the new value and the original value.


- If the new value is less than the original value, the specified  **JournalModule** object moves up to the new position and the other navigation modules that are already at or below that new position move down.
    
- If the new value is greater than the original value, the specified  **JournalModule** object moves down to the new position and the other navigation modules that are between the old position and the new position move up, filling the old position.
    

## See also


[JournalModule Object](Outlook.JournalModule.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]