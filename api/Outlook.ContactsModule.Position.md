---
title: ContactsModule.Position property (Outlook)
keywords: vbaol11.chm2838
f1_keywords:
- vbaol11.chm2838
ms.prod: outlook
api_name:
- Outlook.ContactsModule.Position
ms.assetid: 2e71509d-1e6a-f736-2560-40c1de67711c
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactsModule.Position property (Outlook)

Returns or sets a  **Long** value that represents the ordinal position of the **[ContactsModule](Outlook.ContactsModule.md)** object when it is displayed in the navigation pane. Read/write.


## Syntax

_expression_.**Position**

_expression_ A variable that represents a [ContactsModule](Outlook.ContactsModule.md) object.


## Remarks

This property can only be set to a value between 1 and 9. An error occurs if you attempt to set this property to a value outside that range.

Changing the value of this property for a given  **ContactsModule** object changes the **Position** values of other navigation modules in a **[NavigationModules](Outlook.NavigationModules.md)** collection, depending on the relative change between the new value and the original value.


- If the new value is less than the original value, the specified  **ContactsModule** object moves up to the new position and the other navigation modules that are already at or below that new position move down.
    
- If the new value is greater than the original value, the specified  **ContactsModule** object moves down to the new position and the other navigation modules that are between the old position and the new position move up, filling the old position.
    

## See also


[ContactsModule Object](Outlook.ContactsModule.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]