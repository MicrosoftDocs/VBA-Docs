---
title: NavigationGroup.Position property (Outlook)
keywords: vbaol11.chm2889
f1_keywords:
- vbaol11.chm2889
ms.prod: outlook
api_name:
- Outlook.NavigationGroup.Position
ms.assetid: b6fb7506-e143-97d8-ae36-0812ca8d7355
ms.date: 06/08/2017
localization_priority: Normal
---


# NavigationGroup.Position property (Outlook)

Returns or sets a **Long** value that represents the ordinal position of the **[NavigationGroup](Outlook.NavigationGroup.md)** object when displayed in the navigation pane. Read/write.


## Syntax

_expression_.**Position**

_expression_ A variable that represents a [NavigationGroup](Outlook.NavigationGroup.md) object.


## Remarks

This property can only be set to a value between 1 and the value of the  **[Count](Outlook.NavigationGroups.Count.md)** property for the parent **[NavigationGroups](Outlook.NavigationGroups.md)** object. An error occurs if you attempt to set this property to a value outside that range.

Changing the value of this property for a **NavigationGroup** object changes the **Position** values of other navigation groups contained by a **NavigationGroups** collection, depending on the relative change between the new value and the original value of the **Position** property for that **NavigationGroup** object:


- If the new value is less than the original value, then the specified  **NavigationGroup** object moves up to the new position and pushes the other navigation groups already at or below that new position down.
    
- If the new value is greater than the original value, then the specified  **NavigationGroup** object moves down to the new position and pushes the other navigation groups between the old position and the new position up, filling the old position.
    
If the navigation group is not on the navigation pane, then this property returns -1.


## See also


[NavigationGroup Object](Outlook.NavigationGroup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]