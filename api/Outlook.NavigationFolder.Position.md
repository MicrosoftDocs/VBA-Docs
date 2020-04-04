---
title: NavigationFolder.Position property (Outlook)
keywords: vbaol11.chm2907
f1_keywords:
- vbaol11.chm2907
ms.prod: outlook
api_name:
- Outlook.NavigationFolder.Position
ms.assetid: cfa86104-c191-51f8-4da3-dc3c26d6a7ed
ms.date: 06/08/2017
localization_priority: Normal
---


# NavigationFolder.Position property (Outlook)

Returns or sets an **Long** value that represents the ordinal position of the **[NavigationFolder](Outlook.NavigationFolder.md)** object when displayed in the navigation pane. Read/write.


## Syntax

_expression_.**Position**

_expression_ A variable that represents a [NavigationFolder](Outlook.NavigationFolder.md) object.


## Remarks

This property can only be set to a value between 1 and the value of the  **[Count](Outlook.NavigationFolders.Count.md)** property for the parent **[NavigationFolders](Outlook.NavigationFolders.md)** object. An error occurs if you attempt to set this property to a value outside that range.

Changing the value of this property for a **NavigationFolder** object changes the **Position** values of other navigation folders contained by a **NavigationFolders** collection, depending on the relative change between the new value and the original value of the **Position** property for that **NavigationFolder** object:


- If the new value is less than the original value, then the specified  **NavigationFolder** object moves up to the new position and pushes the other navigation folders already at or below that new position down.
    
- If the new value is greater than the original value, then the specified  **NavigationFolder** object moves down to the new position and pushes the other navigation folders between the old position and the new position up, filling the old position.
    
If the navigation folder has been removed from the navigation pane, then this property returns -1 to indicate that the navigation folder is no longer part of the navigation group.


## See also


[NavigationFolder Object](Outlook.NavigationFolder.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]