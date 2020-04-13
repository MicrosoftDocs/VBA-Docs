---
title: Conflict object (Outlook)
keywords: vbaol11.chm410
f1_keywords:
- vbaol11.chm410
ms.prod: outlook
api_name:
- Outlook.Conflict
ms.assetid: a7c8f12a-08ba-9fff-60b8-a02d1c7f6f33
ms.date: 06/08/2017
localization_priority: Normal
---


# Conflict object (Outlook)

Represents an Outlook item that is in conflict with another Outlook item.


## Remarks

 Each Outlook item has a **[Conflicts](Outlook.Conflicts.md)** collection object associated with it that represents all the items that are in conflict with that item.

Use the  **[Item](Outlook.Conflicts.Item.md)** method to retrieve a particular **Conflict** object from the **Conflicts** collection object, for example:


## Example

The following Visual Basic for Applications (VBA) example retrieves a **Conflict** object from the **Conflicts** collection object.


```vb
Set myConflictItem = myConflicts.Item(1)
```


## Properties



|Name|
|:-----|
|[Application](Outlook.Conflict.Application.md)|
|[Class](Outlook.Conflict.Class.md)|
|[Item](Outlook.Conflict.Item.md)|
|[Name](Outlook.Conflict.Name.md)|
|[Parent](Outlook.Conflict.Parent.md)|
|[Session](Outlook.Conflict.Session.md)|
|[Type](Outlook.Conflict.Type.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]