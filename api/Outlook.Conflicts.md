---
title: Conflicts object (Outlook)
keywords: vbaol11.chm399
f1_keywords:
- vbaol11.chm399
ms.prod: outlook
api_name:
- Outlook.Conflicts
ms.assetid: c4e1c060-519a-a6d1-8fb2-c7dfa1e3e66f
ms.date: 06/08/2017
localization_priority: Normal
---


# Conflicts object (Outlook)

Contains a collection of  **[Conflict](Outlook.Conflict.md)** objects that represent all Microsoft Outlook items that are in conflict with a particular Outlook item.


## Remarks

Use the  **[Conflicts](Outlook.MailItem.Conflicts.md)** property of any Outlook item, such as **[MailItem](Outlook.MailItem.md)**, to return the **Conflicts** object.

Use the  **[Count](Outlook.Conflicts.Count.md)** property of the **Conflicts** object to determine if the item is involved in a conflict. A non-zero value indicates conflict.

Use the  **[Item](Outlook.Conflicts.Item.md)** method to retrieve a particular conflict item from the **Conflicts** collection object.

Use the  **[GetFirst](Outlook.Conflicts.GetFirst.md)**, **[GetNext](Outlook.Conflicts.GetNext.md)**, **[GetPrevious](Outlook.Conflicts.GetPrevious.md)**, and **[GetLast](Outlook.Conflicts.GetLast.md)** methods to traverse the **Conflicts** collection.


## Example

The following Microsoft Visual Basic for Applications (VBA) example uses the  **Count** property of the **Conflicts** object to determine if the item is involved in any conflict. To run this example, make sure an email item is open in the active window.


```vb
Sub CheckConflicts() 
 
 Dim myItem As Outlook.MailItem 
 
 Dim myConflicts As Outlook.Conflicts 
 
 
 
 Set myItem = Application.ActiveInspector.CurrentItem 
 
 Set myConflicts = myItem.Conflicts 
 
 If (myConflicts.Count > 0) Then 
 
 MsgBox ("This item is involved in a conflict.") 
 
 Else 
 
 MsgBox ("This item is not involved in any conflicts.") 
 
 End If 
 
End Sub
```


## Methods



|Name|
|:-----|
|[GetFirst](Outlook.Conflicts.GetFirst.md)|
|[GetLast](Outlook.Conflicts.GetLast.md)|
|[GetNext](Outlook.Conflicts.GetNext.md)|
|[GetPrevious](Outlook.Conflicts.GetPrevious.md)|
|[Item](Outlook.Conflicts.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.Conflicts.Application.md)|
|[Class](Outlook.Conflicts.Class.md)|
|[Count](Outlook.Conflicts.Count.md)|
|[Parent](Outlook.Conflicts.Parent.md)|
|[Session](Outlook.Conflicts.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]