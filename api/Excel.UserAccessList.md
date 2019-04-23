---
title: UserAccessList object (Excel)
keywords: vbaxl10.chm726072
f1_keywords:
- vbaxl10.chm726072
ms.prod: excel
api_name:
- Excel.UserAccessList
ms.assetid: 8b753ffc-e4d5-0824-e465-a3bdb9ed9202
ms.date: 04/03/2019
localization_priority: Normal
---


# UserAccessList object (Excel)

A collection of **[UserAccess](Excel.UserAccess.md)** objects that represents the user access for protected ranges.


## Example

Use the **[Users](Excel.AllowEditRange.Users.md)** property of the protected **AllowEditRange** object to return a **UserAccessList** collection.

After a **UserAccessList** collection is returned, you can use the **Count** property to determine the number of users that have access to a protected range. 

In the following example, Microsoft Excel notifies the user of the number of users that have access to the first protected range. This example assumes that a protected range exists on the active worksheet.

```vb
Sub UseDeleteAll() 
 
 Dim wksSheet As Worksheet 
 
 Set wksSheet = Application.ActiveSheet 
 
 ' Notify the user of the number of users that can access the protected range. 
 MsgBox wksSheet.Protection.AllowEditRanges(1).Users.Count 
 
End Sub
```

## Methods

- [Add](Excel.UserAccessList.Add.md)
- [DeleteAll](Excel.UserAccessList.DeleteAll.md)

## Properties

- [Count](Excel.UserAccessList.Count.md)
- [Item](Excel.UserAccessList.Item.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]