---
title: UserAccessList object (Excel)
keywords: vbaxl10.chm726072
f1_keywords:
- vbaxl10.chm726072
ms.prod: excel
api_name:
- Excel.UserAccessList
ms.assetid: 8b753ffc-e4d5-0824-e465-a3bdb9ed9202
ms.date: 06/08/2017
---


# UserAccessList object (Excel)

A collection of  **[UserAccess](Excel.UserAccess.md)** objects that represent the user access for protected ranges.


## Example

Use the  **[Users](Excel.AllowEditRange.Users.md)** property of the protected **[Range](Excel.Range(object).md)** object to return a **UserAccessList** collection.



Once a  **UserAccessList** collection is returned you can use the **[Count](Excel.UserAccessList.Count.md)** property to determine the number of users that have access to a protected range. In the following example, Microsoft Excel notifies the user the numbers users that have access to the first protected range. This example assumes that a protected range exists on the active worksheet.






```vb
Sub UseDeleteAll() 
 
 Dim wksSheet As Worksheet 
 
 Set wksSheet = Application.ActiveSheet 
 
 ' Notify the user the number of users that can access the protected range. 
 MsgBox wksSheet.Protection.AllowEditRanges(1).Users.Count 
 
End Sub
```


## See also



[Excel Object Model Reference](./overview/Excel/object-model.md)

