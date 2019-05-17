---
title: UserAccess object (Excel)
keywords: vbaxl10.chm727072
f1_keywords:
- vbaxl10.chm727072
ms.prod: excel
api_name:
- Excel.UserAccess
ms.assetid: 44df1865-a5f9-e1b7-b724-41d375e9ea44
ms.date: 04/03/2019
localization_priority: Normal
---


# UserAccess object (Excel)

Represents the user access for a protected range.


## Example

Use the **[Add](Excel.UserAccessList.Add.md)** method or the **[Item](Excel.UserAccessList.Item.md)** property of the **UserAccessList** collection to return a **UserAccess** object.

After a **UserAccess** object is returned, you can determine if access is allowed for a particular range on a worksheet by using the **AllowEdit** property. 

The following example adds a range that can be edited on a protected worksheet and notifies the user of the title of that range.

```vb
Sub UseAllowEditRanges() 
 
 Dim wksSheet As Worksheet 
 
 Set wksSheet = Application.ActiveSheet 
 
 ' Add a range that can be edited on the protected worksheet. 
 wksSheet.Protection.AllowEditRanges.Add "Test", Range("A1") 
 
 ' Notify the user the title of the range that can be edited. 
 MsgBox wksSheet.Protection.AllowEditRanges(1).Title 
 
End Sub
```

## Methods

- [Delete](Excel.UserAccess.Delete.md)

## Properties

- [AllowEdit](Excel.UserAccess.AllowEdit.md)
- [Name](Excel.UserAccess.Name.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]