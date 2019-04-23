---
title: AllowEditRange object (Excel)
keywords: vbaxl10.chm725072
f1_keywords:
- vbaxl10.chm725072
ms.prod: excel
api_name:
- Excel.AllowEditRange
ms.assetid: 2bfd80d1-3a59-162e-194a-8699ca6b0d4b
ms.date: 03/29/2019
localization_priority: Normal
---


# AllowEditRange object (Excel)

Represents the cells that can be edited on a protected worksheet.


## Remarks

Use the **Add** method or the **Item** property of the **[AllowEditRanges](Excel.AllowEditRanges.md)** collection to return an **AllowEditRange** object.

After an **AllowEditRange** object has been returned, you can use the **ChangePassword** method to change the password to access a range that can be edited on a protected worksheet.


## Example

In this example, Microsoft Excel allows edits to range "A1:A4" on the active worksheet, notifies the user, and then changes the password for this specified range and notifies the user of this change.

```vb
Sub UseChangePassword() 
 
 Dim wksOne As Worksheet 
 Dim wksPassword As String 
 
 Set wksOne = Application.ActiveSheet 
 
 wksPassword = InputBox ("Enter password for the worksheet") 
 
 ' Establish a range that can allow edits 
 ' on the protected worksheet. 
 wksOne.Protection.AllowEditRanges.Add _ 
 Title:="Classified", _ 
 Range:=Range("A1:A4"), _ 
 Password:=wksPassword 
 
 MsgBox "Cells A1 to A4 can be edited on the protected worksheet." 
 
 ' Change the password. 
 
 wksPassword = InputBox ("Enter the new password for the worksheet") 
 
 wksOne.Protection.AllowEditRanges(1).ChangePassword _ 
 Password:=wksPassword 
 
 MsgBox "The password for these cells has been changed." 
 
End Sub
```


## Methods

- [ChangePassword](Excel.AllowEditRange.ChangePassword.md)
- [Delete](Excel.AllowEditRange.Delete.md)
- [Unprotect](Excel.AllowEditRange.Unprotect.md)

## Properties

- [Range](Excel.AllowEditRange.Range.md)
- [Title](Excel.AllowEditRange.Title.md)
- [Users](Excel.AllowEditRange.Users.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]