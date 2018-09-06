---
title: AllowEditRanges Object (Excel)
keywords: vbaxl10.chm724072
f1_keywords:
- vbaxl10.chm724072
ms.prod: excel
api_name:
- Excel.AllowEditRanges
ms.assetid: c08bf170-f982-ecca-c026-df4b907e1dde
ms.date: 06/08/2017
---


# AllowEditRanges Object (Excel)

A collection of all the  **[AllowEditRange](Excel.AllowEditRange.md)** objects that represent the cells that can be edited on a protected worksheet.


## Remarks

Use the  **[AllowEditRanges](Excel.Protection.AllowEditRanges.md)** property of the **[Protection](Excel.Protection.md)** object to return an **AllowEditRanges** collection.

Once an  **AllowEditRanges** collection has been returned, you can use the **[Add](Excel.AllowEditRanges.Add.md)** method to add a range that can be edited on a protected worksheet.


## Example

In this example, Microsoft Excel allows edits to range "A1:A4" on the active worksheet and notifies the user of the title and address of the specified range.


```vb
Sub UseAllowEditRanges() 
 
 Dim wksOne As Worksheet 
 Dim wksPassword As String 
 
 Set wksOne = Application.ActiveSheet 
 
 ' Unprotect worksheet. 
 wksOne.Unprotect 
 
 wksPassword = InputBox ("Enter password for the worksheet") 
 
 ' Establish a range that can allow edits 
 ' on the protected worksheet. 
 wksOne.Protection.AllowEditRanges.Add _ 
 Title:="Classified", _ 
 Range:=Range("A1:A4"), _ 
 Password:=wksPassword 
 
 ' Notify the user 
 ' the title and address of the range. 
 With wksOne.Protection.AllowEditRanges.Item(1) 
 MsgBox "Title of range: " & .Title 
 MsgBox "Address of range: " & .Range.Address 
 End With 
 
End Sub
```


## See also



[Excel Object Model Reference](overview/Excel/object-model.md)

