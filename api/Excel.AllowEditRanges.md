---
title: AllowEditRanges object (Excel)
keywords: vbaxl10.chm724072
f1_keywords:
- vbaxl10.chm724072
ms.prod: excel
api_name:
- Excel.AllowEditRanges
ms.assetid: c08bf170-f982-ecca-c026-df4b907e1dde
ms.date: 03/29/2019
localization_priority: Normal
---


# AllowEditRanges object (Excel)

A collection of all the **[AllowEditRange](Excel.AllowEditRange.md)** objects that represent the cells that can be edited on a protected worksheet.


## Remarks

Use the **[AllowEditRanges](Excel.Protection.AllowEditRanges.md)** property of the **Protection** object to return an **AllowEditRanges** collection.

After an **AllowEditRanges** collection has been returned, you can use the **Add** method to add a range that can be edited on a protected worksheet.


## Example

In this example, Microsoft Excel allows edits to range A1:A4 on the active worksheet, and then notifies the user of the title and address of the specified range.

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


## Methods

- [Add](Excel.AllowEditRanges.Add.md)

## Properties

- [Count](Excel.AllowEditRanges.Count.md)
- [Item](Excel.AllowEditRanges.Item.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]