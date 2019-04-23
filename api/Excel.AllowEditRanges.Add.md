---
title: AllowEditRanges.Add method (Excel)
keywords: vbaxl10.chm724075
f1_keywords:
- vbaxl10.chm724075
ms.prod: excel
api_name:
- Excel.AllowEditRanges.Add
ms.assetid: f88d900d-4974-4d8d-6279-0be6376fc232
ms.date: 04/04/2019
localization_priority: Normal
---


# AllowEditRanges.Add method (Excel)

Adds a range that can be edited on a protected worksheet. Returns an **[AllowEditRange](Excel.AllowEditRange.md)** object.


## Syntax

_expression_.**Add** (_Title_, _Range_, _Password_)

_expression_ A variable that represents an **[AllowEditRanges](Excel.AllowEditRanges.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Title_|Required| **String**|The title of the range.|
| _Range_|Required| **Range**| **[Range](Excel.Range(object).md)** object. The range allowed to be edited.|
| _Password_|Optional| **Variant**|The password for the range.|

## Return value

An **AllowEditRange** object that represents the range.


## Example

This example allows edits to range A1:A4 on the active worksheet, notifies the user, changes the password for this specified range, and then notifies the user of this change.


```vb
Sub UseChangePassword() 
 
 Dim wksOne As Worksheet 
 
 Set wksOne = Application.ActiveSheet 
 
 ' Protect the worksheet. 
 wksOne.Protect 
 
 ' Establish a range that can allow edits 
 ' on the protected worksheet. 
 wksOne.Protection.AllowEditRanges.Add _ 
 Title:="Classified", _ 
 Range:=Range("A1:A4"), _ 
 Password:="secret" 
 
 MsgBox "Cells A1 to A4 can be edited on the protected worksheet." 
 
 ' Change the password. 
 wksOne.Protection.AllowEditRanges(1).ChangePassword _ 
 Password:="moresecret" 
 
 MsgBox "The password for these cells has been changed." 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]