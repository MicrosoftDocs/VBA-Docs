---
title: Form.UseDefaultPrinter Property (Access)
keywords: vbaac10.chm13559
f1_keywords:
- vbaac10.chm13559
ms.prod: access
api_name:
- Access.Form.UseDefaultPrinter
ms.assetid: bdb7f428-ee00-5a76-e723-6d1858a6172c
ms.date: 06/08/2017
---


# Form.UseDefaultPrinter Property (Access)

Returns or sets a  **Boolean** indicating whether the specified form uses the default printer for the system; **True** if the form or report uses the default printer. Read/write.


## Syntax

 _expression_. **UseDefaultPrinter**

 _expression_ A variable that represents a **Form** object.


## Remarks

This property is read/write in Design view and read-only in all other views.

When this property is  **True**, the form or report inherits all of its printer settings from the settings of the default printer. Changing the printer associated with a form or report by assigning its **Printer** property to a **Printer** object sets the **UseDefaultPrinter** property to **False**.


## Example

The following example checks to see if the specified form is using the default printer; if not, the user is asked if it should.


```vb
Function CheckPrinter(frmTemp As Form) As Boolean 
 
 If frmTemp.UseDefaultPrinter = False Then 
 If MsgBox("Should this form use " _ 
 &; "the default printer?", _ 
 vbYesNo) = vbYes Then 
 frmTemp.UseDefaultPrinter = True 
 End If 
 End If 
 
End Function
```


## See also


#### Concepts


[Form Object](Access.Form.md)

