---
title: ComboBox.Left Property (Access)
keywords: vbaac10.chm11401
f1_keywords:
- vbaac10.chm11401
ms.prod: access
api_name:
- Access.ComboBox.Left
ms.assetid: d6eeae85-bc8f-c56e-4014-d1a95e32d18e
ms.date: 06/08/2017
---


# ComboBox.Left Property (Access)

You can use the  **Left** property to specify an object's location on a form or report. Read/write **Integer**.


## Syntax

 _expression_. **Left**

 _expression_ A variable that represents a **ComboBox** object.


## Remarks

In Visual Basic, use a numeric expression to set the value of this property. Values are expressed in twips

For reports, you can set these properties only by using a macro or event procedure in Visual Basic while the report is in Print Preview or being printed.

For reports, the  **Left** property setting is the amount the current section is offset from the left of the page. This property is expressed in twips. You can use this property to specify how far down the page you want a section to print in the section's **Format** event procedure.


## Example

The following example checks the  **Left** property setting for the current report. If the value is less than the minimum margin setting, the **NextRecord** and **PrintSection** properties are set to **False** (0). The section doesn't advance to the next record, and the next section isn't printed.


```vb
Sub Detail1_Format(Cancel As Integer, FormatCount As Integer) 
 
 Const conLeftMargin = 1880 
 
 ' Don't advance to next record or print next section 
 ' if Left property setting is less than 1880 twips. 
 If Me.Left < conLeftMargin Then 
 Me.NextRecord = False 
 Me.PrintSection = False 
 End If 
 
End Sub
```


## See also


#### Concepts


[ComboBox Object](Access.ComboBox.md)

