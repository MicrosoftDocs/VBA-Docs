---
title: OptionButton.Left Property (Access)
keywords: vbaac10.chm10582
f1_keywords:
- vbaac10.chm10582
ms.prod: access
api_name:
- Access.OptionButton.Left
ms.assetid: bb7f0e55-e08a-a231-ad6c-55ebdd65cf3b
ms.date: 06/08/2017
---


# OptionButton.Left Property (Access)

You can use the  **Left** property to specify an object's location on a form or report. Read/write **Integer**.


## Syntax

 _expression_. **Left**

 _expression_ A variable that represents an **OptionButton** object.


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


[OptionButton Object](Access.OptionButton.md)

