---
title: Form.HelpFile property (Access)
keywords: vbaac10.chm13393
f1_keywords:
- vbaac10.chm13393
ms.prod: access
api_name:
- Access.Form.HelpFile
ms.assetid: 72b416b1-5257-9560-ebc0-625abc3f7e85
ms.date: 03/12/2019
localization_priority: Normal
---


# Form.HelpFile property (Access)

The name of a help file associated with a form. Read/write **String**.


## Syntax

_expression_.**HelpFile**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Example

This example adds a custom command bar with a combo box that tracks stock data. The example also specifies the Help topic to be displayed for the combo box when the user presses Shift+F1.

```vb
Set myBar = CommandBars _ 
 .Add(Name:="Custom", Position:=msoBarTop, _ 
 Temporary:=True) 
With myBar 
 .Controls.Add Type:=msoControlComboBox, ID:=1 
 .Visible = True 
End With 
With CommandBars("Custom").Controls(1) 
 .AddItem "Get Stock Quote", 1 
 .AddItem "View Chart", 2 
 .AddItem "View Fundamentals", 3 
 .AddItem "View News", 4 
 .Caption = "Stock Data" 
 .DescriptionText = "View Data For Stock" 
 .HelpFile = "C:\corphelp\custom.hlp" 
 .HelpContextID = 47 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]