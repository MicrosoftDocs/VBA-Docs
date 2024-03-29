---
title: Workbook.AddinInstall event (Excel)
keywords: vbaxl10.chm503080
f1_keywords:
- vbaxl10.chm503080
api_name:
- Excel.Workbook.AddinInstall
ms.assetid: 671117b2-590e-9d6f-29ae-5f0bf30d4e99
ms.date: 05/25/2019
ms.localizationpriority: medium
---


# Workbook.AddinInstall event (Excel)

Occurs when the workbook is installed as an add-in.


## Syntax

_expression_.**AddinInstall**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Return value

**Nothing**


## Example

This example adds a control to the standard toolbar when the workbook is installed as an add-in.

```vb
Private Sub Workbook_AddinInstall() 
 With Application.Commandbars("Standard").Controls.Add 
 .Caption = "The AddIn's menu item" 
 .OnAction = "'ThisAddin.xls'!Amacro" 
 End With End Sub 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]