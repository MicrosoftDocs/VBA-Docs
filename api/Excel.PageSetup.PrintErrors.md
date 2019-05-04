---
title: PageSetup.PrintErrors property (Excel)
keywords: vbaxl10.chm473105
f1_keywords:
- vbaxl10.chm473105
ms.prod: excel
api_name:
- Excel.PageSetup.PrintErrors
ms.assetid: 4a864a1e-cbdb-8ef7-536d-d2c5f518f9db
ms.date: 05/03/2019
localization_priority: Normal
---


# PageSetup.PrintErrors property (Excel)

Sets or returns an **[XlPrintErrors](Excel.XlPrintErrors.md)** constant specifying the type of print error displayed. This feature allows users to suppress the display of error values when printing a worksheet. Read/write.


## Syntax

_expression_.**PrintErrors**

_expression_ A variable that represents a **[PageSetup](Excel.PageSetup.md)** object.


## Example

In this example, Microsoft Excel uses a formula that returns an error on the active worksheet. The **PrintErrors** property is set to display dashes. A Print Preview window displays the dashes for the print error. This example assumes that a printer driver has been installed.

```vb
Sub UsePrintErrors() 
 
 Dim wksOne As Worksheet 
 
 Set wksOne = Application.ActiveSheet 
 
 ' Create a formula that returns an error value. 
 Range("A1").Value = 1 
 Range("A2").Value = 0 
 Range("A3").Formula = "=A1/A2" 
 
 ' Change print errors to display dashes. 
 wksOne.PageSetup.PrintErrors = xlPrintErrorsDash 
 
 ' Use the Print Preview window to see the dashes used for print errors. 
 ActiveWindow.SelectedSheets.PrintPreview 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]