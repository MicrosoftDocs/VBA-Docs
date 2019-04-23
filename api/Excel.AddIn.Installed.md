---
title: AddIn.Installed property (Excel)
keywords: vbaxl10.chm185076
f1_keywords:
- vbaxl10.chm185076
ms.prod: excel
api_name:
- Excel.AddIn.Installed
ms.assetid: f8e6e45a-9f6c-2156-dd6f-d3f8e221c282
ms.date: 04/03/2019
localization_priority: Normal
---


# AddIn.Installed property (Excel)

**True** if the add-in is installed or to install the add-in; **False** if the add-in is uninstalled or to uninstall the add-in. Read/write **Boolean**.


## Syntax

_expression_.**Installed**

_expression_ A variable that represents an **[AddIn](Excel.AddIn.md)** object.


## Remarks

Setting this property to **True** installs the add-in and calls its Auto_Add functions. Setting this property to **False** removes the add-in and calls its Auto_Remove functions.


## Example

This example uses a message box to display the installation status of the Solver add-in.

```vb
Set a = AddIns("Solver Add-In") 
If a.Installed = True Then 
 MsgBox "The Solver add-in is installed" 
Else 
 MsgBox "The Solver add-in is not installed" 
End If
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]