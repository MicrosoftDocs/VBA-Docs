---
title: Application.Hinstance property (Excel)
keywords: vbaxl10.chm133278
f1_keywords:
- vbaxl10.chm133278
ms.prod: excel
api_name:
- Excel.Application.Hinstance
ms.assetid: 4551a0a2-0730-1288-7a13-b2beff2a2fca
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.Hinstance property (Excel)

Returns a handle to the instance of Excel represented by the **Application** object. Read-only **Long**.


## Syntax

_expression_.**Hinstance**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

> [!IMPORTANT] 
> This property returns a correct handle only in the 32-bit version of Excel. In Excel, the **[HinstancePtr](Excel.Application.HinstancePtr.md)** property was introduced, which works correctly in both 32-bit and 64-bit versions of Excel.


## Example

In this example, a message box displays the Excel instance handle to the user.

```vb
Sub CheckHinstance() 
 
 MsgBox Application.Hinstance 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]