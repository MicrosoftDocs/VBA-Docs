---
title: Application.CutCopyMode property (Excel)
keywords: vbaxl10.chm133101
f1_keywords:
- vbaxl10.chm133101
ms.prod: excel
api_name:
- Excel.Application.CutCopyMode
ms.assetid: d45d3352-2a33-99ae-22f2-0b1c11466209
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.CutCopyMode property (Excel)

Returns or sets the status of Cut or Copy mode. Can be **True**, **False**, or an **[XLCutCopyMode](Excel.XlCutCopyMode.md)** constant, as shown in the following tables. Read/write **Long**.


## Syntax

_expression_.**CutCopyMode**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

|Return value|Description|
|:-----|:-----|
| **False**|Not in Cut or Copy mode.|
| **xlCopy**|In Copy mode.|
| **xlCut**|In Cut mode.|

<br/>

|Set value|Description|
|:-----|:-----|
| **False**|Cancels Cut or Copy mode and removes the moving border.|
| **True**|Starts Cut or Copy mode and shows the moving border.|

## Example

This example uses a message box to display the status of Cut or Copy mode.

```vb
Select Case Application.CutCopyMode 
 Case Is = False 
 MsgBox "Not in Cut or Copy mode" 
 Case Is = xlCopy 
 MsgBox "In Copy mode" 
 Case Is = xlCut 
 MsgBox "In Cut mode" 
End Select
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
