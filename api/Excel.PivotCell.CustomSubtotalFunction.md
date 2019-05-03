---
title: PivotCell.CustomSubtotalFunction property (Excel)
keywords: vbaxl10.chm692082
f1_keywords:
- vbaxl10.chm692082
ms.prod: excel
api_name:
- Excel.PivotCell.CustomSubtotalFunction
ms.assetid: 35c031a2-7ad4-9cbb-c57b-1f529368d307
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotCell.CustomSubtotalFunction property (Excel)

Returns the custom subtotal function field setting of a **PivotCell** object. Read-only **[XlConsolidationFunction](Excel.XlConsolidationFunction.md)**.


## Syntax

_expression_.**CustomSubtotalFunction**

_expression_ A variable that represents a **[PivotCell](Excel.PivotCell.md)** object.


## Remarks

The **CustomSubtotalFunction** property returns an error if the **PivotCell** object type is not a custom subtotal. This property applies only to non-OLAP source data.


## Example

This example determines if cell C20 contains a custom subtotal function that uses a consolidation function of count, and then it notifies the user. The example assumes that a PivotTable exists on the active worksheet.

```vb
Sub UseCustomSubtotalFunction() 
 
 On Error GoTo Not_A_Function 
 
 ' Determine if custom subtotal function is a count function. 
 If Application.Range("C20").PivotCell.CustomSubtotalFunction = xlCount Then 
 MsgBox "The custom subtotal function is a Count." 
 Else 
 MsgBox "The custom subtotal function is not a Count." 
 End If 
 Exit Sub 
 
Not_A_Function: 
 MsgBox "The selected cell is not a custom subtotal function." 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]