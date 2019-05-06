---
title: PivotValueCell.Creator property (Excel)
keywords: vbaxl10.chm917074
f1_keywords:
- vbaxl10.chm917074
ms.prod: excel
ms.assetid: 85b4c0bf-3654-af39-413e-8c22c00626f3
ms.date: 05/07/2019
localization_priority: Normal
---


# PivotValueCell.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[PivotValueCell](Excel.pivotvaluecell.md)** object.


## Remarks

Because the object was created in Microsoft Excel, this property returns the hexadecimal value, 5843454C, which represents the string XCEL.


## Property value

**XLCREATOR**


## Example

The following code uses the **Creator** property to check whether the specified object is an Excel object.

```vb
Sub FindCreator() 
 Dim myObject As Excel.Workbook 
 Set myObject = ActiveWorkbook 
 If myObject.Creator = &h5843454c Then 
 MsgBox "This is a Microsoft Excel object." 
 Else 
 MsgBox "This is not a Microsoft Excel object." 
 End If 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]