---
title: ProtectedViewWindow.SourceName property (Excel)
keywords: vbaxl10.chm914081
f1_keywords:
- vbaxl10.chm914081
ms.prod: excel
api_name:
- Excel.ProtectedViewWindow.SourceName
ms.assetid: e5347e6e-b9d4-d3b1-ca41-ba577d836e31
ms.date: 05/09/2019
localization_priority: Normal
---


# ProtectedViewWindow.SourceName property (Excel)

Returns the name of the source file that is open in the specified Protected View window. Read-only.


## Syntax

_expression_.**SourceName**

_expression_ A variable that represents a **[ProtectedViewWindow](Excel.ProtectedViewWindow.md)** object.


## Return value

**String**


## Remarks

This property does not return the path for the source file. To return the path, use the **[SourcePath](Excel.ProtectedViewWindow.SourcePath.md)** property of the **ProtectedViewWindow** object.


## Example

The following example returns the path and name of the workbook associated with the specified Protected View window.

```vb
MsgBox ActiveProtectedViewWindow.SourcePath & "\" _ 
 & ActiveProtectedViewWindow.SourceName
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]