---
title: OLEObject.GotFocus event (Excel)
keywords: vbaxl10.chm501073
f1_keywords:
- vbaxl10.chm501073
api_name:
- Excel.OLEObject.GotFocus
ms.assetid: 2bd9a3d8-9305-2354-5ddd-262f4720b444
ms.date: 05/02/2019
ms.localizationpriority: medium
---


# OLEObject.GotFocus event (Excel)

Occurs when an ActiveX control gets input focus.


## Syntax

_expression_.**GotFocus**

_expression_ A variable that represents an **[OLEObject](Excel.OLEObject.md)** object.


## Return value

Nothing


## Example

This example runs when ListBox1 gets the focus.

```vb
Private Sub ListBox1_GotFocus() 
 ' runs when list box gets the focus 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]