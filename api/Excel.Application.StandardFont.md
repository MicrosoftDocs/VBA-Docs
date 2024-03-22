---
title: Application.StandardFont property (Excel)
keywords: vbaxl10.chm133210
f1_keywords:
- vbaxl10.chm133210
api_name:
- Excel.Application.StandardFont
ms.assetid: 6bde5ec0-8868-fa00-52e3-b7387f39f56d
ms.date: 04/05/2019
ms.localizationpriority: medium
---


# Application.StandardFont property (Excel)

Returns or sets the name of the standard font. Read/write **String**.


## Syntax

_expression_.**StandardFont**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

If you change the standard font by using this property, the change doesn't take effect until you restart Microsoft Excel.


## Example

This example sets the standard font to Geneva (on the Macintosh) or Arial (in Windows).

```vb
If Application.OperatingSystem Like "*Macintosh*" Then 
 Application.StandardFont = "Geneva" 
Else 
 Application.StandardFont = "Arial" 
End If
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]