---
title: Application.MapPaperSize property (Excel)
keywords: vbaxl10.chm133286
f1_keywords:
- vbaxl10.chm133286
api_name:
- Excel.Application.MapPaperSize
ms.assetid: c1d83fab-6abc-9103-78cf-047a503688b1
ms.date: 04/05/2019
ms.localizationpriority: medium
---


# Application.MapPaperSize property (Excel)

**True** if documents formatted for the standard paper size of another country/region (for example, A4) are automatically adjusted so that they're printed correctly on the standard paper size (for example, Letter) of your country/region. Read/write **Boolean**.


## Syntax

_expression_.**MapPaperSize**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example determines if Microsoft Excel can adjust the paper size according to the country/region setting, and then notifies the user.

```vb
Sub UseMapPaperSize() 
 
 ' Determine setting and notify user. 
 If Application.MapPaperSize = True Then 
 MsgBox "Microsoft Excel automatically " & _ 
 "adjusts the paper size according to the country/region setting." 
 Else 
 MsgBox "Microsoft Excel does not " & _ 
 "automatically adjusts the paper size according to the country/region setting." 
 End If 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]