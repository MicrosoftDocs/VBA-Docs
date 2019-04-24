---
title: DefaultWebOptions.TargetBrowser property (Excel)
keywords: vbaxl10.chm660090
f1_keywords:
- vbaxl10.chm660090
ms.prod: excel
api_name:
- Excel.DefaultWebOptions.TargetBrowser
ms.assetid: 785efc30-ef17-1745-874d-a3be861d450b
ms.date: 04/25/2019
localization_priority: Normal
---


# DefaultWebOptions.TargetBrowser property (Excel)

Returns or sets an **[MsoTargetBrowser](Office.MsoTargetBrowser.md)** constant indicating the browser version. Read/write.


## Syntax

_expression_.**TargetBrowser**

_expression_ A variable that represents a **[DefaultWebOptions](Excel.DefaultWebOptions.md)** object.


## Example

In this example, Microsoft Excel determines if the browser version for web options is Internet Explorer 5 and notifies the user.

```vb
Sub CheckWebOptions() 
 
    Dim wkbOne As Workbook 
 
    Set wkbOne = Application.Workbooks(1) 
 
    ' Determine if IE5 is the target browser. 
    If wkbOne.WebOptions.TargetBrowser = msoTargetBrowserIE5 Then 
        MsgBox "The target browser is IE5 or later." 
    Else 
        MsgBox "The target browser is not IE5 or later." 
    End If 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]