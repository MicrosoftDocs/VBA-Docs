---
title: DefaultWebOptions.SaveNewWebPagesAsWebArchives property (Excel)
keywords: vbaxl10.chm660091
f1_keywords:
- vbaxl10.chm660091
ms.prod: excel
api_name:
- Excel.DefaultWebOptions.SaveNewWebPagesAsWebArchives
ms.assetid: 659d338e-74b8-8959-d02b-4d7a08cadbf0
ms.date: 06/08/2017
localization_priority: Normal
---


# DefaultWebOptions.SaveNewWebPagesAsWebArchives property (Excel)

 **True** if new Web pages can be saved as Web archives. Read/write **Boolean**.


## Syntax

_expression_. `SaveNewWebPagesAsWebArchives`

_expression_ A variable that represents a [DefaultWebOptions](Excel.DefaultWebOptions.md) object.


## Example

In this example, Microsoft Excel determines the settings for saving new Web pages as Web archives and notifies the user.


```vb
Sub DetermineSettings() 
 
 ' Determine settings and notify user. 
 If Application.DefaultWebOptions.SaveNewWebPagesAsWebArchives = True Then 
 MsgBox "New Web pages will be saved as Web archives." 
 Else 
 MsgBox "New Web pages will not be saved as Web archives." 
 End If 
 
End Sub
```


## See also


[DefaultWebOptions Object](Excel.DefaultWebOptions.md)

