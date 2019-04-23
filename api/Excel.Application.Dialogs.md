---
title: Application.Dialogs property (Excel)
keywords: vbaxl10.chm133118
f1_keywords:
- vbaxl10.chm133118
ms.prod: excel
api_name:
- Excel.Application.Dialogs
ms.assetid: 0d04aa87-9872-23e5-78e3-c9e3da2c8eb5
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.Dialogs property (Excel)

Returns a **[Dialogs](Excel.Dialogs.md)** collection that represents all built-in dialog boxes. Read-only.


## Syntax

_expression_.**Dialogs**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example displays the **Open** dialog box (**File** menu).

```vb
Application.Dialogs(xlDialogOpen).Show
```

<br/>

The following code example opens an email message in Microsoft Outlook with the current workbook attached.

```vb
Sub SendIt() 
    Application.Dialogs(xlDialogSendMail).Show arg1:="ask@mrexcel.com", arg2:="This goes in the subject line" 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
