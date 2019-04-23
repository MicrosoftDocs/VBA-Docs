---
title: Suspend code execution until a pop-up form is closed
ms.prod: access
ms.assetid: d4d419ac-bf43-3356-4c20-e9bb74f9f591
ms.date: 09/25/2018
localization_priority: Normal
---


# Suspend code execution until a pop-up form is closed

To ensure that code in a form suspends operation until a pop-up form is closed, you must open the pop-up form as a modalwindow. The following example illustrates how to use the **[OpenForm](../../../api/Access.DoCmd.OpenForm.md)** method to do this.


```vb
doCmd.OpenForm FormName:=<Name of form to open>, WindowMode:=acDialog
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
