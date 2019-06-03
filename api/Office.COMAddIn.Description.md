---
title: COMAddIn.Description property (Office)
keywords: vbaof11.chm219001
f1_keywords:
- vbaof11.chm219001
ms.prod: office
api_name:
- Office.COMAddIn.Description
ms.assetid: f194ae48-0762-732f-7c9a-f19a92e94d9b
ms.date: 01/02/2019
localization_priority: Normal
---


# COMAddIn.Description property (Office)

Gets or sets a descriptive **String** value for the specified **COMAddin** object. Read/write.


## Syntax

_expression_.**Description**

_expression_ Required. A variable that represents a **[COMAddIn](Office.COMAddIn.md)** object.


## Example

The following example displays the description text of the Microsoft Accessibility COM add-in for drawing.


```vb
MsgBox "The description of this " & _ 
 "COMAddIn is """ & Application.COMAddIns. _ 
 Item("msodraa9.ShapeSelect"). _ 
 Description & """
```


## See also

- [COMAddIn object members](overview/Library-Reference/comaddin-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]