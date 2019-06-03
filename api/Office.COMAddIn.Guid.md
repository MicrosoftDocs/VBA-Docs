---
title: COMAddIn.Guid property (Office)
keywords: vbaof11.chm219004
f1_keywords:
- vbaof11.chm219004
ms.prod: office
api_name:
- Office.COMAddIn.Guid
ms.assetid: 1e3218d9-dce7-21e2-55a7-4435ca58bb35
ms.date: 01/02/2019
localization_priority: Normal
---


# COMAddIn.Guid property (Office)

Gets the class identifier (CLSID) for the specified **COMAddIn** object. Read-only.


## Syntax

_expression_.**Guid**

_expression_ A variable that represents a **[COMAddIn](Office.COMAddIn.md)** object.


## Example

The following example displays the ProgID and CLSID for the first COM add-in in the **COMAddIns** collection in a message box.


```vb
MsgBox "My ProgID is " & _ 
 Application.COMAddIns(1).ProgID & _ 
 " and my CLSID is " & _ 
 Application.COMAddIns(1).Guid
```


## See also

- [COMAddIn object members](overview/Library-Reference/comaddin-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]