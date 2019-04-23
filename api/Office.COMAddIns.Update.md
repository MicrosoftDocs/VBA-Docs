---
title: COMAddIns.Update method (Office)
keywords: vbaof11.chm220004
f1_keywords:
- vbaof11.chm220004
ms.prod: office
api_name:
- Office.COMAddIns.Update
ms.assetid: 4cbaff64-10e8-d792-60b5-29f6de97dc8f
ms.date: 01/03/2019
localization_priority: Normal
---


# COMAddIns.Update method (Office)

Updates the contents of the **COMAddIns** collection from the list of add-ins stored in the Windows registry.


## Syntax

_expression_.**Update**

_expression_ A variable that represents a **[COMAddIns](Office.COMAddIns.md)** object.


## Remarks

Before you can use a given COM add-in in a Microsoft Office application, that add-in must be registered in the Windows registry as a COM component with a corresponding Component Category ID. Normally the setup program for a COM add-in will add the necessary entries to the registry.


## Example

The following example updates the contents of the **COMAddIns** collection from the list of add-ins stored in the Windows registry.


```vb
Application.COMAddIns.Update
```


## See also

- [COMAddIns object members](overview/Library-Reference/comaddins-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]