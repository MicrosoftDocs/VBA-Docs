---
title: IRibbonUI.ActivateTabMso method (Office)
keywords: vbaof11.chm320005
f1_keywords:
- vbaof11.chm320005
ms.prod: office
api_name:
- Office.IRibbonUI.ActivateTabMso
ms.assetid: 74096b3b-c2a7-0247-f3a1-d5e5dc7286e1
ms.date: 01/16/2019
localization_priority: Normal
---


# IRibbonUI.ActivateTabMso method (Office)

Activates the specified built-in tab.


## Syntax

_expression_.**ActivateTabMso** (_ControlID_)

_expression_ An expression that returns an **[IRibbonUI](Office.IRibbonUI.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ControlID_|Required|**String**|Specifies the Id of the custom Ribbon tab to be activated.|

## Return value

Nothing


## Example

The following code makes a built-in tab as specified by the control ID of the active tab.


```vb
Public myRibbon As IRibbonUI 
 
Sub tabActivate(ByVal control As IRibbonControl) 
 myRibbon.ActivateTabMso (control.ID) 
End Sub
```


## See also

- [IRibbonUI object members](overview/library-reference/iribbonui-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]