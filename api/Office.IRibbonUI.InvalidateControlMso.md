---
title: IRibbonUI.InvalidateControlMso method (Office)
keywords: vbaof11.chm320003
f1_keywords:
- vbaof11.chm320003
ms.prod: office
api_name:
- Office.IRibbonUI.InvalidateControlMso
ms.assetid: bfcca0e9-8696-6a0e-ff27-6dfde41dff93
ms.date: 01/16/2019
localization_priority: Normal
---


# IRibbonUI.InvalidateControlMso method (Office)

Used to invalidate a built-in control.


## Syntax

_expression_.**InvalidateControlMso** (_ControlID_)

_expression_ An expression that returns an **[IRibbonUI](Office.IRibbonUI.md)** object.


## Parameters

|Name|Required/Optional|Data type|
|:-----|:-----|:-----|
| _ControlID_|Required|**String**|

## Return value

Nothing


## Remarks

Invalidating a control repaints the screen and causes any callback procedures associated with that control to execute.


## Example


```xml
<customUI … OnLoad="MyAddInInitialize" …>
```


```vb
Sub MyAddInInitialize(Ribbon As IRibbonUI) 
 Set MyRibbon = Ribbon 
End Sub 
 
Sub myFunction() 
 MyRibbon.InvalidateControlMso("TabInsert") ' Invalidates the Insert control 
End Sub
```


## See also

- [IRibbonUI object members](overview/library-reference/iribbonui-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]