---
title: IRibbonUI.ActivateTab method (Office)
keywords: vbaof11.chm320004
f1_keywords:
- vbaof11.chm320004
ms.prod: office
api_name:
- Office.IRibbonUI.ActivateTab
ms.assetid: 32f5205c-6ab1-e3a6-6bae-5f36706c4d0d
ms.date: 01/16/2019
localization_priority: Normal
---


# IRibbonUI.ActivateTab method (Office)

Activates the specified custom tab. This method returns S_FALSE if there is no Ribbon or the Ribbon is collapsed.


## Syntax

_expression_.**ActivateTab** (_ControlID_)

_expression_ An expression that returns an **[IRibbonUI](Office.IRibbonUI.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ControlID_|Required|**String**|Specifies the Id of the custom Ribbon tab to be activated.|

## Return value

Nothing


## Example

The following code makes the custom tab the active tab.

```vb
Public myRibbon As IRibbonUI 
 
Sub tabActivate(ByVal control As IRibbonControl) 
 myRibbon.ActivateTab (control.ID) 
End Sub
```


## See also

- [IRibbonUI object members](overview/library-reference/iribbonui-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]