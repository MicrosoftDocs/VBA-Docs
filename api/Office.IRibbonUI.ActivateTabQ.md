---
title: IRibbonUI.ActivateTabQ method (Office)
keywords: vbaof11.chm320006
f1_keywords:
- vbaof11.chm320006
ms.prod: office
api_name:
- Office.IRibbonUI.ActivateTabQ
ms.assetid: bf664b52-2660-2ce7-a01b-83b459f66e09
ms.date: 01/16/2019
localization_priority: Normal
---


# IRibbonUI.ActivateTabQ method (Office)

Activates the specified custom tab on the Microsoft Office Fluent Ribbon UI. Uses the fully qualified name of the tab, which includes the ID and the namespace of the tab. 


## Syntax

_expression_.**ActivateTabQ** (_ControlID_, _Namespace_)

_expression_ An expression that returns an **[IRibbonUI](Office.IRibbonUI.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ControlID_|Required|**String**|Specifies the Id of the custom Ribbon tab to be activated.|
| _Namespace_|Required|**String**|Specifies the namespace of the tab element.|

## Return value

Nothing


## Example

The following code activates the qualified tab "test:MyTab". It assumes that you have defined the tab in the Ribbon definition file (customUI.xml) as follows. The subroutine that follows is called from the onLoad attribute of the `<customUI>` element in the customUI.xml file.

- `xmlns:test="testnamespace"`
    
- `<tab idQ="test:MyTab">`
    

```vb
Public myRibbon As IRibbonUI 
 
Sub OnLoad(ByVal control As IRibbonControl) 
 myRibbon.ActivateTabQ "MyTab", "testnamespace" 
End Sub
```


## See also

- [IRibbonUI object members](overview/library-reference/iribbonui-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]