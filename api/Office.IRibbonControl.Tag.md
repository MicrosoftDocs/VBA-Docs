---
title: IRibbonControl.Tag property (Office)
keywords: vbaof11.chm288003
f1_keywords:
- vbaof11.chm288003
ms.prod: office
api_name:
- Office.IRibbonControl.Tag
ms.assetid: d0f041c
localization_priority: Normal
---


# IRibbonControl.Tag property (Office)

Used to store arbitrary strings and fetch them at runtime. Read-only.


## Syntax

_expression_.**Tag**

_expression_ An expression that returns an **[IRibbonControl](Office.IRibbonControl.md)** object.


## Return value

String


## Remarks

Normally you can distinguish between controls in a Ribbon user interface XML customization file by using the **Id** property. However, there are restrictions on what IDs can contain (no non-alphanumeric characters, and they must all be unique). The **Tag** property doesn't have these restrictions, so it can be used in the following situations, where ID doesn't work:

- If you need to store a special string with your control such as a filename. For example: tag="C:\path\file.xlsm."
    
- If you want multiple controls to be treated the same way by your callback procedures, but you don't want to maintain a list of all their IDs (which must be unique). For example, you could have buttons on different tabs on the Ribbon, all with tag="blue", and then just choose the **Tag** property instead of the **ID** property when perfroming some common actions.
    
## Example

In the XML used to customize the Ribbon user interface, you can set a tag as follows. When the MyFunction action is called, you can read the **Tag** property, which will be equal to "some string".


```xml
<button id="mybutton" tag="some string" onAction="MyFunction"/>
```


## See also

- [IRibbonControl object members](overview/library-reference/iribboncontrol-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]