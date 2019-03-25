---
title: TabControl.TabFixedWidth property (Access)
keywords: vbaac10.chm12089,vbaac10.chm5506
f1_keywords:
- vbaac10.chm12089,vbaac10.chm5506
ms.prod: access
api_name:
- Access.TabControl.TabFixedWidth
ms.assetid: e4f819b3-fb2d-cddb-c675-6bc831cb2077
ms.date: 03/26/2019
localization_priority: Normal
---


# TabControl.TabFixedWidth property (Access)

You can use the **TabFixedWidth** property to specify or determine the width of the tabs on a tab control. Read/write **Integer**.


## Syntax

_expression_.**TabFixedWidth**

_expression_ A variable that represents a **[TabControl](Access.TabControl.md)** object.


## Remarks

The **TabFixedWidth** property setting is a value that represents the width of tabs in the unit of measurement specified in the **Regional Options** dialog box in the Windows Control Panel. If you set this property to zero, the tabs automatically adjust to the width of the tab contents.

You can also set the default for this property by setting a control's **[DefaultControl](access.form.defaultcontrol.md)** property in Visual Basic.

This property uses an **Integer** value representing the width of the tabs in [twips](../language/glossary/vbe-glossary.md#twip) and can be set in any view.

> [!NOTE] 
> To use a unit of measurement different from the setting in the **Regional Options** dialog box in the Windows Control Panel, specify the unit, such as cm or in (for example, 5 cm or 3 in).

You can't change the color of a tab control. If the tabs don't cover the width of the tab control, the area behind the tabs is displayed. If you place a tab control on an object with a different color than the tab control, you should make sure that the tabs cover the control's background area.


## Example

The following example sets the width of each tab in the tab control **TabCtl1** on the **Mailing List** form to 2000 twips.

```vb
Forms("Mailing List").Controls("TabCtl1").TabFixedWidth = 2000
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]