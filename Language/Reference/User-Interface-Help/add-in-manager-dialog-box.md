---
title: Add-In Manager dialog box
keywords: vbui6.chm181033
f1_keywords:
- vbui6.chm181033
ms.prod: office
ms.assetid: c19b1493-5c13-bcc3-6b45-136d7313ded5
ms.date: 11/26/2018
localization_priority: Normal
---


# Add-In Manager dialog box

![Add-in manager](../../../images/va5lxy1_ZA01201779.gif)

Allows you to register an [add-in](../../Glossary/vbe-glossary.md#add-in), load or unload it, and set its load behavior. If you close only the visible portions of an add-in (by double-clicking its system menu or by clicking its **Close** button, for example), its forms disappear from the screen, but the add-in is still present in memory. The **[Add-in](../visual-basic-add-in-model/objects-visual-basic-add-in-model.md#addin)** object itself will always stay resident in memory until the add-in is disconnected through the **Add-In Manager** dialog box.

To open the Add-In Manager, select **Add-In Manager** from the **[Add-Ins](add-ins-menu.md)** menu.

The following table describes the dialog box options.

|Option|Description|
|:-----|:----------|
|**Available Add-Ins** (Add-Ins list)|Lists available add-ins.|
|**Load Behavior** (Add-Ins list)|Displays the load behavior for the selected add-in.|
|**Description**|Displays a description of what the add-in does.|
|**Load Behavior**|**Loaded/Unloaded** check box: Loads or unloads the selected add-in.<br/><br/>**Load On Startup** check box: Loads the selected add-in on startup of the development environment.<br/><br/>**Command Line** check box: Loads the selected add-in when the development environment is started from the command prompt or from a script.|
|**OK**|Updates the load behavior of selected add-ins.|
|**Cancel**|Cancels all updates made in session.|


## See also

- [Using the Add-In Manager](../../concepts/getting-started/using-the-add-in-manager.md)
- [Dialog boxes](../dialog-boxes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
