---
title: Overview of Customizing the Ribbon
ms.prod: outlook
ms.assetid: ee49751d-9eae-357c-5fa9-0b2dd4ff0890
ms.date: 06/08/2017
localization_priority: Normal
---


# Overview of Customizing the Ribbon

Similar to other Microsoft Office applications such as Microsoft Word, Microsoft Excel, and Microsoft PowerPoint, Microsoft Outlook uses the Microsoft Office Fluent user interface ribbon in its explorer and inspector windows. In an item inspector, such as an email message in compose mode, Outlook uses the ribbon to expose commands in item-specific elements that make it easy for users to identify the commands they need to complete their tasks.

To customize the ribbon programmatically, Outlook uses ribbon extensibility. Each Outlook add-in can specify a custom user interface in an XML markup file, and then implement the  **[IRibbonExtensibility](../../../api/Office.IRibbonExtensibility.md)** interface. Office calls the **[IRibbonExtensibility.GetCustomUI](../../../api/Office.IRibbonExtensibility.GetCustomUI.md)** method before the **ThisAddin.Startup** method to load ribbon customizations for the explorer ribbon, and calls the **GetCustomUI** method the first time that it displays a particular type of inspector. When it is called, the **GetCustomID** method takes a ribbon ID as an argument and loads the corresponding XML that your add-in associates with that ribbon ID. Consider using a `Switch` statement when you implement the **GetCustomID** method to load the ribbon XML for various ribbons; it is probably the most efficient way to accommodate the variety of ribbons that you might customize.

For a complete listing of ribbon identifiers, see  [Implementing the IRibbonExtensibility Interface](implementing-the-iribbonextensibility-interface.md).

For a detailed discussion of the ribbon and ribbon extensibility, see  [Overview of the Office Fluent Ribbon](../../../Library-Reference/Concepts/overview-of-the-office-fluent-ribbon.md).

## See also


 [Detecting Errors](detecting-errors.md)<br>
 [Updating Earlier Code for CommandBars](updating-earlier-code-for-commandbars.md)<br>
 [Overview of the IRibbonUI Object](overview-of-the-iribbonui-object.md)<br>
 [Overview of the IRibbonControl Object](overview-of-the-iribboncontrol-object.md)<br>
 [Office Fluent User Interface Extensibility for Outlook](office-fluent-user-interface-extensibility-for-outlook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]