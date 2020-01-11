---
title: Set startup properties from Visual Basic
ROBOTS: INDEX
keywords: vbaac10.chm14060
f1_keywords:
- vbaac10.chm14060
ms.prod: access
ms.assetid: 73a639d8-38db-cee3-5e16-0d6e1fb54358
ms.date: 06/08/2017
localization_priority: Normal
---


# Set startup properties from Visual Basic

**Applies to:** Access 2013 | Access 2016

In a Microsoft Access database, startup properties are properties of a **Database** object. A **Database** object is a DAO object supplied by the Microsoft Access database engine, but startup properties are defined by Microsoft Access, so they aren't automatically recognized by the Access database engine. If a startup property hasn't been set previously, you must create it and add it to the **Properties** collection of the **Database** object.

In a Microsoft Access project (.adp), startup properties are properties of a **[CurrentProject](https://msdn.microsoft.com/library/e6baae73-1eeb-b48f-d35e-b3e921378561%28Office.15%29.aspx)** object and like the **Database** object in an Access database, startup properties aren't automatically recognized by the Access database engine. If a startup property hasn't been set previously, you must create it and add it to the **[AccessObjectProperties](https://msdn.microsoft.com/library/2df86891-6038-d147-2a32-f1c77b841067%28Office.15%29.aspx)** collection of the **CurrentProject** object.

When you set startup properties from Visual Basic, you should include error-handling code to verify that the property exists in the **Properties** or **AccessObjectProperties** collection. For more information about setting properties defined by Microsoft Access, see [Set Properties of Data Access Objects in Visual Basic](../Data-Access-Objects/set-properties-of-data-access-objects-in-visual-basic.md) or [Set Properties of ActiveX Data Objects in Visual Basic](../ActiveX-Data-Objects/set-properties-of-activex-data-objects-in-visual-basic.md).

The names of the startup properties differ from the text that appears in the **Startup** dialog box. The following table provides the name of each startup property as it is used in Visual Basic code.


|Text in Startup dialog box|Property name|
|:-----|:-----|
|Application Title|**[AppTitle](../../../api/Access.Application.AppTitle.md)**|
|Application Icon|**[AppIcon](../../../api/Access.Application.AppIcon.md)**|
|Display Form/Page|**StartupForm**|
|Display Database Window|**StartupShowDBWindow**|
|Display Status Bar|**StartupShowStatusBar**|
|Menu Bar|**StartupMenuBar**|
|Shortcut Menu Bar|**StartupShortcutMenuBar**|
|Allow Full Menus|**AllowFullMenus**|
|Allow Default Shortcut Menus|**AllowShortcutMenus**|
|Allow Built-In Toolbars|**AllowBuiltInToolbars**|
|Allow Toolbar/Menu Changes|**AllowToolbarChanges**|
|Allow Viewing Code After Error|**AllowBreakIntoCode**|
|Use Access Special Keys|**AllowSpecialKeys**|

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]