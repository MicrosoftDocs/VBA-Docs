---
title: Office for Mac for Visual Basic for Applications (VBA)
description: Use VBA add-ins and macros that you developed for Office for Windows with Office for Mac.
ms.date: 08/24/2018
localization_priority: Normal
---

# Office for Mac

Use VBA add-ins and macros that you developed for Office for Windows with Office for Mac.

**Applies to:** Excel for Mac | PowerPoint for Mac | Word for Mac | Office 2016 for Mac

If you are authoring Macros for Office for Mac, you can use most of the same objects that are available in VBA for Office. For information about VBA for Excel, PowerPoint, and Word, see the following:

- [Excel VBA reference](https://msdn.microsoft.com/library/ee861528.aspx)
- [PowerPoint VBA reference](https://msdn.microsoft.com/library/ee861525.aspx)
- [Word VBA reference](https://msdn.microsoft.com/library/ee861527.aspx)

> [!NOTE] 
> Outlook for Mac and OneNote for Mac do not support VBA. 

## Office 2016 for Mac is sandboxed

Unlike other versions of Office apps that support VBA, Office 2016 for Mac apps are sandboxed.

Sandboxing restricts the apps from accessing resources outside the app container. This affects any add-ins or macros that involve file access or communication across processes. You can minimize the effects of sandboxing by using the new commands described in the following section.

## Creating an installer or putting user content

For instructions on creating an installer for your add-in, please refer to the article here: [nstalling User Content in Office 2016 for Mac](https://macadmins.software/docs/UserContentIn2016.pdf) 

## New VBA commands for Office 2016 for Mac

The following VBA commands are new and unique to Office 2016 for Mac.

|Command|Use to|
|:-----|:-----|
|[GrantAccessToMultipleFiles](../../Office-Mac/GrantAccessToMultipleFiles.md)|Request a user's permission to access multiple files at once.|
|[AppleScriptTask](../../Office-Mac/AppleScriptTask.md)|Call external AppleScript scripts from VB.|
|[MAC_OFFICE_VERSION](../../Office-Mac/MacOfficeVersion.md)|IFDEF between different Mac Office versions at compile time.|

## Ribbon customization in Office for Mac

Office 2016 for Mac supports ribbon customization using Ribbon XML. Note that there are some differences in ribbon support in Office 2016 for Mac and Office for Windows.

|Ribbon customization feature|Office for Windows|Office for Mac|
|:-----|:-----|:-----|
|Ability to customize the ribbon using Ribbon XML|Available|Available|
|Support for document based add-ins|Available|Available|
|Ability to invoke Macros using custom ribbon controls|Available|Available|
|Customization of custom menus|Available|Available|
|Ability to include and invoke Office Fluent Controls within a custom ribbon tab|Available|Most familiar Office Fluent Control Identifiers are compatible with Office for Mac. Some might not be available. For commands that are compatible with Office 2016 for Mac, see [idMSOs compatible with Office 2016 for Mac](#idMSOs-compatible-with-Office-2016-for-Mac).|
|Support for COM add-ins that use custom ribbon controls|Available|Office 2016 for Mac doesn't support third-party COM add-ins.| 

<a name="idMSOs-compatible-with-Office-2016-for-Mac"></a>

## idMSOs compatible with Office 2016 for Mac

For information about the idMSOs that are compatible with Office 2016 for Mac, see the following:

- [idMSOs supported in Excel for Mac](../../Office-Mac/idMSOExcelMac.md)
- [idMSOs supported in PowerPoint for Mac](../../Office-Mac/idMSOPowerPointMac.md)
- [idMSOs supported in Word for Mac](../../Office-Mac/idMSOWordMac.md)

## See also

- [Announcing add-in support for Gmail accounts in Mac Outlook](https://developer.microsoft.com/office/blogs/announcing-add-in-support-for-gmail-accounts-in-mac-outlook)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
