---
title: Create a portion in your application's registry tree to manage settings
ROBOTS: INDEX
keywords: acmain11.chm1032167
f1_keywords:
- acmain11.chm1032167
ms.prod: access
ms.assetid: ed696038-e77f-ce01-a139-d100d17212e5
ms.date: 06/08/2017
localization_priority: Normal
---


# Create a portion in your application's registry tree to manage settings

**Applies to**: Access 2013 | Access 2016

To customize the Microsoft Windows Registry settings, you can create a Microsoft Access database engine portion in your application's registry tree to manage the settings for the Microsoft Access database engine. 

The easiest way to accomplish this is to export the existing Microsoft Access database engine key and then import it into your application's tree with the Regedit.exe Export and Import commands. You can then alter the values in your new registry tree. If you have supplied any values in the Engines subfolder, the Microsoft Access database engine loads those settings when the application starts. Any values not entered in your client application's registry tree are loaded from shadow settings.

For your application to load the appropriate portion of the Windows Registry key you must specify the location with the DAO **INIPath** property. Your application must set the **INIPath** property before executing any other DAO code. The scope of this setting is limited to your application and cannot be changed without restarting your application.

> [!NOTE] 
> Although creating a Microsoft Access database engine portion in your application's registry is more flexible than overwriting the Microsoft Access database engine default entries, it requires that you maintain the registry tree. Every time changes are required in the default settings, you will need to edit the Registry.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]