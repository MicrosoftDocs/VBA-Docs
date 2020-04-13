---
title: The language-specific code page was not specified or could not be found. (Error 3649)
keywords: jeterr40.chm5003649
f1_keywords:
- jeterr40.chm5003649
ms.prod: access
ms.assetid: cdd32ac4-eae4-d0fc-aab1-b6ca18c56024
ms.date: 06/08/2017
localization_priority: Normal
---


# The language-specific code page was not specified or could not be found. (Error 3649)

  

**Applies to:** Access 2013 | Access 2016

You have attempted to open a database that was created with a language that is not installed on your computer. You should determine what language was specified for this database when it was created and then make sure that language is installed on your system. If the database was created with DAO, the language was specified with the locale argument of the **CreateDatabase** method. If the database was created with Microsoft Access, the language was specified with the option "New Database Sort Order" on the **General** tab of the **Options** dialog box, which is available by clicking **Options** on the **Tools** menu.

Languages can be added to your system through the Regional settings of the Control Panel.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]