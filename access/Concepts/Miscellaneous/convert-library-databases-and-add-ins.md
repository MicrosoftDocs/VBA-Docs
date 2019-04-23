---
title: Convert library databases and add-ins
keywords: vbaac10.chm5187103
f1_keywords:
- vbaac10.chm5187103
ms.prod: access
ms.assetid: 676a07f5-4cb2-249e-6e6c-8169700a477a
ms.date: 06/08/2017
localization_priority: Normal
---


# Convert library databases and add-ins

**Applies to:** Access 2013 | Access 2016

If you use add-ins or library databases created in versions of Microsoft Access before 2002, you must convert them to Microsoft Access 2002&ndash;2003 format before you can use them with applications created in Access.

You may also need to make some changes to the objects, macros, and procedures in your library databases and add-ins in order to make sure that they function properly.

## Referencing and loading library databases

Before using a library in Access 2002 or later, you must establish a reference to the library database from each of your applications that uses it. You establish such a reference by clicking **References** on the **Tools** menu while in module Design view. A referenced database must be in Access 2002 format.

A library database should contain only Visual Basic code, which you can call from any application that maintains a reference to that library. In versions 1. _x_ and 2.0 of Access, you load a library database at startup by creating an entry in the Libraries section of your .ini file. Most of the information that's stored in an .ini file in versions 1. _x_ and 2.0 is stored in the Windows registry in later versions. However, there's no need to create a Windows registry key in order to use a library database.


## Circular references between libraries

In versions 1. _x_ and 2.0 of Access, you can make circular library references. However, these aren't allowed in later versions of Access. In other words, once you've created a reference from Library A to Library B, you cannot create a reference from Library B to Library A.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]