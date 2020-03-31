---
title: Guidelines for naming fields, controls, and objects
keywords: vbaac10.chm4348
f1_keywords:
- vbaac10.chm4348
ms.prod: access
ms.assetid: 304d35b1-6f60-445f-b62f-1f0a17b836be
ms.date: 06/08/2019
localization_priority: Normal
---


# Guidelines for naming fields, controls, and objects

**Applies to:** Access 2013 | Access 2016

Names fields, controls, and objects in Microsoft Access:

- Can be up to 64 characters long.
    
- Can include any combination of letters, numbers, spaces, and special characters except a period (.), an exclamation point (!), an accent grave (`), and brackets ([ ]).
    
- Can't begin with leading spaces.
    
- Can't include control characters (ASCII values 0 through 31).
    
Although you can include spaces in field, control, and object names, most examples in the Access documentation show field and control names without spaces because spaces in names can produce naming conflicts in Visual Basic for Applications in some circumstances.

When you name a field, control, or object, it's a good idea to make sure that the name doesn't duplicate the name of a property or other element used by Access; otherwise, your database can produce unexpected behavior in some circumstances. For example, if you refer to the value of a field that is named Name in a table NameInfo by using the syntax NameInfo.Name, Access displays the value of the table's **Name** property rather than the value of the Name field.

Another way to avoid unexpected results is to always use the `!` operator instead of the `.` (dot) operator to refer to the value of a field, control, or object. For example, the following identifier explicitly refers to the value of the Name field rather than the Name property: `[NameInfo]![Name]`.

> [!NOTE] 
> The `!` operator can be used only in Access desktop databases. 

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
