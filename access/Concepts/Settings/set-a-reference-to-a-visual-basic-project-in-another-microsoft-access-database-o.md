---
title: Set a reference to a Visual Basic project in another Access database or project
keywords: vbaac10.chm102162
f1_keywords:
- vbaac10.chm102162
ms.prod: access
ms.assetid: a919be67-84ee-e9de-1cfd-17a456f4d929
ms.date: 09/26/2018
localization_priority: Normal
---


# Set a reference to a Visual Basic project in another Access database or project

Each Access database includes a Visual Basic project. The Visual Basic project is the set of all modules in the project, including both standard modules and class modules. Every Access database, library database, or add-in contained in an .mde file includes a Visual Basic project.

The name of the Access database and the name of the project can differ. The name of the Access database is determined by the name of the .mdb (or .mda or .mde) or .adp file, while the name of the project is determined by the setting of the [CodeProject.Name property (Access)](../../../api/Access.CodeProject.Name.md) option on the **General** tab of the _ProjectName -_ **Project Properties** dialog box, available by clicking _ProjectName_ Properties on the **Tools** menu in the Visual Basic Editor. When you first create a database (.mdb or .adp), the database name and project name are the same by default. However, if you rename the database, the project name doesn't automatically change. Likewise, changing the project name has no effect on the database name.

You can set a reference from a Visual Basic project in one Access database to a project in another Access database, a library database, or an add-in contained in an .mde file. Once you've set a reference, you can run Visual Basic procedures in the referenced project. For example, the Northwind sample database includes a module named Utility Functions that contains a function called IsLoaded. You can set a reference to the project in the Northwind sample database from the project in the current database, and then call the IsLoaded function just as you would if it were defined within the current database.

> [!NOTE] 
> - Set a reference to the project in another Access database when you want to call a public procedure that's defined within a standard module in that database. You can't call procedures that are defined within a class module or procedures in a standard module that are preceded with the **Private** keyword.
> - You can set a reference to the project in a Access database only from another Access database.
> - You can set a reference to a project only in another Access 2002 or later database. To set a reference to a project in a database created in an earlier version of Access, first convert that database to Access 2002 or later.
> - If you set a reference to a project or type library from Access and then move the file that contains that project or type library to a different folder, Access will attempt to locate the file and reestablish the reference. If the RefLibPaths key exists in the registry, Access will first search there. If there's no matching entry, Access will search for the file first in the current folder, then in all the folders on the drive. You can create the RefLibPaths key by using the Registry Editor in Windows, under the registry key \HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\version\Access. For more information about using the Registry Editor, see your Windows documentation.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]