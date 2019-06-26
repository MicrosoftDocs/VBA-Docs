---
title: InvisibleApp.HelpPaths property (Visio)
keywords: vis_sdr.chm17513635
f1_keywords:
- vis_sdr.chm17513635
ms.prod: visio
api_name:
- Visio.InvisibleApp.HelpPaths
ms.assetid: 31e7a73f-85ad-dce0-cfce-3b1a1fdb634d
ms.date: 06/26/2019
localization_priority: Normal
---


# InvisibleApp.HelpPaths property (Visio)

Gets or sets the paths where Microsoft Visio looks for Help files. Read/write.


## Syntax

_expression_.**HelpPaths**

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Return value

String


## Remarks

The **HelpPaths** property is set to an empty string ("") by default.

The string passed to and received from the **HelpPaths** property is the same string shown in the **File Paths** dialog box (**File** tab > **Options** > **Advanced** > **General** > **File Locations**). This string is stored in the **HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Visio\Application\HelpPath** subkey.

When the application looks for Help files, it looks in all paths named in the **HelpPaths** property and all the subfolders of those paths. If you pass the **HelpPaths** property to the **EnumDirectories** method, it returns a complete list of fully qualified paths in the folders passed in.

Setting the **HelpPaths** property replaces existing values for **HelpPaths** in the **File Paths** dialog box. To retain existing values, get the existing string and then append the new file path to that string, as shown in the following code.

```vb
Application.HelpPaths = Application.HelpPaths & ";" & "newpath".
```

> [!WARNING] 
> Modifying the Windows registry in any manner, whether in the Registry Editor or programmatically, always carries some degree of risk. Incorrect modification can cause serious problems that may require you to reinstall your operating system. It is a good practice to always back up a computer's registry first before modifying it. 


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to get and set the **HelpPaths** property of the **Application** object. Before running this macro, replace `fullpath(s)` with the path or paths to the location or locations where you want Visio to look for Help files.

```vb
 
Public Sub GetHelpPaths_Example()  
 
    Dim strCurrentPath As String 
 
    'Retrieve the current path to Visio Help files.  
    strCurrentPath = Application.HelpPaths  
    MsgBox ("The current path for Microsoft Visio Help files is" + strCurrentPath)  
 
End Sub   
 
Public Sub SetHelpPaths_Example()  
 
    Dim strNewPath As String 
 
    'Store the new path.  
    strNewPath = "fullpath(s)"  
 
    'Set the new path in the Application object.  
    Application.HelpPaths = strNewPath  
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]