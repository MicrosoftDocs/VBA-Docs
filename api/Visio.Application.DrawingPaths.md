---
title: Application.DrawingPaths property (Visio)
keywords: vis_sdr.chm10013445
f1_keywords:
- vis_sdr.chm10013445
ms.prod: visio
api_name:
- Visio.Application.DrawingPaths
ms.assetid: 46a0a769-8ef4-1cc3-9206-24c23b0982ee
ms.date: 06/25/2019
localization_priority: Normal
---


# Application.DrawingPaths property (Visio)

Gets or sets the paths where Microsoft Visio looks for drawings. Read/write.


## Syntax

_expression_.**DrawingPaths**

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Return value

String


## Remarks

The **DrawingPaths** property is set to an empty string ("") by default.

The string passed to and received from the **DrawingPaths** property is the same string shown in the **File Locations** dialog box (**File** tab > **Options** > **Advanced** > **General** > **File Locations**). This string is stored in the **HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Visio\Application\DrawingsPath** subkey.

Visio looks for drawings in all paths named in the **DrawingPaths** property and all the subfolders of those paths. If you pass the **DrawingPaths** property to the **EnumDirectories** method, it returns a complete list of fully qualified paths in the folders passed in.

Setting the **DrawingPaths** property replaces existing values for **DrawingPaths** in the **File Locations** dialog box. To retain existing values, get the existing string and then append the new file path to that string, as shown in the following code.

```vb
Application.DrawingPaths = Application.DrawingPaths & ";" & "newpath ".
```

> [!WARNING] 
> Modifying the Windows registry in any manner, whether in the Registry Editor or programmatically, always carries some degree of risk. Incorrect modification can cause serious problems that may require you to reinstall your operating system. It is a good practice to always back up a computer's registry first before modifying it. 

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]