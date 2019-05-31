---
title: AddIns object (PowerPoint)
keywords: vbapp10.chm520000
f1_keywords:
- vbapp10.chm520000
ms.prod: powerpoint
api_name:
- PowerPoint.AddIns
ms.assetid: 8308fd95-a220-469e-c33b-cc46ad1b27b8
ms.date: 06/08/2017
localization_priority: Normal
---


# AddIns object (PowerPoint)

A collection of  **[AddIn](PowerPoint.AddIn.md)** objects that represent all the Microsoft PowerPoint-specific add-ins available to PowerPoint, regardless of whether or not they are loaded. This does not include Component Object Model (COM) add-ins.


## Example

Use the  **AddIns** method to return the **AddIns** collection. The following example displays the names of all the add-ins that are currently loaded in PowerPoint.


```vb
For Each ad In AddIns

    If ad.Loaded Then MsgBox ad.Name

Next
```

Use the  **[Add](PowerPoint.AddIns.Add.md)** method to add a PowerPoint-specific add-in to the list of those available. The **Add** method adds an add-in to the list but does not load the add-in. To load the add-in, set the [Loaded](PowerPoint.AddIn.Loaded.md)property of the add-in to  **True** after you use the **Add** method. You can perform these two actions in a single step, as shown in the following example (note that you use the name of the add-in, not its title, with the **Add** method).




```vb
AddIns.Add("graphdrs.ppa").Loaded = True
```

Use  **AddIns** (_index_), where _index_ is the add-in's title or index number, to return a single **AddIn** object. The following example loads the hypothetical add-in titled "my ppt tools".




```vb
AddIns("my ppt tools").Loaded = True
```

Do not confuse the add-in title with the add-in name, which is the file name of the add-in. You must spell the add-in title exactly as it is spelled in the  **Add-Ins** tab, but the capitalization does not have to match.


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]