---
title: AddIn object (PowerPoint)
keywords: vbapp10.chm521000
f1_keywords:
- vbapp10.chm521000
ms.prod: powerpoint
api_name:
- PowerPoint.AddIn
ms.assetid: e98b609e-97ef-b471-f047-b647bff1e9af
ms.date: 06/08/2017
localization_priority: Normal
---


# AddIn object (PowerPoint)

Represents a single add-in, either loaded or not loaded. 


## Remarks

The **AddIn** object is a member of the **[AddIns](PowerPoint.AddIns.md)** collection. The **AddIns** collection contains all of the Microsoft PowerPoint-specific add-ins available, regardless of whether or not they are loaded. The collection does not include Component Object Model (COM) add-ins.


## Example

Use  **AddIns** (_index_), where _index_ is the add-in's title or index number, to return a single **AddIn** object. The following example loads the My Ppt Tools add-in.


```vb
AddIns("my ppt tools").Loaded = True
```

The add-in title, shown above, should not be confused with the add-in name, which is the file name of the add-in. You must spell the add-in title exactly as it is spelled in the  **Add-Ins** tab, but the capitalization does not have to match.

The index number represents the position of the add-in in the  **Add-Ins** tab. The following example displays the names of all the add-ins that are currently loaded in PowerPoint.




```vb
For i = 1 To AddIns.Count

    If AddIns(i).Loaded Then MsgBox AddIns(i).Name

Next
```

Use the [Add](PowerPoint.AddIns.Add.md)method to add a PowerPoint-specific add-in to the list of those available. Note, however, that using this method does not load the add-in. To load the add-in, set the [Loaded](PowerPoint.AddIn.Loaded.md)property of the add-in to  **True** after you use the **Add** method. You can perform both of these actions in a single step, as shown in the following example (note that you use the name of the add-in, not its title, with the **Add** method).




```vb
AddIns.Add("generic.ppa").Loaded = True
```

Use  **AddIns** (_index_), where _index_ is the add-in's title, to return a reference to the loaded add-in. The following example sets the `presAddin` variable to the add-in titled "my ppt tools" and sets the `myName` variable to the name of the add-in.




```vb
Set presAddin = AddIns("my ppt tools")

With presAddin

    myName = .Name

End With
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]