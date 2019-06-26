---
title: AddIns object (Word)
ms.prod: word
ms.assetid: acf58e58-d3f6-23cf-677b-4780f7cbc24d
ms.date: 06/08/2017
localization_priority: Normal
---


# AddIns object (Word)

A collection of **AddIn** objects that represents all the add-ins available to Word, regardless of whether or not they are currently loaded. The **AddIns** collection includes global templates or Word add-in libraries (WLLs) displayed in the **Templates and Add-ins** dialog box.


## Remarks

Use the **AddIns** property to return the **AddIns** collection. The following example displays the name and the installed state of each available add-in.


```vb
For Each ad In AddIns 
 If ad.Installed = True Then 
 MsgBox ad.Name & " is installed" 
 Else 
 MsgBox ad.Name & " is available but not installed" 
 End If 
Next ad
```

Use the **Add** method to add an add-in to the list of available add-ins and (optionally) install it using the Install argument.




```vb
AddIns.Add FileName:="C:\Templates\Other\Letter.dot", Install:=True
```

To install an add-in shown in the list of available add-ins, use the **Installed** property.




```vb
AddIns("Letter.dot").Installed = True
```

Use **AddIns** (_index_), where _index_ is the add-in name or index number, to return a single **[AddIn](Word.AddIn.md)** object. You must exactly match the spelling (but not necessarily the capitalization) of the name, as it is shown in the **Templates and Add-ins** dialog box. To install an add-in shown in the list of available add-ins, use the **Installed** property. The following example loads the Letter.dot template as a global template.




```vb
AddIns("Letter.dot").Installed = True
```


> [!NOTE] 
> If the add-in is not located in the User Templates, Workgroup Templates, or Startup folder, you must specify the full path and file name when indexing an add-in by name.

Use the **Compiled** property to determine whether an **AddIn** object is a template or a WLL.

## Methods

- [Add](Word.AddIns.Add.md)
- [Item](Word.AddIns.Item.md)
- [Unload](Word.AddIns.Unload.md)

## Properties

- [Application](Word.AddIns.Application.md)
- [Count](Word.AddIns.Count.md)
- [Creator](Word.AddIns.Creator.md)
- [Parent](Word.AddIns.Parent.md)

## See also

- [Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]