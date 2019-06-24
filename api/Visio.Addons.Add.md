---
title: Addons.Add method (Visio)
keywords: vis_sdr.chm12516655
f1_keywords:
- vis_sdr.chm12516655
ms.prod: visio
api_name:
- Visio.Addons.Add
ms.assetid: e0bc6a13-3063-0e1d-09b8-4a9c377695e6
ms.date: 06/24/2019
localization_priority: Normal
---


# Addons.Add method (Visio)

Adds a new **[Addon](Visio.Addon.md)** object to an **Addons** collection.


## Syntax

_expression_.**Add** (_FileName_)

_expression_ A variable that represents an **[Addons](Visio.Addons.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The name of the add-on.|

## Return value

Addon


## Remarks

The **Add** method adds an EXE or VSL file to the collection and returns an **Addon** object if the string expression specifies an EXE file, or **Nothing** if the string expression specifies a VSL file.


## Example

The following macro shows how to add an **Addon** object to the **Addons** collection.

Before running this macro, replace `path\filename` with a valid path and file name for an add-on in your Visio project.

```vb
Public Sub AddAddon_Example() 
 
 Dim vsoAddons As Visio.Addons 
 Dim vsoAddon As Visio.Addon 
 
 'Add an add-on to the Addons collection. 
 Set vsoAddons = Visio.Addons 
 Set vsoAddon = vsoAddons.Add("path\filename") 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]