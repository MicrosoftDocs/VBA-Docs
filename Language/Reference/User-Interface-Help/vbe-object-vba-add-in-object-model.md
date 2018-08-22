---
title: VBE Object (VBA Add-In Object Model)
keywords: vbob6.chm100000
f1_keywords:
- vbob6.chm100000
ms.prod: office
ms.assetid: 82f7d911-5ad9-5e48-c2c0-8a2ebbf14ede
ms.date: 06/08/2017
---


# VBE Object (VBA Add-In Object Model)



The root object that contains all other [objects](../../Glossary/vbe-glossary.md#object) and[collections](../../Glossary/vbe-glossary.md#collection) represented in Visual Basic for Applications.

## Remarks

You can use the following collections to access the objects contained in the  **VBE** object:


- Use the  **VBProjects** collection to access the collection of[projects](../../Glossary/vbe-glossary.md#project).
    
- Use the  **AddIns** collection to access the collection of add-ins.
    
- Use the  **Windows** collection to access the collection of windows.
    
- Use the  **CodePanes** collection to access the collection of[code panes](../../Glossary/vbe-glossary.md#code-pane).
    
- Use the  **CommandBars** collection to access the collection of command bars.
    

Use the  **Events** object to access properties that enable[add-ins](../../Glossary/vbe-glossary.md#add-in) to connect to all events in Visual Basic for Applications. The properties of the **Events** object return objects of the same type as the property name. For example, the **CommandBarEvents** property returns the **CommandBarEvents** object.
You can use the  **SelectedVBComponent** property to return the active component. The active component is the component that is being tracked in the[Project window](../../Glossary/vbe-glossary.md#Project-window). If the selected item in the  **Project** window isn't a component, **SelectedVBComponent** returns **Nothing**.

 **Note**  All objects in this object model have a  **VBE** property that points to the **VBE** object.


