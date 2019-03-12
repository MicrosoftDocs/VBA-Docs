---
title: Check or add an object library reference (VBA)
keywords: vbhw6.chm1107739
f1_keywords:
- vbhw6.chm1107739
ms.prod: office
ms.assetid: a04227a8-80e0-2eb3-52bb-f992d8bb5e68
ms.date: 12/27/2018
localization_priority: Priority
---


# Check or add an object library reference

If you use the objects in other applications as part of your Visual Basic application, you may want to establish a reference to the [object libraries](../Glossary/vbe-glossary.md#object-library) of those applications. Before you can do that, you must first be sure that the application provides an object library.

**To see if an application provides an object library**

1. From the **[Tools](../reference/user-interface-help/tools-menu.md)** menu, choose **References** to display the **[References](../reference/user-interface-help/references-dialog-box.md)** dialog box.
    
2. The **References** dialog box shows all object libraries registered with the operating system. Scroll through the list for the application whose object library you want to reference. If the application isn't listed, you can use the **Browse** button to search for object libraries (\*.olb and \*.tlb) or [executable files](../Glossary/vbe-glossary.md#executable-file) (\*.exe and \*.dll on Windows). References whose check boxes are selected are used by your [project](../Glossary/vbe-glossary.md#project); those that aren't selected are not used, but can be added.
    

**To add an object library reference to your project**

- Select the object library reference in the **Available References** box in the **References** dialog box and choose **OK**. Your Visual Basic project now has a reference to the application's object library. If you open the **[Object Browser](../reference/user-interface-help/object-browser.md)** (press F2) and select the application's library, it displays the objects provided by the selected object library, as well as each object's [methods](../Glossary/vbe-glossary.md#method) and [properties](../Glossary/vbe-glossary.md#property). 

  In the **Object Browser**, you can select a [class](../Glossary/vbe-glossary.md#class) in the **Classes** box and select a method or property in the **Members** box. Use copy and paste to add the syntax to your code.
    

## See also

- [Visual Basic how-to topics](../reference/user-interface-help/visual-basic-how-to-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
