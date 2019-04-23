---
title: Set reference to a type library (VBA)
keywords: vbhw6.chm1105230
f1_keywords:
- vbhw6.chm1105230
ms.prod: office
ms.assetid: 5b695bf5-5ab3-977e-b037-13aea3097b9c
ms.date: 12/27/2018
localization_priority: Normal
---


# Set reference to a type library

Automation (formerly OLE Automation) enables you to use objects from other applications in Visual Basic code. An application that provides its objects for use by other applications also provides information about those objects in a [type library](../Glossary/vbe-glossary.md#type-library). To achieve the best possible performance when using another application's objects, you should set a reference to that application's type library.

**To set a reference to an application's type library**

1. Choose **References** on the **[Tools](../reference/user-interface-help/tools-menu.md)** menu.
    
2. Select the check boxes for the applications with type libraries that you want to reference.
    

If you are writing code that manipulates objects in another application, you should set a reference to that application's type library for best possible access to those objects. You don't have to set a reference to use another application's objects, but doing so provides several advantages for your application.

Your code will run faster if you set a reference to another application's type library before you work with its objects. If you set a reference, you can declare an [object variable](../Glossary/vbe-glossary.md#object-variable) representing an object in the other application as its most specific type. For example, if you are writing code to work with Microsoft Excel objects, you can declare an object variable of type **Excel.Application** if you created a reference to the Excel type library. 

The following code is the fastest way to create a variable to represent the Excel **Application** object.

```vb
Dim appXL As Excel.Application 

```

If you haven't set a reference to the Excel type library, you must declare the [variable](../Glossary/vbe-glossary.md#variable) as a generic variable of type [Object](../Glossary/vbe-glossary.md#object). The following code runs more slowly.

```vb
Dim appXL As Object 

```

If you set a reference to an application's type library, all of its objects and their [methods](../Glossary/vbe-glossary.md#method) and [properties](../Glossary/vbe-glossary.md#property) are listed in the **Object Browser**. This makes it easy to determine what properties and methods are available to each object.

For Microsoft applications that can also serve as Automation servers, you can set references to their type libraries from another application, and control their objects from that application.

## See also

- [Visual Basic how-to topics](../reference/user-interface-help/visual-basic-how-to-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]