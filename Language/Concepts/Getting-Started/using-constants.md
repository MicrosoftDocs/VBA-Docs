---
title: Using constants (VBA)
keywords: vbcn6.chm1076680
f1_keywords:
- vbcn6.chm1076680
ms.prod: office
ms.assetid: 44381f2a-d3a9-0341-80f6-7682a3469951
ms.date: 12/26/2018
localization_priority: Normal
---


# Using constants

Your code might contain frequently occurring constant values, or might depend on certain numbers that are difficult to remember and have no obvious meaning. You can make your code easier to read and maintain by using [constants](../../Glossary/vbe-glossary.md#constant). A constant is a meaningful name that takes the place of a number or string that does not change. You can't modify a constant or assign a new value to it as you can a [variable](../../Glossary/vbe-glossary.md#variable).

## Types of constants

There are three types of constants:

- [Intrinsic constants](../../Glossary/vbe-glossary.md#intrinsic-constants), or system-defined constants, are provided by applications and controls. Other applications that provide [object libraries](../../Glossary/vbe-glossary.md#object-library), such as Microsoft Access, Excel, Project, and Word also provide a list of constants that you can use with their objects, methods, and properties. You can get a list of the constants provided for individual object libraries in the [Object Browser](../../Glossary/vbe-glossary.md#object-browser).

  Visual Basic constants are listed in the Visual Basic for Applications type library and Data Access Object (DAO) library.

  > [!NOTE] 
  > Visual Basic continues to recognize constants in applications created in earlier versions of Visual Basic or Visual Basic for Applications. You can upgrade your constants to those listed in the **[Object Browser](../../reference/user-interface-help/object-browser.md)**. Constants listed in the **Object Browser** don't have to be declared in your application.

- Symbolic or user-defined constants are declared by using the **[Const](../../reference/user-interface-help/const-statement.md)** statement.
    
- [Conditional compiler constants](../../Glossary/vbe-glossary.md#conditional-compiler-constant) are declared by using the **[#Const](../../reference/user-interface-help/const-directive.md)** statement (directive).
    

In earlier versions of Visual Basic, constant names were usually capitalized with underscores. For example:

```vb
TILE_HORIZONTAL 

```

Intrinsic constants are now qualified to avoid confusion when constants with the same name exist in more than one object library, which may have different values assigned to them. There are two ways to qualify constant names:

- By prefix
- By library reference
    
## Qualifying constants by prefix

The intrinsic constants supplied by all objects appear in a mixed-case format, with a 2-character prefix indicating the object library that defines the constant. Constants from the Visual Basic for Applications object library are prefaced with "vb" and constants from the Microsoft Excel object library are prefaced with "xl". The following examples illustrate how prefixes for custom controls vary, depending on the [type library](../../Glossary/vbe-glossary.md#type-library).

- **vbTileHorizontal**
- **xlDialogBorder**
    
## Qualifying constants by library reference

You can also qualify the reference to a constant by using the following syntax.

[ _libname_.] [ _modulename_.] _constname_

The syntax for qualifying constants has these parts:

|Part|Description|
|:-----|:-----|
| _libname_|Optional. The name of the type library that defines the constant. For most custom controls (not available on the Macintosh), this is also the [class](../../Glossary/vbe-glossary.md#class) name of the control. If you don't remember the class name of the control, position the mouse pointer over the control in the toolbox. The class name is displayed in the **ToolTip**.|
| _modulename_|Optional. The name of the module within the type library that defines the constant. You can find the name of the module by using the **Object Browser**.|
| _constname_|The name defined for the constant in the type library.|


For example:

```vb
Threed.LeftJustify 

```


## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]