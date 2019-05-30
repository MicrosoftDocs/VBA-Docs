---
title: Deftype statements (VBA)
keywords: vblr6.chm1008787
f1_keywords:
- vblr6.chm1008787
ms.prod: office
ms.assetid: 14396fc2-494a-9025-d8a5-86174fcc8a74
ms.date: 05/30/2019
localization_priority: Normal
---


# Deftype statements

Used at the [module level](../../Glossary/vbe-glossary.md#module-level) to set the default [data type](../../reference/user-interface-help/data-type-summary.md) for [variables](../../Glossary/vbe-glossary.md#variable), [arguments](../../Glossary/vbe-glossary.md#argument) passed to [procedures](../../Glossary/vbe-glossary.md#procedure), and the return type for **[Function](../../reference/user-interface-help/function-statement.md)** and **[Property Get](../../reference/user-interface-help/property-get-statement.md)** procedures whose names start with the specified characters.

## Syntax

**DefBool** _letterrange_, [ _letterrange_ ] **. . .** <br/>
**DefByte** _letterrange_, [ _letterrange_ ] **. . .** <br/>
**DefInt** _letterrange_, [ _letterrange_ ] **. . .** <br/>
**DefLng** _letterrange_, [ _letterrange_ ] **. . .** <br/>
**DefLngLng** _letterrange_, [ _letterrange_ ] **. . .** (valid on 64-bit platforms only) <br/> 
**DefLngPtr** _letterrange_, [ _letterrange_ ] **. . .** <br/>
**DefCur** _letterrange_, [ _letterrange_ ] **. . .** <br/>
**DefSng** _letterrange_, [ _letterrange_ ] **. . .** <br/>
**DefDbl** _letterrange_, [ _letterrange_ ] **. . .** <br/>
**DefDec** _letterrange_, [ _letterrange_ ] **. . .** <br/>
**DefDate** _letterrange_, [ _letterrange_ ] **. . .** <br/>
**DefStr** _letterrange_, [ _letterrange_ ] **. . .** <br/>
**DefObj** _letterrange_, [ _letterrange_ ] **. . .** <br/>
**DefVar** _letterrange_, [ _letterrange_ ] **. . .**

The required _letterrange_ argument has the following syntax: _letter1_ [ **-** _letter2_ ]

The _letter1_ and _letter2_ arguments specify the name range for which you can set a default data type. Each argument represents the first letter of the variable, argument, **Function** procedure, or **Property Get** procedure name, and can be any letter of the alphabet. The case of letters in _letterrange_ isn't significant.

## Remarks

The statement name determines the data type.

<br/>

|Statement|Data type|
|:-----|:-----|
|**DefBool**|[Boolean](../../Glossary/vbe-glossary.md#boolean-data-type)|
|**DefByte**|[Byte](../../Glossary/vbe-glossary.md#byte-data-type)|
|**DefInt**|[Integer](../../Glossary/vbe-glossary.md#integer-data-type)|
|**DefLng**|[Long](../../Glossary/vbe-glossary.md#long-data-type)|
|**DefLngLng**|[LongLong](../../reference/User-Interface-Help/longlong-data-type.md) (valid on 64-bit platforms only)|
|**DefLngPtr**|[LongPtr](../../reference/User-Interface-Help/longptr-data-type.md)|
|**DefCur**|[Currency](../../Glossary/vbe-glossary.md#currency-data-type)|
|**DefSng**|[Single](../../Glossary/vbe-glossary.md#single-data-type)|
|**DefDbl**|[Double](../../Glossary/vbe-glossary.md#double-data-type)|
|**DefDec**|[Decimal](../../Glossary/vbe-glossary.md#decimal-data-type) (not currently supported)|
|**DefDate**|[Date](../../Glossary/vbe-glossary.md#date-data-type)|
|**DefStr**|[String](../../Glossary/vbe-glossary.md#string-data-type)|
|**DefObj**|[Object](../../Glossary/vbe-glossary.md#object)|
|**DefVar**|[Variant](../../Glossary/vbe-glossary.md#variant-data-type)|

<br/>

For example, in the following program fragment, `Message` is a string variable.

```vb
DefStr A-Q
. . .
Message = "Out of stack space."
```

A **Def**_type_ statement affects only the [module](../../Glossary/vbe-glossary.md#module) where it is used. For example, a **DefInt** statement in one module affects only the default data type of variables, arguments passed to procedures, and the return type for **Function** and **Property Get** procedures declared in that module; the default data type of variables, arguments, and return types in other modules is unaffected. If not explicitly declared with a **Def**_type_ statement, the default data type for all variables, all arguments, all **Function** procedures, and all **Property Get** procedures is **Variant**.

When you specify a letter range, it usually defines the data type for variables that begin with letters in the [first 128 characters of the character set](../../reference/user-interface-help/character-set-0127.md). However, when you specify the letter range A&ndash;Z, you set the default to the specified data type for all variables, including variables that begin with international characters from the [extended part of the character set (128&ndash;255)](../../reference/user-interface-help/character-set-128255.md).

After the range A-Z has been specified, you can't further redefine any subranges of variables by using **Def**_type_ statements. After a range has been specified, if you include a previously defined letter in another **Def**_type_ statement, an error occurs. However, you can explicitly specify the data type of any variable, defined or not, by using a **[Dim](../../reference/user-interface-help/dim-statement.md)** statement with an **As** _type_ clause. 

For example, you can use the following code at the module level to define a variable as a **Double** even though the default data type is **Integer**. 

```vb
DefInt A-Z
Dim TaxRate As Double
```

**Def**_type_ statements don't affect elements of [user-defined types](../../Glossary/vbe-glossary.md#user-defined-type) because the elements must be explicitly declared.


<!--[MISSING EXAMPLE CODE] ## Example

This example shows various uses of the **Def**_type_ statements to set default data types of variables and function procedures whose names start with specified characters. The default data type can be overridden only by explicit assignment by using the **Dim** statement. **Def**_type_ statements can only be used at the module level (that is, not within procedures).--> 


## See also

- [Data types](../../reference/user-interface-help/data-type-summary.md)
- [Statements](../../reference/statements.md)
- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]