---
title: Raise method (Visual Basic for Applications)
keywords: vblr6.chm1014183
f1_keywords:
- vblr6.chm1014183
ms.prod: office
api_name:
- Office.Raise
ms.assetid: 7e3ddb06-db93-ebce-7562-8a15c49261b1
ms.date: 12/14/2018
localization_priority: Normal
---


# Raise method

Generates a [run-time error](../../Glossary/vbe-glossary.md#run-time-error).

## Syntax

_object_.**Raise** _number_, _source_, _description_, _helpfile_, _helpcontext_

<br/>

The **Raise** method has the following object qualifier and [named arguments](../../Glossary/vbe-glossary.md#named-argument):

|Argument|Description|
|:-----|:-----|
|_object_|Required. Always the **[Err](err-object.md)** object.|
|_number_|Required. [Long](../../Glossary/vbe-glossary.md#long-data-type) integer that identifies the nature of the error. Visual Basic errors (both Visual Basic-defined and user-defined errors) are in the range 0&ndash;65535. The range 0&ndash;512 is reserved for system errors; the range 513&ndash;65535 is available for user-defined errors.<br/><br/>When setting the **[Number](number-property-visual-basic-for-applications.md)** property to your own error code in a class module, you add your error code number to the **vbObjectError** [constant](../../Glossary/vbe-glossary.md#constant). For example, to generate the [error number](../../Glossary/vbe-glossary.md#error-number) 513, assign **vbObjectError** + 513 to the **Number** property.|
|_source_|Optional. [String expression](../../Glossary/vbe-glossary.md#string-expression) naming the object or application that generated the error. When setting the **[Source](source-property-visual-basic-for-applications.md)** property for an object, use the form _project.class_. If _source_ is not specified, the programmatic ID of the current Visual Basic [project](../../Glossary/vbe-glossary.md#project) is used.|
|_description_|Optional. String expression describing the error. If unspecified, the value in **Number** is examined. If it can be mapped to a Visual Basic run-time error code, the string that would be returned by the **[Error](error-function.md)** function is used as **[Description](description-property-visual-basic-for-applications.md)**. If there is no Visual Basic error corresponding to **Number**, the "Application-defined or object-defined error" message is used.|
|_helpfile_|Optional. The fully qualified path to the Help file in which help on this error can be found. If unspecified, Visual Basic uses the fully qualified drive, path, and file name of the Visual Basic Help file. See **[HelpFile](helpfile-property.md)**.|
|_helpcontext_|Optional. The context ID identifying a topic within _helpfile_ that provides help for the error. If omitted, the Visual Basic Help file context ID for the error corresponding to the **Number** property is used, if it exists. See **[HelpContext](helpcontext-property-visual-basic-for-applications.md)**.|



## Remarks 

All of the [arguments](../../Glossary/vbe-glossary.md#argument) are optional except _number_. If you use **Raise** without specifying some arguments, and the property settings of the **Err** object contain values that have not been cleared, those values serve as the values for your error.

**Raise** is used for generating run-time errors and can be used instead of the **[Error](error-statement.md)** statement.

**Raise** is useful for generating errors when writing class modules, because the **Err** object gives richer information than is possible if you generate errors with the **Error** statement. For example, with the **Raise** method, the source that generated the error can be specified in the **Source** property, online Help for the error can be referenced, and so on.

## Example

This example uses the **Err** object's **Raise** method to generate an error within an Automation object written in Visual Basic. It has the programmatic ID `MyProj.MyObject`. On the MacIntosh, the default drive name is "HD" and portions of the pathname are separated by colons instead of backslashes.


```vb
Const MyContextID = 1010407    ' Define a constant for contextID.
Function TestName(CurrentName, NewName)
    If Instr(NewName, "bob") Then    ' Test the validity of NewName.
        ' Raise the exception
        Err.Raise vbObjectError + 513, "MyProj.MyObject", _
        "No ""bob"" allowed in your name", "c:\MyProj\MyHelp.Hlp", _
        MyContextID
    End If
End Function
```


## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
