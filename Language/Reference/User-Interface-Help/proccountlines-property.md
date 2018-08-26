---
title: ProcCountLines Property
keywords: vbob6.chm104004
f1_keywords:
- vbob6.chm104004
ms.prod: office
api_name:
- Office.ProcCountLines
ms.assetid: 4527ede7-80e6-45ec-c645-8a1fd623921f
ms.date: 06/08/2017
---


# ProcCountLines Property



Returns the number of lines in the specified [procedure](../../Glossary/vbe-glossary.md#procedure).

## Syntax

_object_**.ProcCountLines(**_procname_, _prockind_**) As Long**
The  **ProcCountLines** syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the Applies To list.|
| _procname_|Required. A [String](../../Glossary/vbe-glossary.md#string-data-type) containing the name of the procedure.|
| _prockind_|Required. Specifies the kind of procedure to locate. Because [property procedures](../../Glossary/vbe-glossary.md#property-procedure) can have multiple representations in the[module](../../Glossary/vbe-glossary.md#module), you must specify the kind of procedure you want to locate. All procedures other than property procedures (that is,  **Sub** and **Function** procedures) use **vbext_pk_Proc**.|

You can use one of the following [constants](../../Glossary/vbe-glossary.md#constant) for the _prockind_[argument](../../Glossary/vbe-glossary.md#argument):


|**Constant**|**Description**|
|:-----|:-----|
|**vbext_pk_Get**|Specifies a procedure that returns the value of a property.|
|**vbext_pk_Let**|Specifies a procedure that assigns a value to a property.|
|**vbext_pk_Set**|Specifies a procedure that sets a reference to an object.|
|**vbext_pk_Proc**|Specifies all procedures other than property procedures.|

## Remarks

The  **ProcCountLines** property returns the count of all blank or comment lines preceding the procedure declaration and, if the procedure is the last procedure in a[code module](../../Glossary/vbe-glossary.md#code-module), any blank lines following the procedure.

