---
title: ObjectProxy (ADO/WFC syntax)
ms.prod: access
ms.assetid: 8e3224b7-0b1d-1e08-eaa7-ceb0b6f5411c
ms.date: 10/12/2018
localization_priority: Normal
---


# ObjectProxy (ADO/WFC syntax)

**Applies to:** Access 2013 | Access 2016

An **ObjectProxy** object represents a server, and is returned by the **createObject** method of the [DataSpace](https://msdn.microsoft.com/library/7db181d5-422b-49fe-b6af-a20f5da520ff%28Office.15%29.aspx) object. The ObjectProxy class has one method, **call**, which can invoke a method on the server and return an object resulting from that invocation.

**package com.ms.wfc.data**

## Methods

### Call method

Invokes a method on the server represented by the ObjectProxy. Optionally, method arguments may be passed as an array of objects.


## Syntax

```vb
public Object ObjectProxy .call( String method  ) 
public Object ObjectProxy .call( String method , Object[] args ) 

```


## Return value

Object
    
- An object resulting from invoking the method.
    

## Parameters

_ObjectProxy_
    
- An **ObjectProxy** object that represents the server.
    
_method_
    
- A String, containing the name of the method to invoke on the server.
    
_args_
    
- Optional. An array of objects that are arguments to the method on the server. Java data types are automatically converted to data types suitable for use on the server.
    
## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]