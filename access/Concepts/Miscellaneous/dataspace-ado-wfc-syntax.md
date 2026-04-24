---
title: DataSpace (ADO/WFC syntax)
ms.assetid: 52bc0aa1-b3e6-4d2c-9a73-a9f185d028c4
ms.date: 10/12/2018
ms.localizationpriority: medium
---


# DataSpace (ADO/WFC syntax)

**Applies to:** Access 2013 | Access 2016

The **createObject** method of the **DataSpace** class specifies both a business object to process client application requests ( _progid_ ) and the communications protocol and server ( _connection_ ). **createObject** returns an [ObjectProxy](objectproxy-ado-wfc-syntax.md) object that represents the server.

**package com.ms.wfc.data**

## Constructor

```vb
public DataSpace() 

```

### Methods

```vb
public static ObjectProxyDataSpace. Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthcreateobj_HV10294242.xml(String 
 progid , String connection ) 

```

### Properties

```vb
public static int Invalid DDUE based on source, error:link not allowed in code, link filename:mdprointernettimeout_HV10294450.xml() 
public static void setInternetTimeout(int plInetTimeout ) 

```

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]