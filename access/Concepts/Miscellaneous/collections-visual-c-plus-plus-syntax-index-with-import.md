---
title: Collections (Visual C++ syntax index with import)
ms.prod: access
ms.assetid: 839b8c78-b6dc-ea2b-fe9c-305b8b47b4b9
ms.date: 10/12/2018
localization_priority: Normal
---


# Collections (Visual C++ syntax index with #import)

**Applies to:** Access 2013 | Access 2016

It is useful to know that collections inherit certain common methods and properties.

All collections inherit the **Count** property and **Refresh** method, and all collections add the **Item** property. The **Errors** collection adds the **Clear** method. The **Parameters** collection inherits the **Append** and **Delete** methods, while the **Fields** collection adds the **Append**, **Delete**, and **Update** methods.

## Properties collection

### Methods

```cs
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthadorefresh_HV10294718.xml( ); 
```

### Properties

```cs
long GetCount( ); 
__declspec(property(get=GetCount)) long Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocount_HV10294234.xml; 
 
PropertyPtr GetItem( const _variant_t & Index ); 
__declspec(property(get=GetItem)) PropertyPtr Invalid DDUE based on source, error:link not allowed in code, link filename:mdproitem_HV10294463.xml[]; 

```


## Errors collection

### Methods


```cs
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthclear_HV10294165.xml( ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthadorefresh_HV10294718.xml( ); 
```

### Properties

```cs
long GetCount( ); 
__declspec(property(get=GetCount)) long Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocount_HV10294234.xml; 
 
PropertyPtr GetItem( const _variant_t & Index ); 
__declspec(property(get=GetItem)) PropertyPtr Invalid DDUE based on source, error:link not allowed in code, link filename:mdproitem_HV10294463.xml[]; 
```


## Parameters collection

### Methods

```cs
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthappend_HV10294078.xml( IDispatch * Object ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthcoldelete_HV10294294.xml( const _variant_t & Index ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthadorefresh_HV10294718.xml( ); 
```

### Properties

```cs
long GetCount( ); 
__declspec(property(get=GetCount)) long Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocount_HV10294234.xml; 
 
PropertyPtr GetItem( const _variant_t & Index ); 
__declspec(property(get=GetItem)) PropertyPtr Invalid DDUE based on source, error:link not allowed in code, link filename:mdproitem_HV10294463.xml[]; 
```


## Fields collection

### Methods

```cs
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthappend_HV10294078.xml( _bstr_t Name , enum DataTypeEnum Type , long DefinedSize , 
 enum FieldAttributeEnum Attrib , const _variant_t & FieldValue  = 
 vtMissing ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthcoldeletefield_HV10294293.xml( const _variant_t & Index ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthadorefresh_HV10294718.xml( ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthupdate_HV10294888.xml( ); 
```

### Properties

```cs
long GetCount( ); 
__declspec(property(get=GetCount)) long Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocount_HV10294234.xml; 
 
PropertyPtr GetItem( const _variant_t & Index ); 
__declspec(property(get=GetItem)) PropertyPtr Invalid DDUE based on source, error:link not allowed in code, link filename:mdproitem_HV10294463.xml[]; 
```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]