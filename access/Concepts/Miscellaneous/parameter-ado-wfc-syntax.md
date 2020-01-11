---
title: Parameter (ADO/WFC syntax)
ms.prod: access
ms.assetid: e5e2ec60-4d62-5959-a8a0-d1f94b54fb7f
ms.date: 10/12/2018
localization_priority: Normal
---


# Parameter (ADO/WFC syntax)

**Applies to:** Access 2013 | Access 2016

**package com.ms.wfc.data**

## Constructor

```vb
public Invalid DDUE based on source, error:link not allowed in code, link filename:mdobjparameter_HV10294590.xml() 
public Parameter(String name ) 
public Parameter(String name , int type ) 
public Parameter(String name , int type , int dir ) 
public Parameter(String name , int type , int dir , int size ) 
public Parameter(String name , int type , int dir , int size , Object value ) 

```

## Methods

```vb
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthappchunk_HV10294090.xml(byte[] bytes ) 
public void appendChunk(char[] chars ) 
public void appendChunk(String chars ) 

```

## Properties

```vb
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdproattributes_HV10294098.xml() 
public void setAttributes(int attr ) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdprodirection_HV10294320.xml() 
public void setDirection(int dir ) 
public String Invalid DDUE based on source, error:link not allowed in code, link filename:mdproname_HV10294535.xml() 
public void setName(String name ) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdpronumericscale_HV10294551.xml() 
public void setNumericScale(int scale) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdproprecision_HV10294615.xml() 
public void setPrecision(int prec) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdprosize_HV10294779.xml() 
public void setSize(int size ) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdprotype_HV10294866.xml() 
public void setType(int type ) 
public com.ms.com.Variant Invalid DDUE based on source, error:link not allowed in code, link filename:mdprovalue_HV10294920.xml() 
public void setValue(Object v ) 
public AdoProperties Invalid DDUE based on source, error:link not allowed in code, link filename:mdcolproperties_HV10294633.xml() 

```


## Parameter accessor methods

The [Value](https://msdn.microsoft.com/library/ff21d122-98e3-2b48-d92f-e696b8079fc5%28Office.15%29.aspx) property of a [Parameter](https://msdn.microsoft.com/library/7577598e-3d0c-30c6-5f24-1cfe98791798%28Office.15%29.aspx) object gets or sets the content of that object. The content is represented as a VARIANT, a type of object that can be assigned a value and any of several data types.

ADO/WFC implements the **Value** property with the **getValue** method, which returns a VARIANT object; and the **setValue** method, which takes a VARIANT as an argument. VARIANTs are highly efficient in certain languages, such as Microsoft Visual Basic. However, you can attain better performance in Microsoft Visual J++ by using native Java data types.

In addition to the **Value** property, ADO/WFC provides _accessor_ methods that use Java data types to get and set the content of **Parameter** objects. Most of these methods have names of the form **get** _DataType_ or **set** _DataType_.

There is one noteworthy exception: There is no **getNull** property; instead, there is an **isNull** property that returns a Boolean value indicating whether the field is null.

```vb
public boolean getBoolean() 
public void setBoolean(boolean v ) 
public byte getByte() 
public void setByte(byte v ) 
public double getDouble() 
public void setDouble(double v ) 
public float getFloat() 
public void setFloat(float v ) 
public int getInt() 
public void setInt(int v ) 
public long getLong() 
public void setLong(long v ) 
public short getShort() 
public void setShort(short v ) 
public String getString() 
public void setString(String v ) 
public boolean isNull() 
public void setNull() 

```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]