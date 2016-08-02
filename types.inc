{****************************************************************************
                        types.h
****************************************************************************}

type

  TEnumVar = (
    VTYPE_EMPTY         = 0,
    VTYPE_NULL,
    VTYPE_I2,                             //int16_t
    VTYPE_I4,                             //int32_t
    VTYPE_R4,                             //float
    VTYPE_R8,                             //double
    VTYPE_DATE,                           //DATE (double)
    VTYPE_TM,                             //struct tm
    VTYPE_PSTR,                           //struct str    string
    VTYPE_INTERFACE,                      //struct iface
    VTYPE_ERROR,                          //int32_t errCode
    VTYPE_BOOL,                           //bool
    VTYPE_VARIANT,                        //struct _tVariant *
    VTYPE_I1,                             //int8_t
    VTYPE_UI1,                            //uint8_t
    VTYPE_UI2,                            //uint16_t
    VTYPE_UI4,                            //uint32_t
    VTYPE_I8,                             //int64_t
    VTYPE_UI8,                            //uint64_t
    VTYPE_INT,                            //int   Depends on architecture
    VTYPE_UINT,                           //unsigned int  Depends on architecture
    VTYPE_HRESULT,                        //long hRes
    VTYPE_PWSTR,                          //struct wstr
    VTYPE_BLOB,                           //means in struct str binary data contain
    VTYPE_CLSID,                          //UUID
    VTYPE_STR_BLOB      = $fff,
    VTYPE_VECTOR        = $1000,
    VTYPE_ARRAY         = $2000,
    VTYPE_BYREF         = $4000,          //Only with struct _tVariant *
    VTYPE_RESERVED      = $8000,
    VTYPE_ILLEGAL       = $ffff);

type
  P1CVariant = ^T1CVariant;

  TTM = record
    tm_sec: cint;
    tm_min: cint;
    tm_hour: cint;
    tm_mday: cint;
    tm_mon: cint;
    tm_year: cint;
    tm_wday: cint;
    tm_yday: cint;
    tm_isdst: cint;
  end;

  TVarValue = record case TEnumVar of
    VTYPE_I1: (i8Val: cint8);
    VTYPE_I2: (shortVal: cint16);
    VTYPE_I4: (lVal: cint32);
    VTYPE_INT: (intVal: cint);
    VTYPE_UINT: (uintVal: cuint);
    VTYPE_I8: (llVal: cint64);
    VTYPE_UI1: (ui8Val: cuint8);
    VTYPE_UI2: (ushortVal: cuint16);
    VTYPE_UI4: (ulVal: cuint32);
    VTYPE_UI8: (ullVal: cuint64);
    VTYPE_ERROR: (errCode: cint32);
    VTYPE_HRESULT: (hRes: clong);
    VTYPE_R4: (fltVal: cfloat);
    VTYPE_R8: (dblVal: cdouble);
    VTYPE_BOOL: (bVal: cbool);
    //VTYPE_???: (chVal: cchar);
    //VTYPE_???: (wchVal: WChar);
    VTYPE_DATE: (date: cdouble);
    //VTYPE_???: (IDVal: TGuid???);
    VTYPE_VARIANT: (pvarVal: P1CVariant);
    VTYPE_TM: (tmVal: TTM);
    VTYPE_INTERFACE: (pInterfaceVal: Pointer; InterfaceID: TGuid);
    VTYPE_PSTR: (pstrVal: PChar; strLen: cuint32);
    VTYPE_PWSTR: (pwstrVal: PWideChar; wstrLen: cuint32);
  end;

  T1CVariant = record
    Value: TVarValue;
    cbElements: cuint32;
    vt: TEnumVar;
  end;

