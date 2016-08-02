{****************************************************************************
                        Converting variants
****************************************************************************}

function From1CVariant(const PropVal: T1CVariant; var Value: Variant): Boolean;
var
  ST: TSystemTime;
begin
  case PropVal.vt of
    VTYPE_EMPTY:
      begin
        Value := Unassigned;
        Result := True;
      end;
    VTYPE_NULL:
      begin
        Value := Null;
        Result := True;
      end;
    VTYPE_I2:
      begin
        Value := PropVal.Value.shortVal;
        Result := True;
      end;
    VTYPE_I4:
      begin
        Value := PropVal.Value.lVal;
        Result := True;
      end;
    VTYPE_R4:
      begin
        Value := PropVal.Value.fltVal;
        Result := True;
      end;
    VTYPE_R8:
      begin
        Value := PropVal.Value.dblVal;
        Result := True;
      end;
    VTYPE_DATE:
      begin
        Value := PropVal.Value.date;
        Result := True;
      end;
    VTYPE_TM:
      begin
        ST.Second := PropVal.Value.tmVal.tm_sec;
        ST.Minute := PropVal.Value.tmVal.tm_min;
        ST.Hour := PropVal.Value.tmVal.tm_hour;
        ST.Day := PropVal.Value.tmVal.tm_mday;
        ST.Month := PropVal.Value.tmVal.tm_mon;
        ST.Year := PropVal.Value.tmVal.tm_year + 1900;
        ST.DayOfWeek := PropVal.Value.tmVal.tm_wday + 1;
        Value := SystemTimeToDateTime(ST);
        Result := True;
      end;
    VTYPE_PSTR:
      begin
        Value := String(PropVal.Value.pstrVal);
        Result := True;
      end;
    VTYPE_ERROR:
      begin
        Value := Unassigned;
        TVarData(Value).vtype := varError;
        TVarData(Value).verror := PropVal.Value.errCode;
        Result := True;
      end;
    VTYPE_BOOL:
      begin
        Value := Boolean(PropVal.Value.bVal);
        Result := True;
      end;
    VTYPE_I1:
      begin
        Value := PropVal.Value.i8Val;
        Result := True;
      end;
    VTYPE_UI1:
      begin
        Value := PropVal.Value.ui8Val;
        Result := True;
      end;
    VTYPE_UI2:
      begin
        Value := PropVal.Value.ushortVal;
        Result := True;
      end;
    VTYPE_UI4:
      begin
        Value := PropVal.Value.ulVal;
        Result := True;
      end;
    VTYPE_I8:
      begin
        Value := PropVal.Value.llVal;
        Result := True;
      end;
    VTYPE_UI8:
      begin
        Value := PropVal.Value.ullVal;
        Result := True;
      end;
    VTYPE_INT:
      begin
        Value := PropVal.Value.intVal;
        Result := True;
      end;
    VTYPE_UINT:
      begin
        Value := PropVal.Value.uintVal;
        Result := True;
      end;
    VTYPE_HRESULT:
      begin
        Value := PropVal.Value.hRes;
        Result := True;
      end;
    VTYPE_PWSTR:
      begin
        Value := WideCharToString(PropVal.Value.pwstrVal);
        Result := True;
      end;
    VTYPE_BLOB:
      begin
        Value := Unassigned;
        TVarData(Value).vtype := BinaryDataFactory.VarType;
        TVarData(Value).vpointer := GetMem(PropVal.Value.strLen + SizeOf(SizeInt));
        P1CBinaryData(TVarData(Value).vpointer)^.Size := PropVal.Value.strLen;
        Move(PropVal.Value.pstrVal^, P1CBinaryData(TVarData(Value).vpointer)^.Data, P1CBinaryData(TVarData(Value).vpointer)^.Size);
        Result := True;
      end;
    //VTYPE_ARRAY:
  else
    Value := Unassigned;
    Result := False;
  end;
end;

function To1CVariant(const Value: Variant; var PropVal: T1CVariant; const MemoryManager: P1CMemoryManager): Boolean;
var
  VarData: TVarData absolute Value;
begin
  case VarData.vtype of
    varempty:
      begin
        PropVal.vt := VTYPE_EMPTY;
        Result := True;
      end;
    varnull:
      begin
        PropVal.vt := VTYPE_NULL;
        Result := True;
      end;
    varsmallint:
      begin
        PropVal.vt := VTYPE_I2;
        PropVal.Value.shortVal := VarData.vsmallint;
        Result := True;
      end;
    varinteger:
      begin
        PropVal.vt := VTYPE_I4;
        PropVal.Value.lVal := VarData.vinteger;
        Result := True;
      end;
    varsingle:
      begin
        PropVal.vt := VTYPE_R4;
        PropVal.Value.fltVal := VarData.vsingle;
        Result := True;
      end;
    vardouble:
      begin
        PropVal.vt := VTYPE_R8;
        PropVal.Value.dblVal := VarData.vdouble;
        Result := True;
      end;
    vardate:
      begin
        PropVal.vt := VTYPE_DATE;
        PropVal.Value.date := VarData.vdate;
        Result := True;
      end;
    varcurrency:
      begin
        PropVal.vt := VTYPE_R8;
        PropVal.Value.dblVal := Value;
        Result := True;
      end;
    varolestr:
      begin
        PropVal.vt := VTYPE_PWSTR;
        PropVal.Value.wstrLen := Length(VarData.volestr);
        Result := MemoryManager^.__vfptr^.AllocMemory(MemoryManager, PropVal.Value.pwstrVal, (PropVal.Value.wstrLen + 1) * SizeOf(WideChar));
        if not Result then
          begin
            PropVal.vt := VTYPE_EMPTY;
            Exit;
          end;
        if PropVal.Value.wstrLen = 0 then
          Move(EmptyWideChar, PropVal.Value.pwstrVal^, (PropVal.Value.wstrLen + 1) * SizeOf(WideChar))
        else
          Move(VarData.volestr^, PropVal.Value.pwstrVal^, (PropVal.Value.wstrLen + 1) * SizeOf(WideChar));
      end;
    varerror:
      begin
        PropVal.vt := VTYPE_ERROR;
        PropVal.Value.errCode := VarData.verror;
        Result := True;
      end;
    varboolean:
      begin
        PropVal.vt := VTYPE_BOOL;
        PropVal.Value.bVal := VarData.vboolean;
        Result := True;
      end;
    varshortint:
      begin
        PropVal.vt := VTYPE_I1;
        PropVal.Value.i8Val := VarData.vshortint;
        Result := True;
      end;
    varbyte:
      begin
        PropVal.vt := VTYPE_UI1;
        PropVal.Value.ui8Val := VarData.vbyte;
        Result := True;
      end;
    varword:
      begin
        PropVal.vt := VTYPE_UI2;
        PropVal.Value.ushortVal := VarData.vword;
        Result := True;
      end;
    varlongword:
      begin
        PropVal.vt := VTYPE_UI4;
        PropVal.Value.ulVal := VarData.vlongword;
        Result := True;
      end;
    varint64:
      begin
        PropVal.vt := VTYPE_I8;
        PropVal.Value.llVal := VarData.vint64;
        Result := True;
      end;
    varqword:
      begin
        PropVal.vt := VTYPE_UI8;
        PropVal.Value.ullVal := VarData.vqword;
        Result := True;
      end;
    //varrecord:
    varstring:
      begin
        PropVal.vt := VTYPE_PWSTR;
        PropVal.Value.wstrLen := Length(String(VarData.vstring));
        Result := MemoryManager^.__vfptr^.AllocMemory(MemoryManager, PropVal.Value.pwstrVal, (PropVal.Value.wstrLen + 1) * SizeOf(WideChar));
        if not Result then
          begin
            PropVal.vt := VTYPE_EMPTY;
            Exit;
          end;
        if PropVal.Value.wstrLen = 0 then
          Move(EmptyWideChar, PropVal.Value.pwstrVal^, (PropVal.Value.wstrLen + 1) * SizeOf(WideChar))
        else
          Move(UTF8Decode(String(VarData.vstring))[1], PropVal.Value.pwstrVal^, (PropVal.Value.wstrLen + 1) * SizeOf(WideChar));
        Result := True;
      end;
    //varustring:
    //vararray:
  else
    if VarData.vType = BinaryDataFactory.VarType then
      begin
        PropVal.vt := VTYPE_BLOB;
        PropVal.Value.strLen := P1CBinaryData(VarData.vpointer)^.Size;
        Result := MemoryManager^.__vfptr^.AllocMemory(MemoryManager, PropVal.Value.pstrVal, PropVal.Value.wstrLen);
        if Result then
          Move(P1CBinaryData(VarData.vpointer)^.Data, PropVal.Value.pstrVal^, P1CBinaryData(VarData.vpointer)^.Size)
        else
          PropVal.vt := VTYPE_EMPTY;
      end
    else
      begin
        PropVal.vt := VTYPE_EMPTY;
        Result := False;
      end;
  end;
end;

