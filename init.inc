{****************************************************************************
                        Init and clear common data
****************************************************************************}

procedure InitUnit;
begin
  AppCapabilities := eAppCapabilitiesInvalid;

  ClassNames := nil;

  FactoryList := TStringList.Create;
  FactoryList.OwnsObjects := True;
  FactoryList.CaseSensitive := False;
  FactoryList.Delimiter := '|';

  ComponentBase.InitDoneBase.__vfptr := GetMem(SizeOf(TInitDoneBaseVTable));
  with ComponentBase.InitDoneBase.__vfptr^ do
    begin
      _Destructor := nil;
      Init := @InitDoneBase_Init;
      setMemManager := @InitDoneBase_setMemManager;
      GetInfo := @InitDoneBase_GetInfo;
      Done := @InitDoneBase_Done;
    end;

  ComponentBase.LanguageExtenderBase.__vfptr := GetMem(SizeOf(TLanguageExtenderBaseVTable));
  with ComponentBase.LanguageExtenderBase.__vfptr^ do
    begin
      _Destructor := nil;
      RegisterExtensionAs := @LanguageExtenderBase_RegisterExtensionAs;
      GetNProps := @LanguageExtenderBase_GetNProps;
      FindProp := @LanguageExtenderBase_FindProp;
      GetPropName := @LanguageExtenderBase_GetPropName;
      GetPropVal := @LanguageExtenderBase_GetPropVal;
      SetPropVal := @LanguageExtenderBase_SetPropVal;
      IsPropReadable := @LanguageExtenderBase_IsPropReadable;
      IsPropWritable := @LanguageExtenderBase_IsPropWritable;
      GetNMethods := @LanguageExtenderBase_GetNMethods;
      FindMethod := @LanguageExtenderBase_FindMethod;
      GetMethodName := @LanguageExtenderBase_GetMethodName;
      GetNParams := @LanguageExtenderBase_GetNParams;
      GetParamDefValue := @LanguageExtenderBase_GetParamDefValue;
      HasRetVal := @LanguageExtenderBase_HasRetVal;
      CallAsProc := @LanguageExtenderBase_CallAsProc;
      CallAsFunc := @LanguageExtenderBase_CallAsFunc;
    end;

  ComponentBase.LocaleBase.__vfptr := GetMem(SizeOf(TLocaleBaseVTable));
  with ComponentBase.LocaleBase.__vfptr^ do
    begin
      _Destructor := nil;
      SetLocale := @LocaleBase_SetLocale;
    end;

  ComponentBase.__vfptr := GetMem(SizeOf(TComponentBaseVTable));
  with ComponentBase.__vfptr^ do
    begin
      _Destructor := nil;
    end;

  ComponentBase.Data := nil;

  BinaryDataFactory := T1CBinaryDataFactory.Create;
end;

procedure ClearUnit;
begin
  if ClassNames <> nil then
    FreeMem(ClassNames);

  FreeAndNil(FactoryList);

  FreeMem(ComponentBase.InitDoneBase.__vfptr);
  FreeMem(ComponentBase.LanguageExtenderBase.__vfptr);
  FreeMem(ComponentBase.LocaleBase.__vfptr);
  FreeMem(ComponentBase.__vfptr);

  FreeAndNil(BinaryDataFactory);

end;


