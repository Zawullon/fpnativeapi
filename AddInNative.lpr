library AddInNative;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX} {$IFDEF UseCThreads} cthreads, cmem, {$ENDIF} {$ENDIF}
  V8AddIn, V8AddInGUI, Example, EasyExample, ExampleForm, GUIExample;

exports

  V8AddIn.GetClassNames, V8AddIn.GetClassObject, V8AddIn.DestroyObject, V8AddIn.SetPlatformCapabilities;

{$R *.res}

begin

  RegisterAddInClass(TAddInExample);

  with TAddInEasyExample do
    begin
      RegisterAddInClass('AddInNativeEasyExtension');
      AddProp('IsEnabled', 'Включен', @GetIsEnabled, @SetIsEnabled);
      AddProp('IsTimerPresent', 'ЕстьТаймер', @GetIsTimerPresent);
      AddProc('Enable', 'Включить', @Enable);
      AddProc('Disable', 'Выключить', @Disable);
      AddProc('ShowInStatusLine', 'ПоказатьВСтрокеСтатуса', @ShowInStatusLine);
      AddProc('StartTimer', 'СтартТаймер', @StartTimer);
      AddProc('StopTimer', 'СтопТаймер', @StopTimer);
      AddFunc('LoadPicture', 'ЗагрузитьКартинку', @LoadPicture);
      AddProc('ShowMessageBox', 'ПоказатьСообщение', @ShowMessageBox);
    end;

  with TFormController do
    begin
      RegisterAddInClass('AddInNativeGUIExtension');
      AddProc('Show', 'Показать', @Show);
      AddProc('Hide', 'Скрыть', @Hide);
      AddProc('SetText', 'УстановитьТекст', @SetText);
    end;

end.

