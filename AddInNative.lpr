library AddInNative;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX} {$IFDEF UseCThreads} cthreads, cmem, {$ENDIF} {$ENDIF}
  V8AddIn, Example;

exports

  V8AddIn.GetClassNames, V8AddIn.GetClassObject, V8AddIn.DestroyObject, V8AddIn.SetPlatformCapabilities;

end.

