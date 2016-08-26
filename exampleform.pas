unit ExampleForm;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls;

type

  { TFormExample }

  TFormExample = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Label1: TLabel;
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  FormExample: TFormExample;

implementation

{$R *.lfm}

end.

