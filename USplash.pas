unit USplash;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Buttons, ShellApi;

type
  TSplashForm = class(TForm)
    Panel1: TPanel;
    Label2: TLabel;
    Bevel1: TBevel;
    LabLoad: TLabel;
    BtnClose: TBitBtn;
    Image1: TImage;
    Label3: TLabel;
    Label4: TLabel;
    procedure Label3Click(Sender: TObject);
  end;

var
  SplashForm: TSplashForm;

implementation

{$R *.dfm}

procedure TSplashForm.Label3Click(Sender: TObject);
begin
  ShellExecute(0, 'open', Pchar('mailto:benbac20@gmail.com'), nil, nil,  sw_show);
end;

end.
