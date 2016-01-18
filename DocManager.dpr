program DocManager;

uses
  Forms, Classes, Windows,
  UMain in 'UMain.pas' {MainForm},
  USplash in 'USplash.pas' {SplashForm};

{$R *.res}

begin
  ShowWindow(Application.Handle, SW_Hide);
  Application.Initialize;
  Application.Title := 'Doc Manager';
  SplashForm := TSplashForm.Create(Application);
  SplashForm.Show;
  SplashForm.Update;
  Application.CreateForm(TMainForm, MainForm);
  SplashForm.Hide;
  SplashForm.Free;
  Application.Run;
end.
