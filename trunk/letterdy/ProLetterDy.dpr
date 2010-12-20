program ProLetterDy;

uses
  Forms,
  FrmMain in 'FrmMain.pas' {frmMainFrame};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := '家长通知书打印V0.4 beta';
  Application.CreateForm(TfrmMainFrame, frmMainFrame);
  Application.Run;
end.
