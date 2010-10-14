program ProLetterDy;

uses
  Forms,
  FrmMain in 'FrmMain.pas' {frmMainFrame};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmMainFrame, frmMainFrame);
  Application.Run;
end.
