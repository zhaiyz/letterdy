program ProLetterDy;

uses
  Forms,
  FrmMain in 'FrmMain.pas' {frmMainFrame};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := '�ҳ�֪ͨ���ӡV0.4 beta';
  Application.CreateForm(TfrmMainFrame, frmMainFrame);
  Application.Run;
end.
