unit FrmMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, DB, ADODB, RpDefine, RpCon, RpConDS, Grids,
  DBGrids, RpRave;

type
  TfrmMainFrame = class(TForm)
    pcMain: TPageControl;
    tsTransfer: TTabSheet;
    tsPrint: TTabSheet;
    tsHome: TTabSheet;
    lblExcel: TLabel;
    edtExcel: TEdit;
    lblSheet: TLabel;
    edtSheet: TEdit;
    btnExcel: TButton;
    lblAccess: TLabel;
    edtAccess: TEdit;
    lblTable: TLabel;
    edtTable: TEdit;
    btnAccess: TButton;
    btnTransfer: TButton;
    dlgExcel: TOpenDialog;
    dlgAccess: TOpenDialog;
    lblFirst: TLabel;
    cbbFirst: TComboBox;
    lblSecond: TLabel;
    cbbSecond: TComboBox;
    lblDatabase: TLabel;
    edtDatabase: TEdit;
    btnDatabase: TButton;
    dlgDatabase: TOpenDialog;
    rvdcTable01: TRvDataSetConnection;
    rvpMain: TRvProject;
    qryTable01: TADOQuery;
    qryTable02: TADOQuery;
    btnPrint: TButton;
    adocMain: TADOConnection;
    rvdcTable02: TRvDataSetConnection;
    procedure btnExcelClick(Sender: TObject);
    procedure btnAccessClick(Sender: TObject);
    procedure btnTransferClick(Sender: TObject);
    procedure btnDatabaseClick(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMainFrame: TfrmMainFrame;

implementation

{$R *.dfm}

procedure TfrmMainFrame.btnAccessClick(Sender: TObject);
begin
  if dlgAccess.Execute then
  begin
    edtAccess.Text := dlgAccess.FileName;
  end;
end;

procedure TfrmMainFrame.btnDatabaseClick(Sender: TObject);
begin
  // ѡ�����ݿ⣬����ʾ�����ݿ�����ı�
  if dlgDatabase.Execute then
  begin
    edtDatabase.Text := dlgDatabase.FileName;
    adocMain.Connected := false;
    adocMain.ConnectionString :=
      'Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+edtDatabase.Text+';'
      +'Persist Security Info=False';
    adocMain.GetTableNames(cbbFirst.Items, false);
    adocMain.GetTableNames(cbbSecond.Items, false);
    adocMain.Connected := false;
  end;
end;

procedure TfrmMainFrame.btnExcelClick(Sender: TObject);
begin
  if dlgExcel.Execute then
  begin
    edtExcel.Text := dlgExcel.FileName;
  end;
end;

procedure TfrmMainFrame.btnPrintClick(Sender: TObject);
begin
  adocMain.Connected := false;
  adocMain.ConnectionString :=
      'Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+edtDatabase.Text+';'
      +'Persist Security Info=False';
  qryTable01.SQL.Clear;
  qryTable01.SQL.Add('SELECT * FROM ' + cbbFirst.Text);
  qryTable01.ExecSQL;
  qryTable02.SQL.Clear;
  qryTable02.SQL.Add('SELECT * FROM ' + cbbSecond.Text);
  qryTable02.ExecSQL;
  rvpMain.Execute;
  adocMain.Connected := false;
end;

procedure TfrmMainFrame.btnTransferClick(Sender: TObject);
var
  SQLStr:string;
begin
  if ((edtExcel.Text <> '')and(edtSheet.Text <> '')and(edtAccess.Text <> '')and(edtTable.Text <> '')) then
  begin
    try
      adocMain.Connected := false;
      adocMain.ConnectionString :=
        'Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+edtAccess.Text+';'
        +'Persist Security Info=False';
      SQLStr:='select * into '+edtTable.Text+' FROM [excel 8.0;database='
        +edtExcel.Text+'].['+edtSheet.Text+'$]';
      adocMain.Execute(SQLStr);
    except
      // �����쳣������ת��ʧ��
      adocMain.Connected := false;
      ShowMessage('����ת��ʧ�ܣ�������������!');
      exit;
    end;
      adocMain.Connected := false;
      ShowMessage('����ת���ɹ����������������!');
  end
else
  begin
    ShowMessage('��������������!');
  end;
end;

end.