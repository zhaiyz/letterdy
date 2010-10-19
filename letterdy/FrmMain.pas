unit FrmMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, DB, ADODB;

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
    adocMain: TADOConnection;
    procedure btnExcelClick(Sender: TObject);
    procedure btnAccessClick(Sender: TObject);
    procedure btnTransferClick(Sender: TObject);
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

procedure TfrmMainFrame.btnExcelClick(Sender: TObject);
begin
  if dlgExcel.Execute then
  begin
    edtExcel.Text := dlgExcel.FileName
  end;
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
      // 出现异常，数据转移失败
      ShowMessage('数据转换失败，请检查输入数据!');
      exit;
    end;
      ShowMessage('数据转换成功，请继续其他操作!');
  end
else
  begin
    ShowMessage('请输入完整数据!');
  end;
end;

end.
