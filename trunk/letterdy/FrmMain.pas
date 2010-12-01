unit FrmMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, DB, ADODB, RpDefine, RpCon, RpConDS, Grids,
  DBGrids, RpRave, RpBase, RpSystem, ExtCtrls, JPEG;

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
    btnPrint: TButton;
    adocMain: TADOConnection;
    Image1: TImage;
    procedure btnExcelClick(Sender: TObject);
    procedure btnAccessClick(Sender: TObject);
    procedure btnTransferClick(Sender: TObject);
    procedure btnDatabaseClick(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
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
  // 选择数据库，并显示出数据库里面的表
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
  if ((edtDatabase.Text <> '') and (cbbFirst.Text <> '') and (cbbSecond.Text <> '')) then
  begin
    try
      adocMain.Connected := false;
      adocMain.ConnectionString :=
          'Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+edtDatabase.Text+';'
          +'Persist Security Info=False';
      qryTable01.SQL.Clear;
      qryTable01.SQL.Add('SELECT * FROM ' + cbbFirst.Text + ' AS test02, ' + cbbSecond.Text + ' AS test03');
      qryTable01.SQL.Add(' WHERE test02.考号=test03.考号');
      qryTable01.ExecSQL;
      rvpMain.Execute;
      adocMain.Connected := false;
    except
      // 出现异常，数据转移失败
      adocMain.Connected := false;
      ShowMessage('不能进行打印，请检查输入数据!');
      exit;
    end;
  end
  else
  begin
    ShowMessage('请输入完整数据!');
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
      adocMain.Connected := false;
      ShowMessage('数据转换失败，请检查输入数据!');
      exit;
    end;
      adocMain.Connected := false;
      ShowMessage('数据转换成功，请继续其他操作!');
  end
  else
  begin
    ShowMessage('请输入完整数据!');
  end;
end;

procedure TfrmMainFrame.FormCreate(Sender: TObject);
begin
    Image1.Picture.LoadFromFile('home.jpg');
end;

end.
