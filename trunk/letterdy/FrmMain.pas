unit FrmMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, DB, ADODB, RpDefine, RpCon, RpConDS, Grids,
  DBGrids, RpRave, RpBase, RpSystem, ExtCtrls, JPEG, RvCsRpt,RVProj,RVClass,
  RVCsStd, StrUtils;

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
    lblTerm: TLabel;
    edtTerm: TEdit;
    lblDate: TLabel;
    edtDate: TEdit;
    memContent: TMemo;
    lblContent: TLabel;
    lblFirstName: TLabel;
    edtFirst: TEdit;
    lblSecondName: TLabel;
    edtSecond: TEdit;
    lblTeacher: TLabel;
    edtTeacher: TEdit;
    lblNumber: TLabel;
    Label1: TLabel;
    edtStuEva: TEdit;
    Label2: TLabel;
    edtAct: TEdit;
    Edit1: TEdit;
    Label3: TLabel;
    Label4: TLabel;
    lblWord: TLabel;
    memWord: TMemo;
    procedure btnExcelClick(Sender: TObject);
    procedure btnAccessClick(Sender: TObject);
    procedure btnTransferClick(Sender: TObject);
    procedure btnDatabaseClick(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure memContentChange(Sender: TObject);
    procedure memWordChange(Sender: TObject);
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
var
  MyPage: TRavePage;
  MyText1: TRaveText;
  MyText2: TRaveText;
  MyText3: TRaveText;
  MyText4: TRaveText;
  MyText5: TRaveText;
  MyMemo: TRaveMemo;
  MyMemo1: TRaveMemo;
  MyMemo2: TRaveMemo;
  MyMemo3: TRaveMemo;
  MyMemo4: TRaveMemo;
begin
  if ((Trim(edtDatabase.Text) <> '') and (Trim(cbbFirst.Text) <> '') and (Trim(cbbSecond.Text) <> '')
      and (Trim(edtTerm.Text) <> '') and (Trim(edtDate.Text) <> '') and (Trim(memContent.Text) <> '')
      and (Trim(edtTeacher.Text) <> '') and (Trim(edtFirst.Text) <> '') and (Trim(edtSecond.Text) <> '')
      and (Trim(edtStuEva.Text) <> '') and (Trim(edtAct.Text) <> '') and (Trim(Edit1.Text) <> '')
      and (Trim(memWord.Text) <> '')) then
  begin
    rvpMain.Open;
    MyPage := rvpMain.ProjMan.FindRaveComponent('report.page', nil) as TRavePage;
    MyText1 := rvpMain.ProjMan.FindRaveComponent('Text17', MyPage) as TRaveText;
    MyText2 := rvpMain.ProjMan.FindRaveComponent('Text15', MyPage) as TRaveText;
    MyText3 := rvpMain.ProjMan.FindRaveComponent('Text18', MyPage) as TRaveText;
    MyText4 := rvpMain.ProjMan.FindRaveComponent('Text20', MyPage) as TRaveText;
    MyText5 := rvpMain.ProjMan.FindRaveComponent('Text22', MyPage) as TRaveText;
    MyMemo :=  rvpMain.ProjMan.FindRaveComponent('Memo11', MyPage) as TRaveMemo;
    MyMemo1 := rvpMain.ProjMan.FindRaveComponent('Memo8', MyPage) as TRaveMemo;
    MyMemo2 := rvpMain.ProjMan.FindRaveComponent('Memo9', MyPage) as TRaveMemo;
    MyMemo3 := rvpMain.ProjMan.FindRaveComponent('Memo10', MyPage) as TRaveMemo;
    MyMemo4 := rvpMain.ProjMan.FindRaveComponent('Memo5', MyPage) as TRaveMemo;
    MyText1.Text := Trim(edtTerm.Text);
    MyText2.Text := Trim(edtDate.Text);
    MyText3.Text := Trim(edtTeacher.Text);
    MyText4.Text := Trim(edtFirst.Text);
    MyText5.Text := Trim(edtSecond.Text);
    MyMemo.Text := Trim(memContent.Text);
    MyMemo1.Text := Trim(edtStuEva.Text);
    MyMemo2.Text := Trim(edtAct.Text);
    MyMemo3.Text := Trim(Edit1.Text);
    MyMemo4.Text := '    ' + Trim(memWord.Text);
    try
      adocMain.Connected := false;
      adocMain.ConnectionString :=
          'Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+edtDatabase.Text+';'
          +'Persist Security Info=False';
      qryTable01.SQL.Clear;
      qryTable01.SQL.Add('SELECT * FROM ' + cbbFirst.Text + ' AS test02, ' + cbbSecond.Text + ' AS test03');
      qryTable01.SQL.Add(' WHERE test02.����=test03.����');
      qryTable01.ExecSQL;
      rvpMain.Execute;
      adocMain.Connected := false;
    except
      // �����쳣������ת��ʧ��
      adocMain.Connected := false;
      ShowMessage('���ܽ��д�ӡ��������������!');
      exit;
    end;
  end
  else
  begin
    ShowMessage('��������������!');
  end;
end;

procedure TfrmMainFrame.btnTransferClick(Sender: TObject);
var
  SQLStr:string;
begin
  if ((Trim(edtExcel.Text) <> '')and(Trim(edtSheet.Text) <> '')and(Trim(edtAccess.Text) <> '')and(Trim(edtTable.Text) <> '')) then
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

procedure TfrmMainFrame.FormCreate(Sender: TObject);
begin
    Image1.Picture.LoadFromFile('home.jpg');
    memWord.Text := '�ֽ�ѧ����ѧϰ�ɼ��������ȥ����һ�ģ�ͬʱ������ı�������ͽ��������±��Ա����ǸĽ���������ѧУ��ø��á�';
    memContent.Text := '����֪ƴ���Ż�ɹ�����ֲŻ�ʤ����������ú�������ٵ�ÿһ���ӣ�ĪҪ����ѧ��ʱ��ź��ʱ����š����Ǵ����ģ�����Ҫ���ӿ̿࣬��ͣ������ǰ��ɣ��ҿ����㡣��Ҫ��ֵ��ף�������ǻۺ�������ս���ң��������ң�����������ļ��ޣ�';
end;

procedure TfrmMainFrame.memContentChange(Sender: TObject);
var
  num: Integer;
begin
  num := Length(memContent.Text) div 2;
  lblNumber.Caption := IntToStr(num) + '/120';
end;

procedure TfrmMainFrame.memWordChange(Sender: TObject);
var
  num: Integer;
begin
  num := Length(memWord.Text) div 2;
  lblWord.Caption := IntToStr(num) + '/80';
end;

end.
