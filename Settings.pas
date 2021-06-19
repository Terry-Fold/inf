unit Settings;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, FileCtrl, Vcl.Buttons, unit1,
  Vcl.Mask, Vcl.DBCtrls, Vcl.ExtCtrls, Data.DB, Data.Win.ADODB, ADOConnection1,
  ADOQuery1, ADOStoredProc1, Vcl.ComCtrls, Vcl.TabNotBk, Registry, IOUtils, ComObj;
const
  SELDIRHELP = 1000;
type
  TSettingsForm = class(TForm)
    OpenDialog: TOpenDialog;
    Panel1: TPanel;
    ApplyBtn: TBitBtn;
    CloseBtn: TBitBtn;
    DataSourceSet: TDataSource;
    ADOQuerySet: TADOQuery1;
    ADOConnSet: TADOConnection1;
    UndoBtn: TBitBtn;
    SetTabNb: TTabbedNotebook;
    MainSettingsBox: TGroupBox;
    SvodLabel: TLabel;
    WorkFoldLabel: TLabel;
    SvodBtn: TButton;
    WDBtn: TButton;
    GroupBox21: TGroupBox;
    GroupBox22: TGroupBox;
    GroupBox23: TGroupBox;
    smsdispnabtext: TDBMemo;
    GroupBox24: TGroupBox;
    viberdispnabtext: TDBMemo;
    GroupBox25: TGroupBox;
    GroupBox26: TGroupBox;
    smsproftext: TDBMemo;
    GroupBox27: TGroupBox;
    viberproftext: TDBMemo;
    GroupBox28: TGroupBox;
    GroupBox29: TGroupBox;
    smsdisptext: TDBMemo;
    GroupBox30: TGroupBox;
    viberdisptext: TDBMemo;
    ConnStrLabel: TLabel;
    Button3: TButton;
    wdmemo: TMemo;
    svdmemo: TMemo;
    Memo5: TMemo;
    ConnStrTxt: TMemo;
    wd1memo: TMemo;
    WDNetbtn: TButton;
    NetWDCheck: TCheckBox;
    Label1: TLabel;
    Label2: TLabel;
    Svodedit: TEdit;
    workdirnetedit: TEdit;
    workdiredit: TEdit;
    ADOStoredProcSet: TADOStoredProc1;
    wdinfmemo: TMemo;
    procedure WDBtnClick(Sender: TObject);
    function CheckCorrect(isNet:boolean):boolean;
    procedure FormShow(Sender: TObject);
    procedure ApplyBtnClick(Sender: TObject);
    procedure CloseBtnClick(Sender: TObject);
    procedure SvodBtnClick(Sender: TObject);
    procedure UndoBtnClick(Sender: TObject);
    procedure WriteConnStr(str:WideString);
    procedure Button3Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure WDNetbtnClick(Sender: TObject);
    procedure NetWDCheckClick(Sender: TObject);
    Function  CreateWorkDirStructure(path:string):boolean;
    function WriteSettings:boolean;
    function LoadSvod(Filepath:string):boolean;
    procedure LoadSettings;
    procedure SvodeditChange(Sender: TObject);
    procedure workdireditChange(Sender: TObject);
    function CheckConnStr(connStr:string):boolean;
    procedure ConnStrTxtChange(Sender: TObject);
    procedure workdirneteditChange(Sender: TObject);
    procedure CheckWDonNC(pth:string);
    procedure ADOQuerySetFetchProgress(DataSet: TCustomADODataSet; Progress,
      MaxProgress: Integer; var EventStatus: TEventStatus);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  SettingsForm: TSettingsForm;

implementation

{$R *.dfm}

uses StartTask;
{$R files.RES}

procedure TSettingsForm.WriteConnStr(str:WideString);
var
 reg : tregistry;
begin
  reg := tregistry.create;
  try
    reg.RootKey := HKEY_CURRENT_USER;
   if  reg.OpenKey('SOFTWARE\Informing\',true) then
    begin
     reg.WriteString('ConnStr', str);
    end;
  finally
    reg.free;
  end;
end;


procedure TSettingsForm.LoadSettings;
begin
  workdiredit.Text:=adoqueryset.FieldByName('WDpth').AsString;
  workdirnetedit.Text:=adoqueryset.FieldByName('WDNetpth').AsString;
  svodedit.Text:=adoqueryset.FieldByName('Svdpth').AsString;
    if adoqueryset.FieldByName('useNetWD').AsString='1' then
      netwdcheck.Checked:=true
        else netwdcheck.Checked:=false;
        wdinfmemo.Visible:=netwdcheck.Checked;
end;

procedure TSettingsForm.CloseBtnClick(Sender: TObject);
begin
SettingsForm.Close;
end;

procedure TSettingsForm.UndoBtnClick(Sender: TObject);
begin
 adoqueryset.Close;
 adoqueryset.Open;
 LoadSettings;
 ConnStrTxt.Text:=ADOConnSet.ConnectionString;
end;

procedure TSettingsForm.Button3Click(Sender: TObject);
var cnstr:widestring;
begin
//cnstr:= adoconnset.ConnectionString;
 with starttaskform do
  begin
       ADOConnStart.Close;
       ADOQuerystart.Close;
       ADOConnStart.ConnectionString:=connstrtxt.Text;
        try
          ADOConnStart.Open;
          if ADOConnStart.Connected then messagedlg( '����� � SQL �������� ������� �����������!' ,mtInformation, [mbOk], 0, mbOk)
            else  messagedlg( '��� ����� � ��! ���������� �� �������!' , mtError, [mbOk], 0, mbOk);
        except
          messagedlg( '��� ����� � ��! �������������� ������!' , mtError, [mbOk], 0, mbOk);
       end;
  end;
 end;

function TSettingsForm.CheckConnStr(connStr:string):boolean;
begin
 with starttaskform do
  begin
       ADOConnStart.Close;
       ADOQuerystart.Close;
       ADOConnStart.ConnectionString:=connStr;
        try
          ADOConnStart.Open;
          if ADOConnStart.Connected then
                            begin
                             ConnStrTxt.Color:=clWindow;
                             result:=true;
                            end
                              else begin
                               ConnStrTxt.Color:=clred;
                               result:=false;
                              end;

        except
          ConnStrTxt.Color:=clred;
          result:=false;
        end;
      end
  end;

procedure TSettingsForm.CheckWDonNC(pth:string);
begin
{
with starttaskform do
  begin
       ADOQuerystart.Close;
       ADOQuerystart.SQL.Clear;
       ADOQuerystart.SQL='';
end;                              }
end;


function TSettingsForm.CheckCorrect(isNet:boolean):boolean;
begin
result:=true;
//���� � ����� ���� �������������
   if not FileExists(Svodedit.Text) then
    begin
      Svodedit.Color:=clred;
      result:=false;
    end
  else Svodedit.Color:=clWindow;
//���� � ������� �����
 if isNet then
  begin
  if workdiredit.text='' then
     begin
      workdiredit.Color:=clred;
      result:=false;
     end
    else  workdiredit.Color:=clWindow;
   end
    else begin
      if not DirectoryExists(workdiredit.text) then
       begin
         workdiredit.Color:=clred;
         result:=false;
        end
          else workdiredit.Color:=clWindow;
    end;
 //������� ���� � ������� ����� (���� ������������)
   if netwdcheck.Checked then
    begin
     if not DirectoryExists(workdirnetedit.text) then
      begin
       workdirnetedit.Color:=clred;
      result:=false;
      end
     else workdirnetedit.Color:=clWindow;
    end;

 //����� ���
  //�����
   if smsdisptext.Lines.Count=0 then
    begin
      smsdisptext.color := clred;
      result:=false;
    end
      else smsdisptext.color:=clWindow;
  //�����
   if smsproftext.Lines.Count=0 then
    begin
      smsproftext.color := clred;
      result:=false;
    end
      else smsproftext.color:=clWindow;
  //��������
   if smsdispnabtext.Lines.Count=0 then
    begin
      smsdispnabtext.color := clred;
      result:=false;
    end
      else smsdispnabtext.color:=clWindow;
//����� Viber
 //�����
  if viberdisptext.Lines.Count=0 then
    begin
      viberdisptext.color := clred;
      result:=false;
    end
      else viberdisptext.color:=clWindow;
 //�����
   if viberproftext.Lines.Count=0 then
    begin
      viberproftext.color := clred;
      result:=false;
    end
      else viberproftext.color:=clWindow;
 //��������
   if viberdispnabtext.Lines.Count=0 then
    begin
      viberdispnabtext.color := clred;
      result:=false;
    end
      else viberdispnabtext.color:=clWindow;
 //������ �����������
  result:=CheckConnStr(ConnStrTxt.Text);
end;

procedure TSettingsForm.ConnStrTxtChange(Sender: TObject);
begin
  CheckConnStr(ConnStrTxt.Text);
end;

Function TSettingsForm.CreateWorkDirStructure(path:string):boolean;
var
_res: TResourceStream;
begin
  try
//�������� ��������� �����
    TDirectory.CreateDirectory(path+'\�������\� �������');
    TDirectory.CreateDirectory(path+'\�������\�� �����');
    TDirectory.CreateDirectory(path+'\�������\����');
    TDirectory.CreateDirectory(path+'\����������');
 //����������� ����� "���� �������������"
   if not CopyFile(PChar(Svodedit.Text), PChar(path+'\����������\���� �������������'+ExtractFileExt(svodedit.Text)), false)
   then  messagedlg( '���������� ����������� ���� "���� �������������" � ������� �����. '+SysErrorMessage(GetLastError) , mtError, [mbOk], 0, mbOk);
 //���������� �������� �� �������� � �������� �����
    _res:=TResourceStream.Create(HInstance,'ToIB','XLSX');
    _res.SaveToFile(path+'\�������\� �������\To_Ib.xlsx');
    _res:=TResourceStream.Create(HInstance,'ToPost','XLSX');
    _res.SaveToFile(path+'\�������\�� �����\To_post.xlsx');
    _res:=TResourceStream.Create(HInstance,'Recon','DOCX');
    _res.SaveToFile(path+'\�������\����\Recon.docx');
    _res:=TResourceStream.Create(HInstance,'Kds','XLSX');
    _res.SaveToFile(path+'\�������\����\KDS.xlsx');
   result:=true;
  except
    result:=false;
  end;
end;

function TSettingsForm.WriteSettings:boolean;
begin
  try
      with mainform do
         begin
            ADOStoredProc.ProcedureName:='edit_Settings';
            ADOStoredProc.Parameters.Refresh;
            ADOStoredProc.Parameters.ParamByName('@SMSDispTxt').Value:=smsdisptext.Text;
            ADOStoredProc.Parameters.ParamByName('@SMSProfTxt').Value:=smsproftext.Text;
            ADOStoredProc.Parameters.ParamByName('@SMSDispNabTxt').Value:=smsdispnabtext.Text;
            ADOStoredProc.Parameters.ParamByName('@ViberDispTxt').Value:=viberdisptext.Text;
            ADOStoredProc.Parameters.ParamByName('@ViberProfTxt').Value:= viberproftext.Text;
            ADOStoredProc.Parameters.ParamByName('@ViberDispNabTxt').Value:=viberdispnabtext.Text;
            ADOStoredProc.Parameters.ParamByName('@WDpth').Value:= workdiredit.Text;
            ADOStoredProc.Parameters.ParamByName('@WDNetpth').Value:= workdirnetedit.Text;
            ADOStoredProc.Parameters.ParamByName('@Svdpth').Value:=svodedit.Text;
             if NetWDCheck.Checked then
              ADOStoredProc.Parameters.ParamByName('@useNetWD').Value:='1'
               else ADOStoredProc.Parameters.ParamByName('@useNetWD').Value:='0';
            ADOStoredProc.ExecProc;
         end;
    result:=true;
  except
    result:=false;
  end;
end;

function TSettingsForm.LoadSvod(Filepath:string):boolean;
begin
  try
    ADOStoredProcSet.ProcedureName:='Load_Svod';
    ADOStoredProcSet.Parameters.Refresh;
    ADOStoredProcSet.Parameters.ParamByName('@FilePath').Value:=Filepath;
    ADOStoredProcSet.ExecProc;
   result:=true;
  except result:=false; end;
end;

procedure TSettingsForm.ADOQuerySetFetchProgress(DataSet: TCustomADODataSet;
  Progress, MaxProgress: Integer; var EventStatus: TEventStatus);
begin
  mainform.ProgressBar.Max:=MaxProgress;
  application.ProcessMessages;
  mainform.ProgressBar.Position:= Progress;
end;

procedure TSettingsForm.ApplyBtnClick(Sender: TObject);
var wdpth:string;
begin
//���� �� ������� SQL ���������� ������� ����
 if netwdcheck.Checked then wdpth:=workdirnetedit.Text
//����� - ���������
 else wdpth:=workdiredit.Text;

//���� ���� ����������� � ����
if MainForm.CheckDBConn then
  begin
    if CheckCorrect(netwdcheck.Checked) then
          begin
         //�������� ������� �����
               if  not CreateWorkDirStructure(wdpth)  then
                begin
                 messagedlg( '���������� ������� ��������� �����!' , mtError, [mbOk], 0, mbOk);
                 MainForm.StatusBar.Panels[0].Text:='�������� ������ ��� ��������� ��������.';
                 abort;
                end;
         //�������� ����� ���� ������������� � ����
                if not LoadSvod(workdiredit.Text+'\����������\���� �������������'+ExtractFileExt(svodedit.Text))
                 then  begin
                   messagedlg( '���������� ��������� "���� �������������" � ����.' , mtError, [mbOk], 0, mbOk);
                   MainForm.StatusBar.Panels[0].Text:='�������� ������ ��� ��������� ��������.';
                   abort;
                 end;
         //������ � ��
              if not WriteSettings then begin
                 messagedlg( '��� ���������� �������� �������� ������!' , mtError, [mbOk], 0, mbOk);
                 MainForm.StatusBar.Panels[0].Text:='�������� ������ ��� ��������� ��������.';
                 abort;
              end;
         //������ � ������

                 try
                   WriteConnStr(ConnStrTxt.Text);
                 except  messagedlg( '���������� ��������� �������� � ������!' , mtError, [mbOk], 0, mbOk);
                         MainForm.StatusBar.Panels[0].Text:='�������� ������ ��� ��������� ��������.';
                         abort;
                 end;

              if netwdcheck.Checked then
                 adoqueryset.FieldByName('useNetWD').AsString;

          if  messagedlg( '��������� ������� ���������.'#13#10'������ ������� ����?' , mtConfirmation,  [mbYes,mbNo], 0)=6 //��������� ������
                then  settingsform.Close;

             MainForm.StatusBar.Panels[0].Text:='��������� �������� ���������� ��������� �������!'   //������ ������ �� ������ ������� �����

             end
           else //���� �������� �� ������������ ���������� �� ��������
             begin
             messagedlg( '��������� ������ ����� �����������!' , mtError, [mbOk], 0, mbOk);
             MainForm.StatusBar.Panels[0].Text:='�������� ������ ��� ��������� ��������.'
             {
               try
                 WriteConnStr(ConnStrTxt.Text);
                except messagedlg( '���������� ��������� �������� � ������!' , mtError, [mbOk], 0, mbOk);
                end;
            if not WriteSettings then
                 messagedlg( '���������� ��������� ���������. �������������� ������!' , mtError, [mbOk], 0, mbOk);
                    }
         //  SettingsForm.Close;   //������� ����
         //  MainForm.StatusBar.Panels[0].Text:='��������� ������� ���������.'  //������ ������ �� ������ ������� �����

            end;

  // SettingsForm.Close;
  end
  //���� ��� ����������� � ��
    else
      begin
        try
          WriteConnStr(ConnStrTxt.Text);
        except  messagedlg( '���������� ��������� �������� � ������!' , mtError, [mbOk], 0, mbOk);
        end;

         messagedlg( '��������� ������� ���������. ������ ������� ����?' , mtConfirmation,  [mbYes,mbNo], 0);
      end;
 end;




procedure TSettingsForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
with mainform do begin
if CheckDBConn then begin
                      statmemo.Clear;
                     AddColoredLine(statmemo,'���������� � �� - �������!', clGreen)
                    end
                      else begin
                        statmemo.Clear;
                        infmemo.Clear;
                        infsstatmemo.Clear;
                        infsstatmemo.Lines.Add('���������� �������� ����������!');
                        AddColoredLine(statmemo,'��� ����� � ��! ������!', clRed);
                        infmemo.Lines.Add('��� ������!');
                      end;
                 checksettings;
                  end;
end;


procedure TSettingsForm.FormShow(Sender: TObject);
begin
 //���� ������� �������� �������� ������ ����������� �� �������
  if MainForm.ReadConnStr<>'not_' then
   begin
    AdoConnSet.Close;
    ADOConnSet.ConnectionString:=MainForm.ReadConnStr;
      try
        ADOConnSet.Open;
       except messagedlg( '������ ����������� � �� �����������!'+#13#10+' ��� ����� � ��������!' , mtError, [mbOk], 0, mbOk)
      end;
       //���������� ������ ����������� � memo
          ConnStrTxt.Text:= MainForm.ReadConnStr;
   end
   //���� �� ������� �������� �������� ������ ����������� �� �������
    else
      begin
        AdoConnSet.Close;
        ConnStrTxt.Text:=ADOConnSet.ConnectionString;
      end;
//��������� ���������� � �����
  try
    adoqueryset.Open;
  except
  end;
   //��������� ���������
     LoadSettings;
   //��������� ������������
     CheckCorrect(NetWDCheck.Checked);

{
smsdisptext.DataField:='smsdisptxt';
smsproftext.DataField:='smsproftxt';
smsdispnabtext.DataField:='smsdispnabtxt';
viberdisptext.DataField:='viberdisptxt';
viberproftext.DataField:='viberproftxt';
viberdispnabtext.DataField:='viberdispnabtxt';
}

end;

procedure TSettingsForm.NetWDCheckClick(Sender: TObject);
begin
workdirnetedit.Enabled := NetWDCheck.Checked;
wdnetbtn.Enabled :=  NetWDCheck.Checked;
wdinfmemo.Visible :=  NetWDCheck.Checked;
if NetWDCheck.Checked then workdiredit.Color:=clWindow
  else  begin
   if not DirectoryExists(workdiredit.text) then
        workdiredit.Color:=clred
         else workdiredit.Color:=clWindow;
  end;
end;

procedure TSettingsForm.WDBtnClick(Sender: TObject);
var  OpenDialog: TFileOpenDialog;
begin
 OpenDialog := TFileOpenDialog.Create(MainForm);
try
  OpenDialog.Options := OpenDialog.Options + [fdoPickFolders];
  if not OpenDialog.Execute then
    Abort;
  adoqueryset.Edit;
  workdiredit.Text := OpenDialog.FileName;
  if NetWDCheck.Checked then
  begin
  if workdiredit.text='' then
     workdiredit.Color:=clred
    else  workdiredit.Color:=clWindow;
   end
    else begin
      if not DirectoryExists(workdiredit.text) then
        workdiredit.Color:=clred
          else workdiredit.Color:=clWindow;
    end;
finally
  OpenDialog.Free;
end;


end;

procedure TSettingsForm.WDNetbtnClick(Sender: TObject);
var OpenDialog: TFileOpenDialog;
begin
OpenDialog := TFileOpenDialog.Create(MainForm);
try
  OpenDialog.Options := OpenDialog.Options + [fdoPickFolders];
  if not OpenDialog.Execute then
    Abort;
  adoqueryset.Edit;
  workdirnetedit.Text := OpenDialog.FileName;
finally
  OpenDialog.Free;
end;
end;



Procedure TSettingsForm.SvodBtnClick(Sender: TObject);
begin
if opendialog.Execute then
 begin
  svodedit.Text:=opendialog.Filename;
  if not FileExists(Svodedit.Text) then
    Svodedit.Color:=clred
      else Svodedit.Color:=clWindow;

 end;
end;


procedure TSettingsForm.SvodeditChange(Sender: TObject);
begin
  if not FileExists(Svodedit.Text) then
    Svodedit.Color:=clred
      else Svodedit.Color:=clWindow;
end;

procedure TSettingsForm.workdireditChange(Sender: TObject);
begin
if not NetWDCheck.Checked then
  begin
   if not DirectoryExists(workdiredit.text) then
        workdiredit.Color:=clred
         else workdiredit.Color:=clWindow;
  end;
end;


procedure TSettingsForm.workdirneteditChange(Sender: TObject);
begin

   if netwdcheck.Checked then
    begin
     if not DirectoryExists(workdirnetedit.text) then
       workdirnetedit.Color:=clred
     else workdirnetedit.Color:=clWindow;
    end;
end;

end.



