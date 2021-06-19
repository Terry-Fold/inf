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
          if ADOConnStart.Connected then messagedlg( 'Связь с SQL сервером успешно установлена!' ,mtInformation, [mbOk], 0, mbOk)
            else  messagedlg( 'Нет связи с БД! Соединение не удалось!' , mtError, [mbOk], 0, mbOk);
        except
          messagedlg( 'Нет связи с БД! Непредвиденная ошибка!' , mtError, [mbOk], 0, mbOk);
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
//Путь к файлу Свод маршрутизации
   if not FileExists(Svodedit.Text) then
    begin
      Svodedit.Color:=clred;
      result:=false;
    end
  else Svodedit.Color:=clWindow;
//Путь к рабочей папке
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
 //Сетевой путь к рабочей папке (если используется)
   if netwdcheck.Checked then
    begin
     if not DirectoryExists(workdirnetedit.text) then
      begin
       workdirnetedit.Color:=clred;
      result:=false;
      end
     else workdirnetedit.Color:=clWindow;
    end;

 //Текст СМС
  //Диспы
   if smsdisptext.Lines.Count=0 then
    begin
      smsdisptext.color := clred;
      result:=false;
    end
      else smsdisptext.color:=clWindow;
  //Профы
   if smsproftext.Lines.Count=0 then
    begin
      smsproftext.color := clred;
      result:=false;
    end
      else smsproftext.color:=clWindow;
  //ДиспНабы
   if smsdispnabtext.Lines.Count=0 then
    begin
      smsdispnabtext.color := clred;
      result:=false;
    end
      else smsdispnabtext.color:=clWindow;
//Текст Viber
 //Диспы
  if viberdisptext.Lines.Count=0 then
    begin
      viberdisptext.color := clred;
      result:=false;
    end
      else viberdisptext.color:=clWindow;
 //Профы
   if viberproftext.Lines.Count=0 then
    begin
      viberproftext.color := clred;
      result:=false;
    end
      else viberproftext.color:=clWindow;
 //ДиспНабы
   if viberdispnabtext.Lines.Count=0 then
    begin
      viberdispnabtext.color := clred;
      result:=false;
    end
      else viberdispnabtext.color:=clWindow;
 //Строка подключение
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
//Создание структуры папок
    TDirectory.CreateDirectory(path+'\Шаблоны\В Инфобип');
    TDirectory.CreateDirectory(path+'\Шаблоны\на Почту');
    TDirectory.CreateDirectory(path+'\Шаблоны\Доки');
    TDirectory.CreateDirectory(path+'\Оповещения');
 //Копирование файла "Свод маршрутизации"
   if not CopyFile(PChar(Svodedit.Text), PChar(path+'\Оповещения\Свод маршрутизации'+ExtractFileExt(svodedit.Text)), false)
   then  messagedlg( 'Невозможно скопировать файл "Свод маршрутизации" в рабочуу папку. '+SysErrorMessage(GetLastError) , mtError, [mbOk], 0, mbOk);
 //распаковка шаблонов из ресурсов в рабочуую папку
    _res:=TResourceStream.Create(HInstance,'ToIB','XLSX');
    _res.SaveToFile(path+'\Шаблоны\В Инфобип\To_Ib.xlsx');
    _res:=TResourceStream.Create(HInstance,'ToPost','XLSX');
    _res.SaveToFile(path+'\Шаблоны\на Почту\To_post.xlsx');
    _res:=TResourceStream.Create(HInstance,'Recon','DOCX');
    _res.SaveToFile(path+'\Шаблоны\Доки\Recon.docx');
    _res:=TResourceStream.Create(HInstance,'Kds','XLSX');
    _res.SaveToFile(path+'\Шаблоны\Доки\KDS.xlsx');
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
//Если на сервере SQL используем сетевой путь
 if netwdcheck.Checked then wdpth:=workdirnetedit.Text
//Иначе - локальный
 else wdpth:=workdiredit.Text;

//Если есть подключение к базе
if MainForm.CheckDBConn then
  begin
    if CheckCorrect(netwdcheck.Checked) then
          begin
         //Создание рабочей папки
               if  not CreateWorkDirStructure(wdpth)  then
                begin
                 messagedlg( 'Невозможно создать структуру папок!' , mtError, [mbOk], 0, mbOk);
                 MainForm.StatusBar.Panels[0].Text:='Возникла ошибка при изменении настроек.';
                 abort;
                end;
         //Загрузка файла Свод маршрутизации в базу
                if not LoadSvod(workdiredit.Text+'\Оповещения\Свод маршрутизации'+ExtractFileExt(svodedit.Text))
                 then  begin
                   messagedlg( 'Невозможно загрузить "Свод маршрутизации" в базу.' , mtError, [mbOk], 0, mbOk);
                   MainForm.StatusBar.Panels[0].Text:='Возникла ошибка при изменении настроек.';
                   abort;
                 end;
         //Запись в БД
              if not WriteSettings then begin
                 messagedlg( 'При сохранении настроек возникла ошибка!' , mtError, [mbOk], 0, mbOk);
                 MainForm.StatusBar.Panels[0].Text:='Возникла ошибка при изменении настроек.';
                 abort;
              end;
         //Запись в реестр

                 try
                   WriteConnStr(ConnStrTxt.Text);
                 except  messagedlg( 'Невозможно сохранить значение в реестр!' , mtError, [mbOk], 0, mbOk);
                         MainForm.StatusBar.Panels[0].Text:='Возникла ошибка при изменении настроек.';
                         abort;
                 end;

              if netwdcheck.Checked then
                 adoqueryset.FieldByName('useNetWD').AsString;

          if  messagedlg( 'Настройки успешно сохранены.'#13#10'Хотите закрыть окно?' , mtConfirmation,  [mbYes,mbNo], 0)=6 //Выключить кнопку
                then  settingsform.Close;

             MainForm.StatusBar.Panels[0].Text:='Изменение настроек приложения завершено успешно!'   //Задаем статус на панели главной формы

             end
           else //Если проверка на корректность заполнения не пройдена
             begin
             messagedlg( 'Исправьте ошибки перед сохранением!' , mtError, [mbOk], 0, mbOk);
             MainForm.StatusBar.Panels[0].Text:='Возникла ошибка при изменении настроек.'
             {
               try
                 WriteConnStr(ConnStrTxt.Text);
                except messagedlg( 'Невозможно сохранить значение в реестр!' , mtError, [mbOk], 0, mbOk);
                end;
            if not WriteSettings then
                 messagedlg( 'Невозможно сохранить настройки. Непредвиденная ошибка!' , mtError, [mbOk], 0, mbOk);
                    }
         //  SettingsForm.Close;   //Закрыть окно
         //  MainForm.StatusBar.Panels[0].Text:='Настройки успешно сохранены.'  //Задаем статус на панели главной формы

            end;

  // SettingsForm.Close;
  end
  //Если нет подключения к БД
    else
      begin
        try
          WriteConnStr(ConnStrTxt.Text);
        except  messagedlg( 'Невозможно сохранить значение в реестр!' , mtError, [mbOk], 0, mbOk);
        end;

         messagedlg( 'Настройки успешно сохранены. Хотите закрыть окно?' , mtConfirmation,  [mbYes,mbNo], 0);
      end;
 end;




procedure TSettingsForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
with mainform do begin
if CheckDBConn then begin
                      statmemo.Clear;
                     AddColoredLine(statmemo,'Соединение с БД - УСПЕШНО!', clGreen)
                    end
                      else begin
                        statmemo.Clear;
                        infmemo.Clear;
                        infsstatmemo.Clear;
                        infsstatmemo.Lines.Add('Невозможно получить статистику!');
                        AddColoredLine(statmemo,'Нет связи с БД! ОШИБКА!', clRed);
                        infmemo.Lines.Add('Нет данных!');
                      end;
                 checksettings;
                  end;
end;


procedure TSettingsForm.FormShow(Sender: TObject);
begin
 //Если удалось получить значение строки подключения из реестра
  if MainForm.ReadConnStr<>'not_' then
   begin
    AdoConnSet.Close;
    ADOConnSet.ConnectionString:=MainForm.ReadConnStr;
      try
        ADOConnSet.Open;
       except messagedlg( 'Строка подключения к БД некорректна!'+#13#10+' Нет связи с сервером!' , mtError, [mbOk], 0, mbOk)
      end;
       //Заполнение строки подключения в memo
          ConnStrTxt.Text:= MainForm.ReadConnStr;
   end
   //Если не удалось получить значение строки подключения из реестра
    else
      begin
        AdoConnSet.Close;
        ConnStrTxt.Text:=ADOConnSet.ConnectionString;
      end;
//Открываем соединение с базой
  try
    adoqueryset.Open;
  except
  end;
   //Загружаем настройки
     LoadSettings;
   //Проверяем корректность
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



