unit AboutUnit;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.StdCtrls, Vcl.Buttons,
  Vcl.Imaging.pngimage, ShellAPI;

type
  TAboutForm = class(TForm)
    AboutImage: TImage;
    OKButt: TBitBtn;
    urlLink: TLabel;
    AboutMemoFull: TMemo;
    Panel1: TPanel;
    Panel2: TPanel;
    AboutMemo: TMemo;
    procedure OKButtClick(Sender: TObject);
    procedure urlLinkClick(Sender: TObject);
    procedure AboutMemoClick(Sender: TObject);
    procedure AboutMemoMouseEnter(Sender: TObject);
    procedure AboutMemoMouseLeave(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  AboutForm: TAboutForm;

implementation

{$R *.dfm}

procedure TAboutForm.AboutMemoClick(Sender: TObject);
begin
  ShellExecute(handle, 'open', 'https://www.gnu.org/licenses/gpl-3.0.txt', nil, nil, SW_SHOW);
end;

procedure TAboutForm.AboutMemoMouseEnter(Sender: TObject);
begin
 urlLink.Font.Color:=clBlue;
end;

procedure TAboutForm.AboutMemoMouseLeave(Sender: TObject);
begin
   urlLink.Font.Color:=clNavy;
end;

procedure TAboutForm.OKButtClick(Sender: TObject);
begin
AboutForm.Close;
end;

procedure TAboutForm.urlLinkClick(Sender: TObject);
begin
  ShellExecute(handle, 'open', 'https://github.com/Gidrotr0nik/Informing_', nil, nil, SW_SHOW);

end;

end.
