unit Unit_bing;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, OleCtrls, SHDocVw, StdCtrls, IdBaseComponent, IdComponent,
  IdTCPConnection, IdTCPClient, IdHTTP,RegExpr,Registry,UrlMon,ExtActns,ShlObj
  ,ComObj,WinInet;

type
  TForm1 = class(TForm)
    wb1: TWebBrowser;
    idhtp1: TIdHTTP;
    mmo1: TMemo;
    btnopenweb: TButton;
    btnsearch: TButton;
    mmo2: TMemo;
    edt1: TEdit;
    dlgOpen1: TOpenDialog;
    btn1: TButton;
    procedure FormCreate(Sender: TObject);
    procedure btnopenwebClick(Sender: TObject);
    procedure btnsearchClick(Sender: TObject);

    procedure btn1Click(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;



var
  Form1: TForm1;
  imgpath,sfile:string;
  Mycs:TRTLCriticalSection;

   procedure openanddownpic();
    function DownloadFile(SourceFile, DestFile: string): Boolean;
    procedure setbg();
implementation

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
var
  types:Integer;
begin

  InitializeCriticalSection(Mycs);
    types:=INTERNET_CONNECTION_MODEM + INTERNET_CONNECTION_LAN + INTERNET_CONNECTION_PROXY;

  if  InternetGetConnectedState(@types,0) then
  begin
     //WindowState:=wsMaximized;
     wb1.Navigate('http://cn.bing.com');
     sfile:='C:\Users\Administrator\AppData\Roaming\Microsoft\Windows\Themes\TranscodedWallpaper.bmp';
  end
  else
  begin
    ShowMessage('网络未连接');
    form1.Free;
  end;

     EnterCriticalSection(Mycs);

   wb1.Navigate('http://cn.bing.com');
   sfile:='C:\Users\Administrator\AppData\Roaming\Microsoft\Windows\Themes\TranscodedWallpaper.bmp';

 openanddownpic;
  Sleep(1000);
  setbg;
   {
   btnopenwebClick(Sender);
   Sleep(1000);
   btn1Click(Sender);
   }
end;

procedure TForm1.btnopenwebClick(Sender: TObject);
var
  s:TStringList;
  regexpr:TRegExpr;
  i:Integer;
begin
    s:=TStringList.Create;
    regexpr:=TRegExpr.Create;
    regexpr.Expression:= Trim(edt1.Text);//'/g_img={url:'(.*)'';
    try
    s.Text:=idhtp1.Get('http://cn.bing.com');
    mmo1.Text:=s.Text;
    mmo2.Lines.Clear;
    regexpr.InputString:=mmo1.Text;
     if regexpr.Exec then
      begin
         for i:=1 to regexpr.SubExprMatchCount do
         begin
            mmo2.Lines.Add(Format('Match[%d]=',[I])+regexpr.Match[i]);
         end;
         wb1.Navigate(regexpr.Match[1]);
         imgpath:=regexpr.Match[1];
         mmo2.Lines.Add('匹配成功');
      end
      else
      begin
       // http://www\.mynet\.com/register\.asp\?id=(\d+)&name=(\w+)
       // g_img={url:'(\*)
        mmo2.Lines.Add('匹配失败');
      end;
    finally
      s.Free;
      FreeAndNil(regexpr);
    end;

     while  FileExists(sfile) do
    begin
      SetFileAttributes(PChar(sfile),FILE_ATTRIBUTE_ARCHIVE);
      if DeleteFile(sfile) then
      // ShowMessage('删除源文件成功');
    end;

    DownloadFile(imgpath,sfile);
end;

procedure TForm1.btnsearchClick(Sender: TObject);
var
  reg:TRegistry;
  stemp:string;
begin
  stemp:='C:\Users\Administrator\AppData\Roaming\Microsoft\Windows\Themes\001.bmp';

    // dlgOpen1.FileName:=sfile;
    reg:=TRegistry.Create;
    reg.RootKey:= HKEY_CURRENT_USER;
    reg.OpenKey('Control Panel\Desktop',False);
    reg.WriteString('TileWallPaper','0');
   // reg.WriteString('WallPaper',sfile);
     reg.WriteString('WallPaper',stemp);
    reg.CloseKey;
    reg.Free;
   Systemparametersinfo(SPI_SETDESKWallpaper,0,Nil,SPIF_SendChange);
   // Systemparametersinfo(SPI_SETDESKWallpaper,0,PChar(sfile),SPIF_SendChange);
    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NiL, NiL);

end;

function DownloadFile(SourceFile, DestFile: string): Boolean;
var
  haserror:Boolean;
begin
     haserror:=False;
     with  TDownLoadURL.Create(form1) do
     begin
       try
       URL:=SourceFile;
       Filename:=DestFile;
       //isdowning:=True;
       //OnDownloadProgress:=url_
       ExecuteTarget(nil);
       except on e:Exception do
       begin
         ShowMessage(e.Message);
         Free;
         haserror:=True;
       end;
      end;
      Result:=not haserror;
     end;
end;
procedure TForm1.btn1Click(Sender: TObject);
var
  hObj: IUnknown;
  ADesktop: IActiveDesktop;
stemp:string;
begin
 // stemp:='C:\Users\Administrator\AppData\Roaming\Microsoft\Windows\Themes\TranscodedWallpaper.bmp';
 begin
  hObj := CreateComObject(CLSID_ActiveDesktop);
  ADesktop := hObj as IActiveDesktop;
  ADesktop.SetWallpaper(PWideChar(WideString(sfile)), 0);
  ADesktop.ApplyChanges(AD_APPLY_ALL or AD_APPLY_FORCE);
 end;
end;

procedure openanddownpic;
var
  s:TStringList;
  regexpr:TRegExpr;
  i:Integer;
begin
    s:=TStringList.Create;
    regexpr:=TRegExpr.Create;
    regexpr.Expression:= Trim(Form1.edt1.Text);//'/g_img={url:'(.*)'';
    try
    s.Text:=Form1.idhtp1.Get('http://cn.bing.com');
    Form1.mmo1.Text:=s.Text;
    Form1.mmo2.Lines.Clear;
    regexpr.InputString:=Form1.mmo1.Text;
     if regexpr.Exec then
      begin
         for i:=1 to regexpr.SubExprMatchCount do
         begin
            Form1.mmo2.Lines.Add(Format('Match[%d]=',[I])+regexpr.Match[i]);
         end;
         Form1.wb1.Navigate(regexpr.Match[1]);
         imgpath:=regexpr.Match[1];
         Form1.mmo2.Lines.Add('匹配成功');
      end
     else
      begin
       // http://www\.mynet\.com/register\.asp\?id=(\d+)&name=(\w+)
       // g_img={url:'(\*)
        Form1.mmo2.Lines.Add('匹配失败');
        Form1.Close;
      end;
    finally
      s.Free;
      FreeAndNil(regexpr);
    end;


     while  FileExists(sfile) do
    begin
      SetFileAttributes(PChar(sfile),FILE_ATTRIBUTE_ARCHIVE);
      if DeleteFile(sfile) then
      //ShowMessage('删除源文件成功');
    end;
    DownloadFile(imgpath,sfile);
end;

procedure setbg;
var
  hObj: IUnknown;
  ADesktop: IActiveDesktop;
stemp:string;
begin
 // stemp:='C:\Users\Administrator\AppData\Roaming\Microsoft\Windows\Themes\TranscodedWallpaper.bmp';
 begin

  hObj := CreateComObject(CLSID_ActiveDesktop);
  ADesktop := hObj as IActiveDesktop;
  ADesktop.SetWallpaper(PWideChar(WideString(sfile)), 0);
  ADesktop.ApplyChanges(AD_APPLY_ALL or AD_APPLY_FORCE);
 end;
  DeleteCriticalSection(Mycs);
  form1.Free;
end;
end.
