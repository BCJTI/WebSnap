program WebSnap;

{$APPTYPE CONSOLE}

uses
  Windows, SysUtils, Classes, Graphics, SHDocVw, ExtCtrls, ActiveX, JPEG,
  WinInet, Dialogs, PngImage;

type
  TEventHandlers = class
    procedure OnTimerTick(Sender : TObject);
    procedure wbIEDocumentComplete(ASender: TObject;
      const pDisp: IDispatch; var URL: OleVariant);
  end;

var
  i: Integer;
  sParam, sVal: string;
  dtInit, dtEnd: TDateTime;

  url, outFile, outExt: string;
  width, height, wait, quality: Integer;
  crop: array of Integer;
  EventHandlers : TEventHandlers;
  wbIE: TWebBrowser;
  tmr: TTimer;

const
  slashN = Chr(13) + Chr(10);
  sHelp = 'WebSnap --help' + slashN +
          ' ' + slashN +
          'WebSnap is inspired by CutyCapt (cutycapt.sourceforge.net) ' + slashN +
          '-----------------------------------------------------------------------------' + slashN +
          'Usage: WebSnap --url=http://www.google.com/ --out=outfile.jpg' + slashN +
          '-----------------------------------------------------------------------------' + slashN +
          ' --help                    Print help and exit' + slashN +
          ' --url=<url>               The URL to capture (http/https/file)' + slashN +
          ' --out=<path>.<f>          The target file (jpg|bmp|jpg ... <f>)' + slashN +
          ' --quality=<int>           0-100 JPG quality if out is a JPG (default: 70)' + slashN +
          ' --width=<int>             Browser width for Snapshot (default: 1366)' + slashN +
          ' --height=<int>            Browser height for Snapshot (default: 667)' + slashN +
          ' --wait=<ms>               After load, wait (default: 2000)' + slashN +
          ' --crop=<int,int,int,int>  Crop image (default: 0, 0, --width, --height)' + slashN +
          '-----------------------------------------------------------------------------' + slashN +
          ' <f> is jpg,png,bmp' + slashN +
          '-----------------------------------------------------------------------------';

function Token(Texto: string; Separador: Char; Bloco: Integer): string;
var
  Contador: Integer;
  p1, p2: Integer;
  sResult: string;
begin
  Result := '';
  sResult := '';
  if Bloco > 0 then begin
    Contador := 1;
    p1 := 1;
    if Bloco > 1 then begin
      while (Contador < Bloco) and (p1 < Length(Texto)) do begin
        if Texto[p1] = Separador then
          inc(Contador);
        inc(p1);
      end;
    end;
    if Contador < Bloco then begin
      Exit;
    end;
    p2 := p1;
    while (p2 <= Length(Texto)) and (Texto[p2] <> Separador) do
      inc(p2);
    sResult := Copy(Texto, p1, p2 - p1);
  end;
  Result := sResult;
end;

function CountToken(Value: string; Block: Char): Integer;
var
  I: Integer;
begin
  Result := 0;
  if Trim(Value) <> '' then begin
    if Value[Length(Value)] <> Block then
      Value := Value + Block;
    for I := 1 to Length(Value) do
      if Value[I] = Block then
        Result := Result + 1;
  end;
end;

procedure CropBitmap(InBitmap: TBitmap; X, Y, W, H :Integer);
begin
  BitBlt(InBitmap.Canvas.Handle, 0, 0, W, H, InBitmap.Canvas.Handle, X, Y, SRCCOPY);
  InBitmap.Width := W;
  InBitmap.Height:= H;
end;

procedure SaveImage(sArq: String);
var
  ViewObject: IViewObject;
  sourceDrawRect: TRect;
  imgBMP: TImage;
  imgJPG: TJPegImage;
  imgPNG: TPngImage;
begin
  if wbIE.Document <> nil then
  try
    imgBMP := TImage.Create(nil);
    imgBMP.Width := wbIE.Width;
    imgBMP.Height := wbIE.Height;

    wbIE.Document.QueryInterface(IViewObject, ViewObject);
    if ViewObject <> nil then begin
      try
        sourceDrawRect := Rect(0, 0, imgBMP.Width, imgBMP.Height);
        ViewObject.Draw(DVASPECT_CONTENT, 1, nil, nil, wbIE.Handle,
        imgBMP.Canvas.Handle, @sourceDrawRect, nil, nil, 0);
      finally
        ViewObject._Release;
      end;

      CropBitmap(imgBMP.Picture.Bitmap, crop[0], crop[1], crop[2], crop[3]);

      if(outExt = '.bmp') then begin
        imgBMP.Picture.SaveToFile(sArq);
      end
      else if(outExt = '.jpg') then begin
        imgJPG := TJPegImage.Create;
        imgJPG.CompressionQuality := quality;
        imgJPG.Assign(imgBMP.Picture.Bitmap);
        imgJPG.SaveToFile(sArq);
      end
      else if(outExt = '.png') then begin
        imgPNG := TPngImage.Create;
        imgPNG.Assign(imgBMP.Picture.Bitmap);
        imgPNG.SaveToFile(sArq);
      end;

      dtEnd := Now;
      Writeln('Outfile Saved in ' + FormatDateTime('ss:zzz', dtEnd - dtInit));
      Halt(0);
    end;
  except
    on E: Exception do begin
      Writeln('An error occured:');
      Writeln(E.Message);
      Halt(1);
    end;
  end;
end;

procedure TEventHandlers.OnTimerTick(Sender : TObject);
begin
  tmr.Enabled := False;
  SaveImage(outFile);
end;

procedure TEventHandlers.wbIEDocumentComplete(ASender: TObject;
  const pDisp: IDispatch; var URL: OleVariant);
begin
  tmr.Enabled := true;
end;


procedure DeleteIECache;
var
  lpEntryInfo: PInternetCacheEntryInfo;
  hCacheDir: LongWord;
  dwEntrySize: LongWord;
begin { DeleteIECache }
  dwEntrySize := 0;

  FindFirstUrlCacheEntry(nil, TInternetCacheEntryInfo(nil^), dwEntrySize);

  GetMem(lpEntryInfo, dwEntrySize);

  if dwEntrySize>0 then
    lpEntryInfo^.dwStructSize := dwEntrySize;

  hCacheDir := FindFirstUrlCacheEntry(nil, lpEntryInfo^, dwEntrySize);

  if hCacheDir<>0 then
  begin
    repeat
      DeleteUrlCacheEntry(lpEntryInfo^.lpszSourceUrlName);
      FreeMem(lpEntryInfo, dwEntrySize);
      dwEntrySize := 0;
      FindNextUrlCacheEntry(hCacheDir, TInternetCacheEntryInfo(nil^), dwEntrySize);
      GetMem(lpEntryInfo, dwEntrySize);
      if dwEntrySize>0 then
        lpEntryInfo^.dwStructSize := dwEntrySize;
    until not FindNextUrlCacheEntry(hCacheDir, lpEntryInfo^, dwEntrySize)
  end; { hCacheDir<>0 }
  FreeMem(lpEntryInfo, dwEntrySize);

  FindCloseUrlCache(hCacheDir)
end; { DeleteIECache }

procedure MsgPump;
var
  Unicode: Boolean;
  Msg: TMsg;
begin
  while GetMessage(Msg, 0, 0, 0) do begin
    Unicode := (Msg.hwnd = 0) or IsWindowUnicode(Msg.hwnd);
    TranslateMessage(Msg);
    if Unicode then
      DispatchMessageW(Msg)
    else
      DispatchMessageA(Msg);
  end;
end;

begin
  dtInit := Now;

  url     := '';
  outFile := '';
  quality := 70;
  width   := 1366;
  height  := 667;
  wait    := 2000;
  SetLength(crop, 4);
  crop[0] := 0; crop[1] := 0; crop[2] := width; crop[3] := height;

  try
    for i := 1 to ParamCount do begin
      sParam := Copy(ParamStr(i), 1, Pos('=', ParamStr(i)));
      sVal   := Copy(ParamStr(i), Pos('=', ParamStr(i)) + 1,  Length(ParamStr(i)));

      if sParam = '--url=' then
        url := sVal
      else if sParam = '--out=' then begin
        outFile := sVal;
        outExt  := LowerCase(Copy(sVal, Length(sVal) - 3, 4));
        if( not ((outExt = '.jpg') or (outExt = '.png') or (outExt = '.bmp')) ) then begin
          Raise Exception.Create('Invalid output format "' + outExt + '"');
        end;
      end
      else if sParam = '--width=' then
        width := StrToInt(sVal)
      else if sParam = '--quality=' then
        quality := StrToInt(sVal)
      else if sParam = '--height=' then
        height := StrToInt(sVal)
      else if sParam = '--wait=' then
        wait := StrToInt(sVal)
      else if sParam = '--crop=' then begin
        if CountToken(sVal, ',') <> 4 then begin
          Raise Exception.Create('Invalid value for --crop');
        end;
        crop[0] := StrToInt(Token(sVal, ',', 1));
        crop[1] := StrToInt(Token(sVal, ',', 2));
        crop[2] := StrToInt(Token(sVal, ',', 3));
        crop[3] := StrToInt(Token(sVal, ',', 4));
      end;
    end;
  except
    on E: Exception do begin
      Writeln('Invalid Param. Check Params list:' + slashN);
      Writeln(sHelp);
      Writeln('Error:');
      Writeln(E.Message);
      halt(1);
    end;
  end;

  if (ParamCount = 1) or (ParamStr(1) = '--help') or (url = '') or (outFile = '') then begin
    WriteLn(Output, sHelp);
    Exit;
  end;

  //Here de game begins
  try

    EventHandlers := TEventHandlers.Create();

    tmr := TTimer.Create(nil);
    tmr.Enabled  := False;
    tmr.Interval := wait;
    tmr.OnTimer  := EventHandlers.OnTimerTick;

    DeleteIECache;
    wbIE := TWebBrowser.Create(nil);
    wbIE.OnDocumentComplete := EventHandlers.wbIEDocumentComplete;
    wbIE.HandleNeeded;
    wbIE.Silent := True;
    wbIE.Width  := width;
    wbIE.Height := height;
    wbIE.Navigate(url);

    MsgPump;
  except
    on E: Exception do begin
      Writeln(E.ClassName, ': ', E.Message);
      Halt(1);
    end;
  end;


end.
