unit autocomplete;

{$MODE Delphi}
(*
   The "autocomplete" unit can be used for classes based on TEdit,
   such as TMemo, TRichEdit, TEdit, TLabeledEdit etc.
   The initialization takes place, for example, as follows:
        AutoComplete1: = TAutoComplete1.Create (Form1, 'TextFile', Memo1);
   This adds AutoComplete to the 'Memo1' component.
   The text file contains the dictionary which is important for Memo1.

   Only words that are longer than 4 characters should be entered in the dictionary.
   AutoComplete only reacts if you enter 3 characters or more.

   if a pop-up menu is defined for the component, then this is displayed around the
   Menu item "Copy selected text to dictionary." expanded

   Important: With AutoComlete the event handlers 'OnChange', 'OnKeyDown'
            and 'OnKeyPress' overwritten.
*)

interface

uses classes, Controls, Graphics, Forms, Dialogs,
  StdCtrls, ComCtrls, Menus, LCLIntf, LCLtype, LMessages, SysUtils, Windows, addresslist, MemoryRecordUnit;

  type
  TAutoComplete = class(TListView)
    Separator: TMenuItem;
    ApplyWord: TMenuItem;

    procedure AutoCompleteClick(Sender: TObject);
    procedure AutoCompleteKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure AutoCompleteKeyPress(Sender: TObject; var Key: Char);
    procedure EditKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure EditKeyPress(Sender: TObject; var Key: Char);
    procedure EditOnChange(Sender: TObject);
    procedure ApplyWordClick(Sender: TObject);
    procedure StringListOnChange(Sender: TObject);
    procedure XEditContextPopup(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
  private
    { private Declarations }
    Comp: TControl;
    mX, mY: Integer;
    SelStart, SelLength: Integer;
    strList: TStringList;
    wordCount: Integer;
    AutoCompleteWasVisible: Boolean;
    XEdit: TEdit;
    WordList: String;
    strListHasChanged: Boolean;
    HintHidePause: Integer;
  public
    { public Declarations }
    constructor Create(AOwner: TComponent; var al: TAddresslist; const Edit: TObject); Overload;
    destructor Destroy; Override;
  end;

  const
  Delimiters = [' ', #$A, #$D, ';', '.', ':', ',', '"', #9, #39, '(', ')', '[', ']', '{', '}', '!', '?', '=', '/', '\', '-'];
  MinLen = 5;  // Min word length to be included in the search

  implementation

constructor TAutoComplete.Create(AOwner: TComponent; var al: TAddresslist; const Edit: TObject);
var
  i: integer;
begin
  inherited Create(AOwner);
  Parent := TWinControl(AOwner);             // Set the parent to avoid run-time error 65
  Height := 65;
  Width := 192;  // 192
  Visible := False;
  ViewStyle := vsReport;
  Sorttype := stText;
  ShowColumnHeaders := False;
  HideSelection := False;
  Color := clWindow;
  XEdit := TEdit(Edit);
  XEdit.OnKeyDown := EditKeyDown;
  XEdit.OnKeyPress := EditKeyPress;
  XEdit.OnChange := EditOnChange;
  XEdit.OnContextPopup := XEditContextPopup;
  mX := 0;
  mY := 0;
  Comp := TControl(XEdit);
  while Comp.HasParent Do
  begin
    mX := mX + Comp.Left;
    mY := mY + Comp.Top;
    Comp := Comp.Parent;
  end;
  if Assigned(XEdit.PopupMenu) then
  begin
    Separator := TMenuItem.Create(self);
    Separator.Caption := '-';
    ApplyWord := TMenuItem.Create(self);
    ApplyWord.Caption := '&Transfer selected word to the dictionary.';
    ApplyWord.OnClick := ApplyWordClick;
    XEdit.PopupMenu.Items.Add(Separator);
    XEdit.PopupMenu.Items.Add(ApplyWord);
  end;
  strList := TStringList.Create;
  strList.Sorted := True;
  strList.Duplicates := dupIgnore;
  strList.OnChange := StringListOnchange;

  for i := 0 to al.count -1 do
  begin
    WordList := WordList + al.MemRecItems[i].Description;
  end;

  strListHasChanged := False;
  wordCount := strList.Count;
  Columns.Add;
  Columns[0].Width := Width - 20;
  OnClick := AutoCompleteClick;
  OnKeyDown := AutoCompleteKeyDown;
  OnKeyPress := AutoCompleteKeyPress;
  Hint := 'Use ↑↓ keys to navigate' + #$A + #$D + 'Select with Enter/Left Mouse Click' + #$A + #$D + 'Cancel with ESC.';
  ShowHint := True;
  HintHidePause := Application.HintHidePause;
end;


destructor TAutoComplete.Destroy;
begin
  if strListHasChanged then
   try
    strList.SaveToFile(WordList);
   finally
    strList.Free;
    inherited Destroy;
   end;
end;

procedure TAutoComplete.AutoCompleteClick(Sender: TObject);
var
  str: String;
begin
  XEdit.SelStart := SelStart;
  XEdit.SelLength := SelLength;
  if ItemFocused = nil then
    str := TopItem.Caption
  else
    str := ItemFocused.Caption;
  XEdit.SelText := str + ' ';
  XEdit.SetFocus;
  AutoCompleteWasVisible := True;
  Visible := False;
  Application.HintHidePause := HintHidePause;
end;

procedure TAutoComplete.AutoCompleteKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  XEdit.SetFocus;
end;

procedure TAutoComplete.AutoCompleteKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = Char(VK_RETURN) then
    AutoCompleteClick(Sender)
  else
  if Key = Char(VK_ESCAPE) then
  begin
    Visible := False;
    Application.HintHidePause := HintHidePause;
  end
  else
  begin
    XEdit.SetFocus;
    keybd_event(Byte(Key), 0, KEYEVENTF_EXTendEDKEY, 0);
  end;
end;

procedure TAutoComplete.EditKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Visible and (not Focused) then
  begin
    if Key In [VK_RETURN, VK_TAB, VK_ESCAPE, VK_LEFT, VK_UP, VK_RIGHT, VK_DOWN] then
    begin
      SetFocus;
      if Key = VK_RETURN then
      begin
        Key := 0;
        AutoCompleteClick(Sender);
      end
      else
      begin
        keybd_event(Byte(Key), 0, 0, 0);
        Key := 0;
      end;
    end;
  end;
end;

procedure TAutoComplete.EditKeyPress(Sender: TObject; var Key: Char);
begin
  if (Key = #13) and AutoCompleteWasVisible then
  begin
    Key := #0;
    AutoCompleteWasVisible := False;
  end;
end;


procedure TAutoComplete.EditOnChange(Sender: TObject);
var
  i, k, l: Integer;
  s1, s2: String;
  pt: TPoint;
begin
  i := -1;
  // Get the length of the word
  repeat
    i := i + 1;
  until ((XEdit.SelStart - i <= 0) or
      (XEdit.Text[XEdit.SelStart - i] in Delimiters));
  Visible := False;
  Application.HintHidePause := HintHidePause;
  if i > 2 then
  begin
    //Wort ermitteln
    s1 := Copy(XEdit.Text, XEdit.SelStart - i + 1, i);
    Clear;
    SelStart := XEdit.SelStart - i;
    XEdit.SelStart := SelStart + 1;
    GetCaretPos(pt);
    XEdit.SelStart := SelStart + i;
    SelLength := i;
    For k := 0 To wordCount - 1 Do
    begin
      s2 := '';
      l := AnsiStrLIComp(PAnsiChar(s1), PAnsiChar(strList.Strings[k]), i);
      if l = 0 then
        Items.Add.Caption := strList.Strings[k]
      else
      if l < 0 then // if the queried word is smaller than the following
        break;      // then break
    end;
    if Items.Count > 0 then
    begin
      Items[0].Selected := True;
      ItemFocused := Items[0];
      Left := pt.X - 10 + mX;
      Top := pt.Y - XEdit.Font.Height * 6 Div 4 + mY;
      Application.HintHidePause := 10000;
      Visible := True;
    end;
  end;
end;

procedure TAutoComplete.ApplyWordClick(Sender: TObject);
var
  str: String;
begin
  if XEdit.SelLength >= MinLen then
  begin
    str := Trim(XEdit.SelText);
    strList.Add(str);
    wordCount := strList.Count;
  end;
end;

procedure TAutoComplete.XEditContextPopup(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
begin
  if XEdit.SelLength >= MinLen then
    ApplyWord.Enabled := True
  else
    ApplyWord.Enabled := False;
end;


procedure TAutoComplete.StringListOnChange(Sender: TObject);
begin
  strListHasChanged := True;
end;

end.
