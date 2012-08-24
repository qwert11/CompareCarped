unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, ComObj, ExtCtrls, ActnList, StdActns, Buttons,
  ImgList, CheckLst;

type
  TForm1 = class(TForm)
    strngrd1: TStringGrid;
    strngrd2: TStringGrid;
    pnl1: TPanel;
    spl1: TSplitter;
    strngrd3: TStringGrid;
    pnl2: TPanel;
    spl2: TSplitter;
    pnlImport: TPanel;
    spl3: TSplitter;
    pnl3: TPanel;
    lbl1: TLabel;
    lbl2: TLabel;
    edtFromBuh: TEdit;
    edtMyTable: TEdit;
    btnBuhTable: TSpeedButton;
    btnMyTable: TSpeedButton;
    actlst1: TActionList;
    flpnBuhTable: TFileOpen;
    flpnMyTable: TFileOpen;
    il1: TImageList;
    chklstCompare: TCheckListBox;
    procedure FormActivate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure flpnBuhTableAccept(Sender: TObject);
    procedure flpnMyTableAccept(Sender: TObject);
    procedure chklstCompareEnter(Sender: TObject);
  private
    { Private declarations }
    procedure RemoveIfEmptyBuh(strngrd1: TStringGrid);
//    procedure RemoveIfEmptyMy(strngrd2: TStringGrid);
    function ImtTblMy: Boolean;
    function ImtTblBuh: Boolean;
    function ImportTable(SG: TStringGrid; fileName: string;
        iPage: Integer = 1): Boolean;
    function ComPare(): Boolean;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  PathBuh,
  PathMy: string;

implementation

uses Unit2, Auto1U;

{$R *.dfm}

var
  _3GridsForCompare: T3GridsForCompare;

procedure AutoSizeGridColumn(Grid: TStringGrid; column, min, max: Integer);
  { Set for max and min some minimal/maximial Values}
  { Bei max and min kann eine Minimal- resp. Maximalbreite angegeben werden}
var
  i: Integer;
  temp: Integer;
  tempmax: Integer;
begin
  tempmax := 0;
  for i := 0 to (Grid.RowCount - 1) do
  begin
    temp := Grid.Canvas.TextWidth(Grid.cells[column, i]);
    if temp > tempmax then tempmax := temp;
    if tempmax > max then
    begin
      tempmax := max;
      break;
    end;
  end;
  if tempmax < min then tempmax := min;
  Grid.ColWidths[column] := tempmax + Grid.GridLineWidth + 3;
end;



function Xls_To_StringGrid(AGrid: TStringGrid; AXLSFile: string;
      iPage: Integer): Boolean;
const
  xlCellTypeLastCell = $0000000B;
var
  XLApp, Sheet: OLEVariant;
  RangeMatrix: Variant;
  x, y, k, r: Integer;
begin
  Result := False;
  if AXLSFile = '' then
    Exit;

  Screen.Cursor:=crAppStart;
  XLApp := CreateOleObject('Excel.Application');
  try
     XLApp.Visible := False;
     XLApp.Workbooks.Open(AXLSFile);
     Sheet := XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets[iPage];
     Sheet.Activate;
     Sheet.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;
     x := XLApp.ActiveCell.Row;
     y := XLApp.ActiveCell.Column;
     AGrid.RowCount := x;
     AGrid.ColCount := y;
     RangeMatrix := XLApp.Range['A1', XLApp.Cells.Item[X, Y]].Value;
     k := 1;
     repeat
       for r := 1 to y do
        AGrid.Cells[(r - 1), (k - 1)] := RangeMatrix[K, R];
       Inc(k, 1);
       AGrid.RowCount := k + 1;
     until k > x;
     RangeMatrix := Unassigned;
     Result := True;
  finally
     if not VarIsEmpty(XLApp) then
     begin
       XLApp.Quit;
       XLAPP := Unassigned;
       Sheet := Unassigned;
       Screen.Cursor:=crDefault;
     end;
  end;

end;


Type
  TFakeGrid=class(TCustomGrid);

procedure SelectUsefulField (sg: TStringGrid);
var
  iCol,
  iRow: Integer;
  b: Boolean;
begin
  iRow := 0;
  repeat
    b := False;
    for iCol := 0 to sg.ColCount - 1 do
      if sg.Cells[iCol, iRow] <> '' then
        begin
          b := True;
          Break;
        end;
    if not b then
      begin
        TFakeGrid(sg).DeleteRow(iRow);
        Dec(iRow);
      end;
    Inc(iRow);
  until iRow >= sg.RowCount;
end;

procedure TForm1.RemoveIfEmptyBuh(strngrd1: TStringGrid);
const
  scNameArt = 'Товар/Склад/Документ';
  scTarck = 'Дорожки ИБРАГИМ Килимові доріжки';
  scPicture = 'Картины';
  //scMetal = '';
  scElse = 'Коврики Коврики нарезные Овечьи шкуры';

  function RemoveBetweenGoods(const s: string): Boolean;
  var
    RowIDX: Integer;
  begin
    Result := False;
    RowIDX :=  0;
    repeat

      if (Pos(Trim(strngrd1.Cells[0, RowIDX]), s) <> 0) and
          (Trim(strngrd1.Cells[3, RowIDX]) = '') then
        repeat
          strngrd1.Row := RowIDX;
          strngrd1.Refresh;
          Inc(RowIDX);
          if (strngrd1.Cells[1, RowIDX] = '') and
              (strngrd1.Cells[3, RowIDX] <> '') then
            begin
              TFakeGrid(strngrd1).DeleteRow(RowIDX);
              Dec(RowIDX);
              Result := True;
            end;
        until (RowIDX >= strngrd1.RowCount - 1) or
          (strngrd1.Cells[3, RowIDX] = '')
      else
        Inc(RowIDX);
    until ((RowIDX >= strngrd1.RowCount - 1)) or (strngrd1.Cells[0, RowIDX] = '');
  end;
begin
  SelectUsefulField(strngrd1);
  RemoveBetweenGoods(scTitleCarped);
  RemoveBetweenGoods(scPicture);
  RemoveBetweenGoods(scTarck);
  RemoveBetweenGoods(scElse);
end;

procedure TForm1.FormActivate(Sender: TObject);
begin
  chklstCompare.Checked[4] := True;
  {$BOOLEVAL ON}
  if ImtTblMy and
      ImtTblBuh
  then
    ComPare;
  OnActivate := nil;
end;


//procedure TForm1.RemoveIfEmptyMy(strngrd2: TStringGrid);
//const
//  scArticul = 'Арт. Артикул.';
//  scTitle = 'Товар Название';
//  scCount = 'В налич. В нал. шт. В налич. В нал. В нал. (м)';
//  scLength = 'Ш Ширина';
//  scWidth = 'В нал. (м) Дл. Длина';
//var
//  iCol,
//  iRow,
//  iColIfTitle: Integer;
//  s: string;
//  b: Boolean;
//begin
//  SelectUsefulField(strngrd2);
//  strngrd2.Refresh;
//  iRow := 0;
//  s := scArticul + scTitle + scCount + scLength + scWidth;
//  b := False;
//  while (not b) and (iRow < strngrd2.RowCount - 1) do
//  begin
//    iCol := 0;
//    while iCol < strngrd2.ColCount do
//    begin
//      if (b or (Trim(strngrd2.Cells[iCol, iRow]) <> '')) and
//          (Pos(Trim(strngrd2.Cells[iCol, iRow]), s) = 0) then
//        begin
//          // проверка на заглавный рядок
//          if not b then
//            begin
//              iColIfTitle := iCol;
//              repeat
//                if (Pos(Trim(strngrd2.Cells[iColIfTitle, iRow]), s) <> 0) then
//                  b := True
//                else
//                  Inc(iColIfTitle);
//              until b or (iColIfTitle >= strngrd2.ColCount - 1);
//            end;
//
//          TFakeGrid(strngrd2).DeleteColumn(iCol);
//          Dec(iCol);
//        end;
//      Inc(iCol);
//    end;
//    Inc(iRow);
//  end;
//end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  _3GridsForCompare.Free;
end;

function TForm1.ImportTable(SG: TStringGrid; fileName: string;
    iPage: Integer): Boolean;
  procedure grFocus(gr: TStringGrid; Row: Integer);
  begin
    gr.TopRow := Row;
    gr.LeftCol := 0;
    gr.Row := Row;
  end;
var
  Columns_: Integer;
begin
  Result := False;
  if not Xls_To_StringGrid(SG, fileName, iPage) then
    Exit;

  // подгоняем размеры полей Buh
  for Columns_ := 0 to SG.ColCount - 1 do
    AutoSizeGridColumn(SG, Columns_, 20, 200);

  // удаляем пустые строки и фокусируемся на первой записи
  if SG.Name = 'strngrd1' then
    RemoveIfEmptyBuh(SG);
  if SG.Name = 'strngrd2' then
    RemoveIfEmptyBuh(SG);
  grFocus(SG, 0);
  Result := True;
end;

function TForm1.ComPare: Boolean;
var
  Columns_: Integer;
begin
  Result := False;
  if (flpnBuhTable.Dialog.FileName <> '') and
      (flpnMyTable.Dialog.FileName <> '')
  then
    _3GridsForCompare := T3GridsForCompare.Create(strngrd1, strngrd2, strngrd3)
  else
    Exit;

  // подгоняем размеры полей Public
  for Columns_ := 0 to strngrd3.ColCount - 1 do
    AutoSizeGridColumn(strngrd3, Columns_, 20, 200);
  Result := True;
end;

procedure TForm1.flpnBuhTableAccept(Sender: TObject);
begin
  edtFromBuh.Text := flpnBuhTable.Dialog.FileName;

  if ImtTblBuh then
    begin
      if flpnMyTable.Dialog.FileName <> '' then
        ComPare;
    end
  else
    begin
      flpnBuhTable.Dialog.FileName := '';
      edtFromBuh.Text := '';
    end;
end;

procedure TForm1.flpnMyTableAccept(Sender: TObject);
begin
  edtMyTable.Text := flpnMyTable.Dialog.FileName;

  if ImtTblMy then
    begin
      if flpnBuhTable.Dialog.FileName <> '' then
        ComPare;
    end
  else
    begin
      flpnMyTable.Dialog.FileName := '';
      edtMyTable.Text := '';
    end;
end;

function TForm1.ImtTblBuh: Boolean;
begin
  Result := False;
  if ImportTable(strngrd1, flpnBuhTable.Dialog.FileName) then
    begin
      //RemoveIfEmptyBuh(strngrd1);
      Result := True;
    end;
end;

function TForm1.ImtTblMy: Boolean;
begin
  Result := False;
  if ImportTable(strngrd2, flpnMyTable.Dialog.FileName) then
    begin
      //RemoveIfEmptyMy(strngrd2);
      Result := True;
    end;
end;

procedure TForm1.chklstCompareEnter(Sender: TObject);
begin
  //ShowMessage('');
end;

end.
