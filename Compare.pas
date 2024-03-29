unit Compare;

interface

uses
  Unit2, SysUtils, Classes;

function Compare2Str(str1, str2: string): Integer;
function Compare2InfoRecord(InfoRecord1, InfoRecord2: TInfoRecord): Integer;



implementation

procedure DeleteGarbage(var str: string);
var
  i: Integer;
begin
  if str = '' then
    Exit;

  // ������� ��������
  i := 1;
  repeat
    if not (str[i] in ['A'..'z', '�'..'�']) then
      Delete(str, i, 1);
    Inc(i);
  until i > Length(str);
end;

function Compare2Str(str1, str2: string): Integer;
var
  i, l1, l2: Integer;
begin
  Result := 0;

  Trim(str1);
  Trim(str2);

  if (str1 = '') then
    Abort;
  if (str2 = '') then
    Abort;

  DeleteGarbage(str1);
  DeleteGarbage(str2);

  str1 := LowerCase(str1);
  str2 := LowerCase(str2);

  if (str1 = '') or (str2 = '') then
    Error(reAccessViolation);

  i := Pos(str1, str2);
  if i = 0 then
    i := Pos(str2, str1);

  if i <> 0 then
    begin
      l1 := Length(str1);
      l2 := Length(str2);
      if l1 > l2 then
        Result := 100 - Abs((Round((l2 - l1) * 100 / l1)))
      else
        if l1 < l2 then
          Result := 100 - Abs(Round((l1 - l2) * 100 / l2))
        else
           Result := 100;
    end;
end;

function Compare2InfoRecord(InfoRecord1, InfoRecord2: TInfoRecord): Integer;
label
  GoToLabel;
var
  i, ii: Integer;
begin
  Result := 0;
  if not(((InfoRecord1.itsWidth = InfoRecord2.itsWidth) and
      (InfoRecord1.itsLength = InfoRecord2.itsLength)) or
        ((InfoRecord1.itsLength = InfoRecord2.itsWidth) and
          (InfoRecord1.itsWidth = InfoRecord2.itsLength))) then
    Exit;

  DeleteGarbage(InfoRecord1.itsTitle);
  DeleteGarbage(InfoRecord2.itsTitle);
//  for i := 0 to InfoRecord1.itsListArt.Count - 1 do
//  begin
//    for ii := 0 to InfoRecord2.itsListArt.Count - 1 do
//      if CompareText(InfoRecord1.itsListArt[i],
//          InfoRecord2.itsListArt[ii]) = 0 then
//        begin
//          Result := 50;
//          goto GoToLabel;
//        end;
//  end;
  for i := 0 to InfoRecord1.itsListArt.Count - 1 do
    if InfoRecord2.itsListArt.Find(InfoRecord1.itsListArt[i], ii) then
      begin
        Result := 50;
        Break;
      end;

//  GoToLabel:

  // �������� 150 % �.�. 50% ��� � 100 % �� Compare2Str
  Result := Result + Compare2Str(InfoRecord1.itsTitle, InfoRecord2.itsTitle);
end;

end.
