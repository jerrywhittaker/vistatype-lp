// ---------------------------------------------------------------------------------------
// NOT IN USE. Nothing includes this file and nothing calls it.
//
// Jerry's decision, 8/13/2026: the installer went back to running Merge-Qat.ps1 and
// Remove-Qat.ps1, because taking PowerShell out did NOT stop Windows Defender deleting the
// Setup.exe - same detection, Trojan:Win32/Bearfoos.B!ml - and the PowerShell version is the
// one with years of field history behind it.
//
// It is kept because it WORKS and was proven equal to the scripts on every path (see the
// note further down), so reviving it is a one-line #include in vistatype.iss plus the two
// calls in CurStepChanged and CurUninstallStepChanged. The reason to revive it would not be
// Defender: it is that -ExecutionPolicy Bypass is exactly what institutional IT blocks, and
// on such a machine the [Run] entries fail SILENTLY, leaving the transcriber with no toolbar
// and no ribbon tabs and nothing said about why.
// ---------------------------------------------------------------------------------------

// ---------------------------------------------------------------------------------------
// qat-code.iss  --  the Quick Access Toolbar and ribbon-tab setup, in Inno Setup's own
//                   scripting language. Included by vistatype.iss inside its [Code] section.
//
// This is a port of installer/scripts/Merge-Qat.ps1 and Remove-Qat.ps1, which are no longer
// shipped or run. They stay in the repo as the reference this was written from, and their
// comments carry the reasoning behind every rule below - read those first if any of this
// needs changing.
//
// WHY IT WAS PORTED (Jerry, 8/13/2026). Windows Defender began deleting the installer, as
// Trojan:Win32/Bearfoos.B!ml - the "!ml" meaning a machine-learning guess rather than a match
// against anything known. An unsigned installer that launches powershell.exe with the
// execution policy bypassed, hidden, and drops a macro-enabled Word template into the STARTUP
// folder is the shape of Office macro malware, and a model cannot tell it from the real thing.
// Taking PowerShell out removes the loudest of those signals. It does NOT make the installer
// signed, and signing is the only reliable fix.
//
// The XML is done with MSXML2.DOMDocument.6.0 over COM, present on every version of Windows
// this add-in supports. Two things were measured on the build box before a line of this was
// written: MSXML is reachable from Inno's script, and its .save writes UTF-8 with NO
// byte-order mark, which is what Word requires of Word.officeUI.
//
// A NODE HANDED BACK BY VtCopyTree IS DEAD ONCE IT HAS BEEN APPENDED. Pascal Script does not
// keep the COM reference alive across a function return the way the tree does, so
//     El := VtCopyTree(...); Parent.appendChild(El); ... El.getAttribute('idQ')
// takes the installer down with a heap corruption (exit code 0xC0000374) at that last step.
// It cost four builds to find, because the log stopped one line short of the truth every
// time. Append the call's result directly - Parent.appendChild(VtCopyTree(...)) - and work
// out anything you need about the node from the SOURCE element beforehand.
//
// PROVEN EQUAL TO THE POWERSHELL IT REPLACED, 8/13/2026. Same starting toolbar, run twice -
// once through Merge-Qat.ps1 and once through this - and the resulting Word.officeUI compared
// structurally, in BOTH the Roaming and the Local copy:
//     "leave my toolbar alone"          identical, manifest identical
//     "install VistaType's toolbar"     identical, manifest identical
//     uninstall                         identical apart from an empty <mso:tabs/> being
//                                       written self-closed here and as an open/close pair by
//                                       .NET's XmlWriter. Same XML, same to Word.
// Redo that comparison if this file is ever changed; the scripts are still in the repo for it.
//
// NOTHING HERE HOLDS A NODE IN AN ARRAY. The first version collected child elements into an
// "array of Variant" and worked from that; it took the installer down with a heap corruption
// (exit code 0xC0000374) on the first real install. Pascal Script and COM references in a
// dynamic array do not mix. Every loop below walks the LIVE node list instead, and any loop
// that removes something walks it BACKWARDS so the indexes it has yet to visit cannot move.
// ---------------------------------------------------------------------------------------

const
  VT_MSO_NS   = 'http://schemas.microsoft.com/office/2009/07/customui';
  VT_MSOX_NS  = 'http://schemas.microsoft.com/office/2006/01/customui/special';
  VT_XMLNS_NS = 'http://www.w3.org/2000/xmlns/';
  VT_KEEP     = 'vt_tmp_keep';

var
  VtLogFile:    String;
  VtDotmPath:   String;
  VtPmUris:     TStringList;   // namespace URI ->
  VtPmPrefixes: TStringList;   //   prefix to write it under in THIS document
  VtMode:       String;        // Vista | Mine | Restore | None
  VtTabs:       String;        // Install | Skip | Remove

// ---------------------------------------------------------------------------------------
// small helpers
// ---------------------------------------------------------------------------------------

function VtStr(V: Variant): String;
begin
  Result := '';
  if VarIsNull(V) or VarIsEmpty(V) then Exit;
  Result := V;
end;

procedure VtLog(const S: String);
var
  Dir: String;
begin
  try
    Dir := ExpandConstant('{userappdata}\VistaType LP');
    if not DirExists(Dir) then ForceDirectories(Dir);
    if VtLogFile = '' then VtLogFile := Dir + '\qat.log';
    SaveStringToFile(VtLogFile, GetDateTimeString('yyyy-mm-dd hh:nn:ss', '-', ':') + '  ' + S + #13#10, True);
  except
    // a log that cannot be written must never stop an install
  end;
end;

function VtIdQPrefix(const IdQ: String): String;
var
  P: Integer;
begin
  P := Pos(':', IdQ);
  if P = 0 then Result := '' else Result := Copy(IdQ, 1, P - 1);
end;

function VtIdQLocal(const IdQ: String): String;
var
  P: Integer;
begin
  P := Pos(':', IdQ);
  if P = 0 then Result := IdQ else Result := Copy(IdQ, P + 1, Length(IdQ));
end;

function VtEndsWith(const S, Suffix: String): Boolean;
begin
  Result := (Length(S) >= Length(Suffix)) and
            (Copy(S, Length(S) - Length(Suffix) + 1, Length(Suffix)) = Suffix);
end;

function VtStartsWith(const S, Prefix: String): Boolean;
begin
  Result := Copy(S, 1, Length(Prefix)) = Prefix;
end;

function VtIsElement(N: Variant): Boolean;
begin
  Result := False;
  if VarIsNull(N) or VarIsEmpty(N) then Exit;
  Result := N.nodeType = 1;
end;

function VtCountElements(Node: Variant): Integer;
var
  I: Integer;
begin
  Result := 0;
  for I := 0 to Node.childNodes.length - 1 do
    if VtIsElement(Node.childNodes.item[I]) then Result := Result + 1;
end;

// ---------------------------------------------------------------------------------------
// namespaces
// ---------------------------------------------------------------------------------------

// The URI an idQ prefix stands for. mso and msox are well known; anything else is declared on
// the document's root element.
function VtResolveUri(Doc: Variant; const Prefix: String): String;
var
  A: Variant;
begin
  Result := '';
  if Prefix = 'mso'  then begin Result := VT_MSO_NS;  Exit; end;
  if Prefix = 'msox' then begin Result := VT_MSOX_NS; Exit; end;
  if Prefix = '' then Exit;
  A := Doc.documentElement.getAttributeNode('xmlns:' + Prefix);
  // A is a NODE. Handing it to VtStr threw "could not convert variant of type (Dispatch) into
  // type (OleStr)" and took the whole run with it - measured 8/13/2026, first install.
  if VarIsNull(A) or VarIsEmpty(A) then Exit;
  Result := VtStr(A.text);
end;

// Reuse whatever prefix already means this URI; failing that declare the preferred one; and if
// that name is taken for something else, vt1, vt2 ...
//
// Resolving by URI and not by prefix NAME is the point. The preferred prefix may already mean
// something else on the machine - Jerry's own build box has x1 bound to SWIFT - and writing
// VistaType's entries under it there would point them at SWIFT.
function VtEnsurePrefix(Doc: Variant; const Uri, Preferred: String): String;
var
  Root, A: Variant;
  I, N: Integer;
  Nm, Attempt: String;
begin
  Root := Doc.documentElement;
  for I := 0 to Root.attributes.length - 1 do
  begin
    A := Root.attributes.item[I];
    Nm := VtStr(A.nodeName);
    if VtStartsWith(Nm, 'xmlns:') and (VtStr(A.text) = Uri) then
    begin
      Result := Copy(Nm, 7, Length(Nm));
      Exit;
    end;
  end;

  Attempt := Preferred;
  N := 1;
  A := Root.getAttributeNode('xmlns:' + Attempt);
  while not (VarIsNull(A) or VarIsEmpty(A)) do
  begin
    Attempt := 'vt' + IntToStr(N);
    N := N + 1;
    A := Root.getAttributeNode('xmlns:' + Attempt);
  end;

  A := Doc.createNode(2, 'xmlns:' + Attempt, VT_XMLNS_NS);
  A.text := Uri;
  Root.setAttributeNode(A);
  Result := Attempt;
end;

procedure VtPrefixMapClear;
begin
  if VtPmUris = nil then VtPmUris := TStringList.Create;
  if VtPmPrefixes = nil then VtPmPrefixes := TStringList.Create;
  VtPmUris.Clear;
  VtPmPrefixes.Clear;
end;

procedure VtPrefixMapAdd(const Uri, Prefix: String);
begin
  if VtPmUris.IndexOf(Uri) >= 0 then Exit;
  VtPmUris.Add(Uri);
  VtPmPrefixes.Add(Prefix);
end;

function VtPrefixFor(const Uri: String): String;
var
  I: Integer;
begin
  Result := '';
  I := VtPmUris.IndexOf(Uri);
  if I >= 0 then Result := VtPmPrefixes[I];
end;

// ---------------------------------------------------------------------------------------
// documents
// ---------------------------------------------------------------------------------------

function VtNewDoc(): Variant;
begin
  Result := CreateOleObject('MSXML2.DOMDocument.6.0');
  Result.async := False;
  Result.validateOnParse := False;
  Result.resolveExternals := False;
  Result.preserveWhiteSpace := True;
  Result.setProperty('SelectionLanguage', 'XPath');
  Result.setProperty('SelectionNamespaces', 'xmlns:mso="' + VT_MSO_NS + '"');
end;

function VtParseFile(const Path: String; var Doc: Variant): Boolean;
begin
  Result := False;
  Doc := VtNewDoc();
  if not Doc.load(Path) then Exit;
  if VarIsNull(Doc.documentElement) then Exit;
  Result := True;
end;

// UTF-8, no byte-order mark, with the same declaration the PowerShell wrote. Measured on the
// build box: MSXML's save produces exactly that.
procedure VtSaveDoc(Doc: Variant; const Path: String);
var
  PI: Variant;
begin
  if Doc.firstChild.nodeType <> 7 then
  begin
    PI := Doc.createProcessingInstruction('xml', 'version="1.0" encoding="utf-8"');
    Doc.insertBefore(PI, Doc.firstChild);
  end;
  Doc.save(Path);
end;

function VtGetOrCreate(Doc, Parent: Variant; const Name: String): Variant;
begin
  Result := Parent.selectSingleNode('mso:' + Name);
  if VarIsNull(Result) then
  begin
    Result := Doc.createNode(1, 'mso:' + Name, VT_MSO_NS);
    Parent.appendChild(Result);
  end;
end;

// Copy a template element, and everything under it, into another document. Recursive: a tab is
// three deep (tab > group > control); a toolbar entry is flat.
//
// Nothing is imported wholesale - each element is built fresh in the destination - so an idQ
// can be rewritten to whatever prefix that document uses for the namespace, and no dangling
// namespace declaration can travel with it.
function VtCopyTree(DestDoc, SrcDoc, Item: Variant): Variant;
var
  El, A, C: Variant;
  I: Integer;
  Nm, Val, Uri, Pfx: String;
begin
  El := DestDoc.createNode(1, 'mso:' + VtStr(Item.baseName), VT_MSO_NS);

  for I := 0 to Item.attributes.length - 1 do
  begin
    A := Item.attributes.item[I];
    Nm := VtStr(A.nodeName);
    if VtStartsWith(Nm, 'xmlns:') or (Nm = 'xmlns') or (Nm = VT_KEEP) then Continue;
    Val := VtStr(A.text);
    if Nm = 'idQ' then
    begin
      Uri := VtResolveUri(SrcDoc, VtIdQPrefix(Val));
      if Uri <> '' then
      begin
        Pfx := VtPrefixFor(Uri);
        if Pfx <> '' then Val := Pfx + ':' + VtIdQLocal(Val);
      end;
    end;
    El.setAttribute(Nm, Val);
  end;

  for I := 0 to Item.childNodes.length - 1 do
  begin
    C := Item.childNodes.item[I];
    if VtIsElement(C) then El.appendChild(VtCopyTree(DestDoc, SrcDoc, C));
  end;

  Result := El;
end;

// Load one of the shipped templates and pin its x1 namespace to this machine's add-in path.
// HolderName is "sharedControls" for a toolbar, "tabs" for the ribbon tabs. Hands back the
// holder element; the caller walks its live children.
function VtLoadTemplate(const Path, HolderName: String; var Doc, Holder: Variant): Boolean;
var
  Xml: String;
begin
  Result := False;
  Holder := Null;
  if Path = '' then Exit;
  if not FileExists(Path) then
  begin
    VtLog('  !! template not found: ' + Path);
    Exit;
  end;
  if not VtParseFile(Path, Doc) then
  begin
    VtLog('  !! could not read template: ' + Path);
    Exit;
  end;

  // The placeholder sits in the root xmlns:x1 declaration. Round-tripping through the
  // document's own serialization keeps the encoding right whatever the file holds.
  Xml := VtStr(Doc.xml);
  if Pos('__VT_DOTM_PATH__', Xml) > 0 then
  begin
    StringChangeEx(Xml, '__VT_DOTM_PATH__', VtDotmPath, True);
    Doc := VtNewDoc();
    Doc.loadXML(Xml);
    if VarIsNull(Doc.documentElement) then
    begin
      VtLog('  !! template would not re-read after substitution: ' + Path);
      Exit;
    end;
  end;

  Holder := Doc.selectSingleNode('//mso:' + HolderName);
  if VarIsNull(Holder) then
  begin
    VtLog('  !! no ' + HolderName + ' in ' + Path);
    Exit;
  end;
  Result := True;
end;

// Every namespace the template declares, mapped onto a prefix free in the target document.
procedure VtBuildPrefixMap(SrcDoc, DestDoc: Variant);
var
  I: Integer;
  A: Variant;
  Nm, Uri: String;
begin
  for I := 0 to SrcDoc.documentElement.attributes.length - 1 do
  begin
    A := SrcDoc.documentElement.attributes.item[I];
    Nm := VtStr(A.nodeName);
    if not VtStartsWith(Nm, 'xmlns:') then Continue;
    Uri := VtStr(A.text);
    if Uri = VT_MSO_NS then Continue;          // mso is always mso
    VtPrefixMapAdd(Uri, VtEnsurePrefix(DestDoc, Uri, Copy(Nm, 7, Length(Nm))));
  end;
end;

// Remove every element child of a node. Backwards, so the indexes still to come cannot move.
procedure VtClearElements(Node: Variant);
var
  I: Integer;
  C: Variant;
begin
  for I := Node.childNodes.length - 1 downto 0 do
  begin
    C := Node.childNodes.item[I];
    if VtIsElement(C) then Node.removeChild(C);
  end;
end;

// ---------------------------------------------------------------------------------------
// recognizing what is ours
// ---------------------------------------------------------------------------------------

// The add-in control names a tab or group carries. This is the fingerprint that survives Word
// rewriting ids: Word can renumber a tab or a group, but it cannot change which of the add-in's
// controls they point at without the user editing the group's contents.
procedure VtOurControlNames(Doc, El: Variant; Names: TStringList);
var
  L: Variant;
  I: Integer;
  IdQ, Uri: String;
begin
  Names.Clear;
  L := El.selectNodes('.//*[@idQ]');
  for I := 0 to L.length - 1 do
  begin
    IdQ := VtStr(L.item[I].getAttribute('idQ'));
    Uri := VtResolveUri(Doc, VtIdQPrefix(IdQ));
    if (Uri <> '') and VtEndsWith(Uppercase(Uri), 'LPANDBRL.DOTM') then
      Names.Add(VtIdQLocal(IdQ));
  end;
end;

function VtHasOurControls(Doc, El: Variant): Boolean;
var
  Names: TStringList;
begin
  Names := TStringList.Create;
  try
    VtOurControlNames(Doc, El, Names);
    Result := Names.Count > 0;
  finally
    Names.Free;
  end;
end;

// A group id with VistaType's decoration taken off, so an old one and a current one compare
// equal. The generator built "vt_grp_mso_c1_18B5F8FF" from the id Word itself had written,
// "mso_c1.18B5F8FF" - same hash, prefixed, dots turned into underscores.
function VtNormalizedGroupId(const Id: String): String;
begin
  Result := Id;
  if VtStartsWith(Result, 'vt_grp_') then Result := Copy(Result, 8, Length(Result));
  StringChangeEx(Result, '.', '_', True);
end;

// Is this group one of ours? Three ways, and ALL THREE are needed:
//   * the id we write,
//   * the add-in controls it holds (Word may have renumbered the id),
//   * the same id as one of the template's groups with the decoration off - the only thing
//     that sees a PRE-3.0 group, because those hold no resolvable controls at all.
// A group the user added satisfies none of them and is never touched.
function VtIsOurGroup(Doc, G: Variant; TemplateGroupIds: TStringList): Boolean;
begin
  Result := True;
  if VtStartsWith(VtStr(G.getAttribute('id')), 'vt_grp_') then Exit;
  if VtHasOurControls(Doc, G) then Exit;
  if TemplateGroupIds <> nil then
    if TemplateGroupIds.IndexOf(VtNormalizedGroupId(VtStr(G.getAttribute('id')))) >= 0 then Exit;
  Result := False;
end;

// Is this tab a VistaType tab from BEFORE 3.0, when the install replaced the user's whole
// Word.officeUI? Those tabs are mso_c1.F9211 "Braille Macros", mso_c1.56E551D "VistaType LP"
// and mso_c1.4EA2EBA "LP and BRL QAT Icons" - Word-generated ids, holding controls that no
// longer resolve to the add-in. Every machine that ran VistaType before 3.0 still carries
// them, so on upgrade the transcriber gets two tabs of each name and the older one is missing
// every button added since. Jerry hit that on 8/6/2026; it took eight builds to find.
//
// The group ids are the proof. Only say yes when EVERY group matches and there are at least
// two, so a tab of the user's own that happens to share one group is never touched.
function VtIsLegacyOurTab(Tab: Variant; TemplateGroupIds: TStringList): Boolean;
var
  I, NGroups: Integer;
  C: Variant;
begin
  Result := False;
  if TemplateGroupIds = nil then Exit;
  NGroups := 0;
  for I := 0 to Tab.childNodes.length - 1 do
  begin
    C := Tab.childNodes.item[I];
    if not VtIsElement(C) then Continue;
    if VtStr(C.baseName) <> 'group' then Continue;
    NGroups := NGroups + 1;
    if TemplateGroupIds.IndexOf(VtNormalizedGroupId(VtStr(C.getAttribute('id')))) < 0 then Exit;
  end;
  Result := NGroups >= 2;
end;

// A tab is ours if it carries our id shape, or any control pointing into the add-in.
function VtIsOurTab(Doc, Tab: Variant): Boolean;
begin
  Result := VtStartsWith(VtStr(Tab.getAttribute('id')), 'vt_tab_') or VtHasOurControls(Doc, Tab);
end;

// Strip the groups we own from a tab. Backwards, for the usual reason.
procedure VtStripOurGroups(Doc, Tab: Variant; TemplateGroupIds: TStringList);
var
  I: Integer;
  C: Variant;
begin
  for I := Tab.childNodes.length - 1 downto 0 do
  begin
    C := Tab.childNodes.item[I];
    if not VtIsElement(C) then Continue;
    if VtIsOurGroup(Doc, C, TemplateGroupIds) then Tab.removeChild(C);
  end;
end;

// ---------------------------------------------------------------------------------------
// taking previous VistaType work back out
// ---------------------------------------------------------------------------------------

function VtRemoveOurTabs(Doc, TabsNode: Variant): Integer;
var
  I: Integer;
  Tab: Variant;
begin
  Result := 0;
  for I := TabsNode.childNodes.length - 1 downto 0 do
  begin
    Tab := TabsNode.childNodes.item[I];
    if not VtIsElement(Tab) then Continue;
    if VtStr(Tab.baseName) <> 'tab' then Continue;
    if not VtIsOurTab(Doc, Tab) then Continue;

    VtStripOurGroups(Doc, Tab, nil);

    if VtCountElements(Tab) = 0 then
    begin
      TabsNode.removeChild(Tab);
      Result := Result + 1;
    end
    else
      VtLog('  kept ' + VtStr(Tab.getAttribute('id')) + ': it still holds a group the user added');
  end;
end;

// The toolbar entries a previous VistaType install laid down, for Mine and Restore to strip
// before appending fresh ones. FullHolder may be Null when the curated template is missing.
function VtRemoveOurEntries(Doc, Shared, FullHolder: Variant): Integer;
var
  I, J, N: Integer;
  IdQ, Uri: String;
  Want, Have: TStringList;
  PrefixMatches: Boolean;
  C: Variant;
begin
  Result := 0;

  // 1) Anything whose namespace is the add-in itself is unambiguously ours.
  for I := Shared.childNodes.length - 1 downto 0 do
  begin
    C := Shared.childNodes.item[I];
    if not VtIsElement(C) then Continue;
    IdQ := VtStr(C.getAttribute('idQ'));
    if IdQ = '' then Continue;
    Uri := VtResolveUri(Doc, VtIdQPrefix(IdQ));
    if (Uri <> '') and VtEndsWith(Uppercase(Uri), 'LPANDBRL.DOTM') then
    begin
      Shared.removeChild(C);
      Result := Result + 1;
    end;
  end;

  // 2) The built-in Word entries the curated toolbar imposes cannot be told from a user's own
  //    by inspection, so only strip them when the curated block is present IN FULL and in
  //    template order - a fingerprint strong enough to be sure we wrote it.
  if not VarIsNull(FullHolder) then
  begin
    Want := TStringList.Create;
    Have := TStringList.Create;
    try
      for I := 0 to FullHolder.childNodes.length - 1 do
      begin
        C := FullHolder.childNodes.item[I];
        if not VtIsElement(C) then Continue;
        IdQ := VtStr(C.getAttribute('idQ'));
        if (IdQ <> '') and (VtIdQPrefix(IdQ) = 'mso') then Want.Add(VtIdQLocal(IdQ));
      end;
      for I := 0 to Shared.childNodes.length - 1 do
      begin
        C := Shared.childNodes.item[I];
        if not VtIsElement(C) then Continue;
        IdQ := VtStr(C.getAttribute('idQ'));
        if (IdQ <> '') and (VtIdQPrefix(IdQ) = 'mso') then Have.Add(VtIdQLocal(IdQ));
      end;

      if (Want.Count > 0) and (Have.Count >= Want.Count) then
      begin
        PrefixMatches := True;
        for I := 0 to Want.Count - 1 do
          if Have[I] <> Want[I] then
          begin
            PrefixMatches := False;
            Break;
          end;

        if PrefixMatches then
        begin
          // Forwards here, taking the FIRST Want.Count of them out: the block sits at the
          // front, and each removal shifts the rest down, so the index stays where it is.
          N := 0;
          J := 0;
          while (N < Want.Count) and (J < Shared.childNodes.length) do
          begin
            C := Shared.childNodes.item[J];
            if not VtIsElement(C) then
            begin
              J := J + 1;
              Continue;
            end;
            IdQ := VtStr(C.getAttribute('idQ'));
            if (IdQ <> '') and (VtIdQPrefix(IdQ) = 'mso') then
            begin
              Shared.removeChild(C);
              Result := Result + 1;
              N := N + 1;
            end
            else
              J := J + 1;
          end;
          VtLog('  recognized VistaType''s standard toolbar and took it back out');
        end;
      end;
    finally
      Want.Free;
      Have.Free;
    end;
  end;

  // 3) Our separators, in the special namespace, only ever come from us.
  for I := Shared.childNodes.length - 1 downto 0 do
  begin
    C := Shared.childNodes.item[I];
    if not VtIsElement(C) then Continue;
    if VtStr(C.baseName) <> 'separator' then Continue;
    IdQ := VtStr(C.getAttribute('idQ'));
    if IdQ = '' then Continue;
    IdQ := VtIdQLocal(IdQ);
    if VtStartsWith(IdQ, 'vtsep') or VtStartsWith(IdQ, 'sep') then
    begin
      Shared.removeChild(C);
      Result := Result + 1;
    end;
  end;
end;

// ---------------------------------------------------------------------------------------
// the install side  (was Merge-Qat.ps1)
// ---------------------------------------------------------------------------------------

function VtFileHasContent(const Path: String): Boolean;
var
  Size: Integer;
begin
  Result := FileExists(Path) and FileSize(Path, Size) and (Size > 0);
end;

// What this install wrote, so uninstall can take out precisely that and nothing else.
procedure VtWriteManifest(const Target: String; Items, TabIds: TStringList);
var
  L: TStringList;
  I: Integer;
begin
  L := TStringList.Create;
  try
    L.Add('# VistaType manifest -- what this install wrote. Do not edit.');
    L.Add('mode=' + VtMode);
    L.Add('tabs=' + VtTabs);
    L.Add('written=' + GetDateTimeString('yyyy-mm-dd hh:nn:ss', '-', ':'));
    for I := 0 to Items.Count - 1 do L.Add('item=' + Items[I]);
    // Tabs carry id, not idQ, so they get their own line type.
    for I := 0 to TabIds.Count - 1 do L.Add('tab=' + TabIds[I]);
    L.SaveToFile(Target + '.vtqatmanifest');
  finally
    L.Free;
  end;
end;

// The manifest key for an element we have just written into Doc.
function VtManifestKey(Doc, El: Variant): String;
var
  IdQ: String;
begin
  Result := '';
  IdQ := VtStr(El.getAttribute('idQ'));
  if IdQ = '' then Exit;
  Result := VtStr(El.baseName) + '|' + VtResolveUri(Doc, VtIdQPrefix(IdQ)) + '|' + VtIdQLocal(IdQ);
end;

procedure VtDoTabs(Doc, Ribbon: Variant; const TabsTpl: String; WrittenTabs: TStringList);
var
  TabsNode, TabsDoc, TabsHolder, T, E, Winner: Variant;
  WantNames, HaveNames, WantGroups: TStringList;
  I, J, N, Overlap, Score, BestScore: Integer;
  Installed, Refreshed, Duplicates: Integer;
  WantId: String;
begin
  TabsNode := VtGetOrCreate(Doc, Ribbon, 'tabs');

  if VtTabs = 'Remove' then
  begin
    VtLog('  ribbon tabs: removed ' + IntToStr(VtRemoveOurTabs(Doc, TabsNode)));
    Exit;
  end;
  if VtTabs <> 'Install' then Exit;

  if not VtLoadTemplate(TabsTpl, 'tabs', TabsDoc, TabsHolder) then Exit;
  VtBuildPrefixMap(TabsDoc, Doc);

  Installed := 0; Refreshed := 0; Duplicates := 0;

  for I := 0 to TabsHolder.childNodes.length - 1 do
  begin
    T := TabsHolder.childNodes.item[I];
    if not VtIsElement(T) then Continue;
    WantId := VtStr(T.getAttribute('id'));

    WantNames  := TStringList.Create;
    HaveNames  := TStringList.Create;
    WantGroups := TStringList.Create;
    try
      VtOurControlNames(TabsDoc, T, WantNames);

      // This tab's group ids with the decoration off, which is what identifies both a
      // Word-renumbered copy and a pre-3.0 one.
      for J := 0 to T.childNodes.length - 1 do
      begin
        E := T.childNodes.item[J];
        if VtIsElement(E) and (VtStr(E.baseName) = 'group') then
          WantGroups.Add(VtNormalizedGroupId(VtStr(E.getAttribute('id'))));
      end;

      // Pass one, read only: which of the tabs already there answers to this template tab,
      // and which answers best. The winner is marked rather than remembered, because holding
      // a node while the list is edited is what crashed the first version of this file.
      Winner := Null;
      BestScore := -1;
      for J := 0 to TabsNode.childNodes.length - 1 do
      begin
        E := TabsNode.childNodes.item[J];
        if not VtIsElement(E) then Continue;
        if VtStr(E.baseName) <> 'tab' then Continue;

        Score := -1;
        if VtStr(E.getAttribute('id')) = WantId then
          Score := 1000
        else
        begin
          VtOurControlNames(Doc, E, HaveNames);
          if HaveNames.Count > 0 then
          begin
            Overlap := 0;
            for N := 0 to HaveNames.Count - 1 do
              if WantNames.IndexOf(HaveNames[N]) >= 0 then Overlap := Overlap + 1;
            if Overlap > 0 then Score := Overlap;
          end
          else if VtIsLegacyOurTab(E, WantGroups) then
          begin
            VtLog('  found a pre-3.0 VistaType tab: ' + VtStr(E.getAttribute('id')) +
                  ' ''' + VtStr(E.getAttribute('label')) + '''');
            Score := 1;
          end;
        end;

        if Score >= 0 then
        begin
          E.setAttribute(VT_KEEP, IntToStr(Score));   // marks it as a candidate
          if Score > BestScore then
          begin
            BestScore := Score;
            Winner := E;
          end;
        end;
      end;

      if not VarIsNull(Winner) then
      begin
        // Refresh the groups we own and leave everything else - the tab element itself, its
        // place, its name, and any group the user added.
        //
        // WantGroups is passed so a PRE-3.0 group is recognized too. Leaving it out was the
        // 3.0.96 disaster: the legacy groups were not removed, our seven appended beside
        // them, and the tab came out with fourteen groups, half of them drawing as empty
        // placeholders because their controls point nowhere.
        VtStripOurGroups(Doc, Winner, WantGroups);
        for J := 0 to T.childNodes.length - 1 do
        begin
          E := T.childNodes.item[J];
          if VtIsElement(E) then Winner.appendChild(VtCopyTree(Doc, TabsDoc, E));
        end;
        Winner.setAttribute(VT_KEEP, 'winner');
        WrittenTabs.Add(WantId);
        Refreshed := Refreshed + 1;

        // Pass two: every other candidate is a spare copy. Strip our groups from it, and if
        // that leaves it empty take the tab out. A group the user added keeps it.
        for J := TabsNode.childNodes.length - 1 downto 0 do
        begin
          E := TabsNode.childNodes.item[J];
          if not VtIsElement(E) then Continue;
          if VtStr(E.getAttribute(VT_KEEP)) = '' then Continue;
          if VtStr(E.getAttribute(VT_KEEP)) = 'winner' then Continue;
          VtStripOurGroups(Doc, E, WantGroups);
          if VtCountElements(E) = 0 then
          begin
            TabsNode.removeChild(E);
            Duplicates := Duplicates + 1;
            VtLog('  removed a duplicate ' + WantId + ' tab');
          end
          else
            VtLog('  emptied a duplicate ' + WantId + ' tab but kept it: it holds a group the user added');
        end;
      end
      else
      begin
        TabsNode.appendChild(VtCopyTree(Doc, TabsDoc, T));
        WrittenTabs.Add(WantId);
        Installed := Installed + 1;
      end;

      // The marker is scratch, and must never reach the file.
      for J := 0 to TabsNode.childNodes.length - 1 do
      begin
        E := TabsNode.childNodes.item[J];
        if VtIsElement(E) then E.removeAttribute(VT_KEEP);
      end;
    finally
      WantNames.Free;
      HaveNames.Free;
      WantGroups.Free;
    end;
  end;

  VtLog('  ribbon tabs: ' + IntToStr(Installed) + ' added, ' + IntToStr(Refreshed) +
        ' refreshed, ' + IntToStr(Duplicates) + ' duplicate(s) removed');
end;

procedure VtApplyOne(const Target, FullTpl, IconsTpl, TabsTpl: String);
var
  Bak, Prev, Dir, IdQ, Uri, K: String;
  Doc, BakDoc, BakShared, Root, Ribbon, Qat, Shared, El, C: Variant;
  TplDoc, TplHolder, FullDoc, FullHolder: Variant;
  Written, WrittenTabs, Have: TStringList;
  I, N, Added, Stripped: Integer;
  RestoreFromBackup: Boolean;
begin
  Bak  := Target + '.vtqatbak';
  Prev := Target + '.vtqatprev';

  Dir := ExtractFilePath(Target);
  if not DirExists(Dir) then ForceDirectories(Dir);

  // The pristine pre-VistaType toolbar, captured ONCE and never overwritten. On a machine that
  // has had VistaType for years this is still their true original - it is what makes "put back
  // the toolbar I had" possible at all. Never delete it.
  if not FileExists(Bak) then
  begin
    if FileExists(Target) then FileCopy(Target, Bak, False)
    else SaveStringToFile(Bak, '', False);   // empty = there was no toolbar file before us
  end;
  // A snapshot of whatever was there a moment ago, so one install is undoable by hand.
  if FileExists(Target) then FileCopy(Target, Prev, False);

  // Restore is deliberately NOT done by copying the backup over the file. The backup predates
  // every ribbon change the user has made since installing - their own tabs and groups, and
  // from 3.0.34 VistaType's tabs too - so overwriting would silently discard all of it. The
  // backup's toolbar entries are grafted into the LIVE document instead. (3.0.33 overwrote;
  // that was wrong.)
  RestoreFromBackup := (VtMode = 'Restore') and VtFileHasContent(Bak);
  if (VtMode = 'Restore') and not RestoreFromBackup then
    VtLog('  no pre-VistaType toolbar was saved for ' + Target + '; starting from what is there');

  if FileExists(Target) then
  begin
    if not VtParseFile(Target, Doc) then
    begin
      VtLog('  !! ' + Target + ' could not be read as XML; leaving it alone');
      Exit;
    end;
  end
  else
  begin
    Doc := VtNewDoc();
    Doc.loadXML('<mso:customUI xmlns:mso="' + VT_MSO_NS + '"><mso:ribbon><mso:qat>' +
                '<mso:sharedControls></mso:sharedControls></mso:qat></mso:ribbon></mso:customUI>');
  end;

  Root   := Doc.documentElement;
  Ribbon := VtGetOrCreate(Doc, Root,   'ribbon');
  Qat    := VtGetOrCreate(Doc, Ribbon, 'qat');
  Shared := VtGetOrCreate(Doc, Qat,    'sharedControls');

  // Reserve prefixes for the namespaces VistaType writes, whatever the user already has.
  VtPrefixMapClear;
  VtPrefixMapAdd(VtDotmPath, VtEnsurePrefix(Doc, VtDotmPath, 'x1'));
  VtPrefixMapAdd(VT_MSOX_NS, VtEnsurePrefix(Doc, VT_MSOX_NS, 'msox'));

  Written     := TStringList.Create;
  WrittenTabs := TStringList.Create;
  try
    // ---- Restore: swap the live toolbar entries for the saved ones, in place -------------
    if RestoreFromBackup then
    begin
      if VtParseFile(Bak, BakDoc) then
      begin
        BakShared := BakDoc.selectSingleNode('//mso:sharedControls');
        if not VarIsNull(BakShared) then
        begin
          VtBuildPrefixMap(BakDoc, Doc);
          VtClearElements(Shared);
          N := 0;
          for I := 0 to BakShared.childNodes.length - 1 do
          begin
            C := BakShared.childNodes.item[I];
            if not VtIsElement(C) then Continue;
            Shared.appendChild(VtCopyTree(Doc, BakDoc, C));
            N := N + 1;
          end;
          VtLog('  put back ' + IntToStr(N) + ' saved toolbar entry(ies) for ' + Target +
                ', leaving every ribbon change alone');
        end
        else
          VtLog('  the saved copy for ' + Target + ' has no toolbar section; nothing to put back');
      end;
    end;

    // ---- the toolbar --------------------------------------------------------------------
    if VtMode = 'None' then
      VtLog('  leaving the Quick Access Toolbar untouched, as chosen')

    else if VtMode = 'Vista' then
    begin
      // Whole replacement. No merging: their toolbar is saved, not blended.
      if VtLoadTemplate(FullTpl, 'sharedControls', TplDoc, TplHolder) then
      begin
        VtBuildPrefixMap(TplDoc, Doc);
        VtClearElements(Shared);
        N := 0;
        for I := 0 to TplHolder.childNodes.length - 1 do
        begin
          C := TplHolder.childNodes.item[I];
          if not VtIsElement(C) then Continue;
          K := VtManifestKey(TplDoc, C);
          Shared.appendChild(VtCopyTree(Doc, TplDoc, C));
          if K <> '' then Written.Add(K);
          N := N + 1;
        end;
        VtLog('  installed VistaType''s standard toolbar (' + IntToStr(N) + ' entries) into ' + Target);
      end;
    end

    else
    begin
      // Mine / Restore: keep what is there, take out anything a previous VistaType install put
      // in, then append our icons on the end.
      FullHolder := Null;
      VtLoadTemplate(FullTpl, 'sharedControls', FullDoc, FullHolder);

      Stripped := VtRemoveOurEntries(Doc, Shared, FullHolder);
      if Stripped > 0 then
        VtLog('  removed ' + IntToStr(Stripped) + ' entry(ies) left by a previous VistaType install');

      if VtLoadTemplate(IconsTpl, 'sharedControls', TplDoc, TplHolder) then
      begin
        VtBuildPrefixMap(TplDoc, Doc);

        Have := TStringList.Create;
        try
          for I := 0 to Shared.childNodes.length - 1 do
          begin
            C := Shared.childNodes.item[I];
            if not VtIsElement(C) then Continue;
            IdQ := VtStr(C.getAttribute('idQ'));
            if IdQ = '' then Continue;
            Have.Add(VtResolveUri(Doc, VtIdQPrefix(IdQ)) + '|' + VtIdQLocal(IdQ));
          end;

          Added := 0;
          for I := 0 to TplHolder.childNodes.length - 1 do
          begin
            C := TplHolder.childNodes.item[I];
            if not VtIsElement(C) then Continue;
            IdQ := VtStr(C.getAttribute('idQ'));
            if IdQ <> '' then
            begin
              Uri := VtResolveUri(TplDoc, VtIdQPrefix(IdQ));
              if Have.IndexOf(Uri + '|' + VtIdQLocal(IdQ)) >= 0 then Continue;
            end;
            { The key comes from the TEMPLATE node, before anything is appended - see the
              note at the top of this file about what happens to a node handed back by
              VtCopyTree once appendChild has taken it. The manifest records the namespace
              URI, and VtCopyTree only ever rewrites the PREFIX, so the two agree. }
            K := VtManifestKey(TplDoc, C);
            Shared.appendChild(VtCopyTree(Doc, TplDoc, C));
            if K <> '' then Written.Add(K);
            Added := Added + 1;
          end;
        finally
          Have.Free;
        end;
        VtLog('  kept the existing toolbar and added ' + IntToStr(Added) + ' VistaType icon(s) to ' + Target);
      end;
    end;

    // ---- the ribbon tabs -----------------------------------------------------------------
    // Word does not list add-in tabs in Customize the Ribbon, so tabs defined inside the add-in
    // cannot be hidden, reordered or renamed. Written here they are ordinary custom tabs and
    // all three become possible.
    //
    // The policy, worth stating plainly: THE CONTENTS OF OUR TABS ARE OURS; THE TAB'S PLACE,
    // NAME AND VISIBILITY ARE THE USER'S. An upgrade refreshes the buttons inside an existing
    // VistaType tab and touches nothing else - a tab they moved, renamed or unticked stays
    // moved, renamed and unticked.
    if VtTabs <> 'Skip' then VtDoTabs(Doc, Ribbon, TabsTpl, WrittenTabs);

    VtSaveDoc(Doc, Target);
    VtWriteManifest(Target, Written, WrittenTabs);
  finally
    Written.Free;
    WrittenTabs.Free;
  end;
end;

function VtTargets(): TStringList;
begin
  Result := TStringList.Create;
  Result.Add(ExpandConstant('{userappdata}\Microsoft\Office\Word.officeUI'));
  Result.Add(ExpandConstant('{localappdata}\Microsoft\Office\Word.officeUI'));
end;

procedure VtQatApply(const Mode, Tabs: String);
var
  T: TStringList;
  I: Integer;
  Base: String;
begin
  VtMode := Mode;
  VtTabs := Tabs;
  VtDotmPath := ExpandConstant('{userappdata}\Microsoft\Word\STARTUP\LPandBRL.dotm');
  Base := ExpandConstant('{userappdata}\VistaType LP\');

  VtLog('Qat setup -Mode ' + Mode + ' -Tabs ' + Tabs + '  (in the installer, no PowerShell)');

  if (Mode = 'None') and (Tabs = 'Skip') then
  begin
    VtLog('  nothing to do: toolbar and ribbon both left alone, as chosen');
    Exit;
  end;

  T := VtTargets;
  try
    for I := 0 to T.Count - 1 do
    begin
      try
        VtApplyOne(T[I], Base + 'qat-template.officeUI', Base + 'qat-icons-only.officeUI',
                   Base + 'ribbon-tabs.officeUI');
      except
        // One bad or unreadable file must not abort the install, and must not stop the other
        // location being written - Word only reads one of the two.
        VtLog('  !! ' + T[I] + ' : ' + GetExceptionMessage);
      end;
    end;
  finally
    T.Free;
  end;

  // Tell the add-in whether the user now has their own copies of the tabs. VtTabVisible in
  // RibbonCallbacks.bas reads this: when it is 1 the tabs built into the add-in go dark, so the
  // two never appear at once. Absent - a hand-copied install, or the option declined - means
  // the built-in tabs show as they always have, and nobody ends up with no tabs.
  if Tabs <> 'Skip' then
  begin
    try
      if Tabs = 'Install' then
        RegWriteStringValue(HKEY_CURRENT_USER, 'Software\VistaType LP', 'UserRibbonTabs', '1')
      else
        RegWriteStringValue(HKEY_CURRENT_USER, 'Software\VistaType LP', 'UserRibbonTabs', '0');
      VtLog('  UserRibbonTabs recorded');
    except
      VtLog('  !! could not record UserRibbonTabs: ' + GetExceptionMessage);
    end;
  end;
end;

// ---------------------------------------------------------------------------------------
// the uninstall side  (was Remove-Qat.ps1)
// ---------------------------------------------------------------------------------------
// What it will NOT do, and why: an older version copied the pre-install backup over the live
// file wholesale. That reverted the toolbar to its state on the day VistaType was installed,
// silently discarding every icon - and every ribbon change - the user had added in the months
// since. An uninstaller has no business undoing work that was never ours.

procedure VtCleanOne(const Target: String);
var
  Bak, Man, Prev, Mode, K, IdQ, Uri: String;
  Doc, BakDoc, Shared, BakShared, TabsNode, C, Imported, After: Variant;
  OurKeys, Kept, ManLines: TStringList;
  I, Removed, TabsGone, Restored, LeftCount: Integer;
  IsOurs, HaveBackup: Boolean;
begin
  Bak  := Target + '.vtqatbak';
  Man  := Target + '.vtqatmanifest';
  Prev := Target + '.vtqatprev';

  if not FileExists(Target) then
  begin
    VtLog('  ' + Target + ' : nothing there');
    Exit;
  end;

  OurKeys := TStringList.Create;
  Kept    := TStringList.Create;
  try
    Mode := 'unknown';
    if FileExists(Man) then
    begin
      ManLines := TStringList.Create;
      try
        ManLines.LoadFromFile(Man);
        for I := 0 to ManLines.Count - 1 do
        begin
          if VtStartsWith(ManLines[I], 'mode=') then Mode := Copy(ManLines[I], 6, Length(ManLines[I]));
          if VtStartsWith(ManLines[I], 'item=') then OurKeys.Add(Copy(ManLines[I], 6, Length(ManLines[I])));
        end;
      finally
        ManLines.Free;
      end;
      VtLog('  ' + Target + ' : manifest found (installed as ''' + Mode + ''', ' +
            IntToStr(OurKeys.Count) + ' entries)');
    end
    else
      VtLog('  ' + Target + ' : no manifest; identifying our entries by add-in path');

    if not VtParseFile(Target, Doc) then
    begin
      VtLog('  !! ' + Target + ' could not be read as XML; leaving it alone');
      Exit;
    end;

    // Ribbon tabs first, and INDEPENDENTLY of the toolbar. Someone who chose "leave my toolbar
    // alone" still has our tabs, and giving up because there is no toolbar section would
    // strand them on the ribbon for good.
    TabsGone := 0;
    TabsNode := Doc.selectSingleNode('//mso:ribbon/mso:tabs');
    if not VarIsNull(TabsNode) then TabsGone := VtRemoveOurTabs(Doc, TabsNode);

    Shared := Doc.selectSingleNode('//mso:qat/mso:sharedControls');
    if VarIsNull(Shared) then
    begin
      if TabsGone > 0 then
      begin
        VtSaveDoc(Doc, Target);
        VtLog('  ' + Target + ' : removed ' + IntToStr(TabsGone) +
              ' VistaType ribbon tab(s); no toolbar section to clean');
      end
      else
        VtLog('  ' + Target + ' : nothing of ours here');
      DeleteFile(Man);
      DeleteFile(Prev);
      Exit;
    end;

    Removed := 0;
    for I := Shared.childNodes.length - 1 downto 0 do
    begin
      C := Shared.childNodes.item[I];
      if not VtIsElement(C) then Continue;
      IsOurs := False;
      if OurKeys.Count > 0 then
      begin
        K := VtManifestKey(Doc, C);
        IsOurs := (K <> '') and (OurKeys.IndexOf(K) >= 0);
      end
      else
      begin
        // Fallback: our controls are the ones whose namespace IS the add-in, plus the legacy
        // VT_ markers from an older scheme.
        IdQ := VtStr(C.getAttribute('idQ'));
        if IdQ <> '' then
        begin
          IsOurs := VtStartsWith(IdQ, 'x1:VT_') or VtStartsWith(IdQ, 'msox:VT_');
          if not IsOurs then
          begin
            Uri := VtResolveUri(Doc, VtIdQPrefix(IdQ));
            IsOurs := (Uri <> '') and VtEndsWith(Uppercase(Uri), 'LPANDBRL.DOTM');
          end;
        end;
      end;
      if IsOurs then
      begin
        Shared.removeChild(C);
        Removed := Removed + 1;
      end;
    end;

    LeftCount  := VtCountElements(Shared);
    HaveBackup := VtFileHasContent(Bak);

    // VistaType had replaced their toolbar whole if the install was "Vista", or if taking our
    // entries out leaves nothing at all. Either way the saved original is what they want back.
    // They may have added icons on top of ours since; those must survive, so the saved entries
    // go back at the FRONT and anything of theirs stays where it is.
    //
    // ALWAYS work in the live document, never in the backup. The obvious shortcut - load the
    // backup, add the leftovers, save the backup over the target - loses everything in the live
    // file that is not a toolbar entry, which is every ribbon change since the install. 3.0.33
    // did exactly that.
    if HaveBackup and ((Mode = 'Vista') or (LeftCount = 0)) then
    begin
      if VtParseFile(Bak, BakDoc) then
      begin
        BakShared := BakDoc.selectSingleNode('//mso:qat/mso:sharedControls');
        if not VarIsNull(BakShared) then
        begin
          for I := 0 to Shared.childNodes.length - 1 do
          begin
            C := Shared.childNodes.item[I];
            if not VtIsElement(C) then Continue;
            K := VtManifestKey(Doc, C);
            if K <> '' then Kept.Add(K);
          end;

          VtPrefixMapClear;
          VtBuildPrefixMap(BakDoc, Doc);

          { The saved entries go back at the FRONT, in their original order, ahead of
            anything the user added on top of ours. Each is inserted immediately before
            whatever element was first to begin with, which keeps their order without ever
            holding on to a node VtCopyTree has just handed back - see the note at the top
            of this file. }
          After := Null;
          for I := 0 to Shared.childNodes.length - 1 do
            if VtIsElement(Shared.childNodes.item[I]) then
            begin
              After := Shared.childNodes.item[I];
              Break;
            end;

          Restored := 0;
          for I := 0 to BakShared.childNodes.length - 1 do
          begin
            C := BakShared.childNodes.item[I];
            if not VtIsElement(C) then Continue;
            K := VtManifestKey(BakDoc, C);
            if (K <> '') and (Kept.IndexOf(K) >= 0) then Continue;   // already there
            if VarIsNull(After) then Shared.appendChild(VtCopyTree(Doc, BakDoc, C))
            else Shared.insertBefore(VtCopyTree(Doc, BakDoc, C), After);
            Restored := Restored + 1;
          end;
          VtSaveDoc(Doc, Target);
          VtLog('  ' + Target + ' : put back ' + IntToStr(Restored) +
                ' saved toolbar entry(ies), keeping ' + IntToStr(LeftCount) +
                ' added since, and every ribbon change');
        end
        else
        begin
          // The backup has no toolbar section, so there is nothing to put back. Keep the live
          // document - ours already stripped - rather than copying the backup over it.
          VtSaveDoc(Doc, Target);
          VtLog('  ' + Target + ' : nothing saved to put back; removed our entries only');
        end;
      end;
    end
    else if (LeftCount = 0) and FileExists(Bak) and (not HaveBackup) then
    begin
      // Empty sentinel: there was no toolbar file before us and nothing of theirs is left, so
      // take the file away again. If they HAD added icons, LeftCount would not be zero and the
      // file is kept - the old version deleted it regardless and lost them.
      DeleteFile(Target);
      VtLog('  ' + Target + ' : removed the file VistaType created (nothing of the user''s in it)');
    end
    else
    begin
      VtSaveDoc(Doc, Target);
      VtLog('  ' + Target + ' : removed ' + IntToStr(Removed) + ' VistaType entry(ies); kept ' +
            IntToStr(LeftCount) + ' of the user''s');
    end;

    if TabsGone > 0 then
      VtLog('  ' + Target + ' : removed ' + IntToStr(TabsGone) + ' VistaType ribbon tab(s)');

    // Our own bookkeeping goes; their saved original stays. It is the only copy of the toolbar
    // they had before VistaType and it costs nothing to leave behind.
    DeleteFile(Man);
    DeleteFile(Prev);
    if FileExists(Bak) then
      VtLog('  ' + Target + ' : leaving ' + ExtractFileName(Bak) + ' in place (their pre-VistaType toolbar)');
  finally
    OurKeys.Free;
    Kept.Free;
  end;
end;

procedure VtQatRemove();
var
  T: TStringList;
  I: Integer;
begin
  VtDotmPath := ExpandConstant('{userappdata}\Microsoft\Word\STARTUP\LPandBRL.dotm');
  VtLog('Qat removal (in the uninstaller, no PowerShell)');

  // The add-in may well still be installed - Word can disable it, or the user may reinstall -
  // so tell it the user no longer has their own copies of the tabs. VtTabVisible then shows the
  // built-in ones again, rather than leaving the ribbon with no VistaType tabs at all.
  try
    if RegKeyExists(HKEY_CURRENT_USER, 'Software\VistaType LP') then
    begin
      RegWriteStringValue(HKEY_CURRENT_USER, 'Software\VistaType LP', 'UserRibbonTabs', '0');
      VtLog('  UserRibbonTabs = 0 (the add-in''s own tabs show again)');
    end;
  except
    VtLog('  !! could not clear UserRibbonTabs: ' + GetExceptionMessage);
  end;

  T := VtTargets;
  try
    for I := 0 to T.Count - 1 do
    begin
      try
        VtCleanOne(T[I]);
      except
        // A malformed toolbar file must not take the uninstall down with it, and must not stop
        // the other location being cleaned.
        VtLog('  !! ' + T[I] + ' : ' + GetExceptionMessage);
      end;
    end;
  finally
    T.Free;
  end;
end;
