/* Neurosurg Exam – scrittura di file .xlsx senza librerie esterne (zip non compresso + SpreadsheetML).
   XLSXW.build([{name, rows}]) → Blob. Celle: numero, stringa, oppure {d:'AAAA-MM-GG'} per una data. La prima riga è l'intestazione. */
window.XLSXW=(function(){
var enc=new TextEncoder();
var CRC=(function(){var t=new Uint32Array(256);for(var n=0;n<256;n++){var c=n;for(var k=0;k<8;k++)c=c&1?0xEDB88320^(c>>>1):c>>>1;t[n]=c>>>0}return t})();
function crc32(b){var c=0xFFFFFFFF;for(var i=0;i<b.length;i++)c=CRC[(c^b[i])&255]^(c>>>8);return (c^0xFFFFFFFF)>>>0}

function zip(files){
  var parts=[],cd=[],off=0;
  files.forEach(function(f){
    var nm=enc.encode(f.name),d=enc.encode(f.data),crc=crc32(d);
    var h=new DataView(new ArrayBuffer(30));
    h.setUint32(0,0x04034b50,true);h.setUint16(4,20,true);h.setUint16(6,0x0800,true);h.setUint16(12,0x21,true);
    h.setUint32(14,crc,true);h.setUint32(18,d.length,true);h.setUint32(22,d.length,true);h.setUint16(26,nm.length,true);
    parts.push(new Uint8Array(h.buffer),nm,d);
    var c=new DataView(new ArrayBuffer(46));
    c.setUint32(0,0x02014b50,true);c.setUint16(4,20,true);c.setUint16(6,20,true);c.setUint16(8,0x0800,true);c.setUint16(14,0x21,true);
    c.setUint32(16,crc,true);c.setUint32(20,d.length,true);c.setUint32(24,d.length,true);c.setUint16(28,nm.length,true);c.setUint32(42,off,true);
    cd.push(new Uint8Array(c.buffer),nm);
    off+=30+nm.length+d.length;
  });
  var size=cd.reduce(function(a,b){return a+b.length},0);
  var e=new DataView(new ArrayBuffer(22));
  e.setUint32(0,0x06054b50,true);e.setUint16(8,files.length,true);e.setUint16(10,files.length,true);e.setUint32(12,size,true);e.setUint32(16,off,true);
  return new Blob(parts.concat(cd,[new Uint8Array(e.buffer)]),{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}

function xesc(s){return String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,'').replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]})}
function col(i){var s='';i++;while(i>0){var m=(i-1)%26;s=String.fromCharCode(65+m)+s;i=Math.floor((i-1)/26)}return s}
function serial(iso){var p=iso.split('-').map(Number);return (Date.UTC(p[0],p[1]-1,p[2])-Date.UTC(1899,11,30))/864e5}

function sheet(rows){
  var x='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'+
    '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'+
    '<sheetFormatPr defaultRowHeight="15" baseColWidth="16"/><sheetData>';
  rows.forEach(function(r,ri){
    x+='<row r="'+(ri+1)+'">';
    r.forEach(function(v,ci){
      if(v===null||v===undefined||v==='')return;
      var ref=col(ci)+(ri+1);
      if(typeof v==='number'&&isFinite(v))x+='<c r="'+ref+'"><v>'+v+'</v></c>';
      else if(v&&v.d&&/^\d{4}-\d{2}-\d{2}$/.test(v.d))x+='<c r="'+ref+'" s="2"><v>'+serial(v.d)+'</v></c>';
      else x+='<c r="'+ref+'" t="inlineStr"'+(ri===0?' s="1"':'')+'><is><t xml:space="preserve">'+xesc(String(v).slice(0,32000))+'</t></is></c>';
    });
    x+='</row>';
  });
  return x+'</sheetData></worksheet>';
}

var NS='http://schemas.openxmlformats.org/',REL=NS+'officeDocument/2006/relationships/';
function build(sheets){
  var files=[],ct='',wb='',wr='';
  sheets.forEach(function(s,i){
    var n=i+1;
    files.push({name:'xl/worksheets/sheet'+n+'.xml',data:sheet(s.rows)});
    ct+='<Override PartName="/xl/worksheets/sheet'+n+'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
    wb+='<sheet name="'+xesc(s.name.replace(/[\\\/?*\[\]:]/g,' ').slice(0,31))+'" sheetId="'+n+'" r:id="rId'+n+'"/>';
    wr+='<Relationship Id="rId'+n+'" Type="'+REL+'worksheet" Target="worksheets/sheet'+n+'.xml"/>';
  });
  var head='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  files.push(
    {name:'[Content_Types].xml',data:head+'<Types xmlns="'+NS+'package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'+ct+'</Types>'},
    {name:'_rels/.rels',data:head+'<Relationships xmlns="'+NS+'package/2006/relationships"><Relationship Id="rId1" Type="'+REL+'officeDocument" Target="xl/workbook.xml"/></Relationships>'},
    {name:'xl/workbook.xml',data:head+'<workbook xmlns="'+NS+'spreadsheetml/2006/main" xmlns:r="'+NS+'officeDocument/2006/relationships"><sheets>'+wb+'</sheets></workbook>'},
    {name:'xl/_rels/workbook.xml.rels',data:head+'<Relationships xmlns="'+NS+'package/2006/relationships">'+wr+'<Relationship Id="rId'+(sheets.length+1)+'" Type="'+REL+'styles" Target="styles.xml"/></Relationships>'},
    {name:'xl/styles.xml',data:head+'<styleSheet xmlns="'+NS+'spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'}
  );
  return zip(files);
}
return {build:build};
})();
