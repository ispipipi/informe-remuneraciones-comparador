const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const fixture = JSON.parse(fs.readFileSync(0, 'utf8'));
const seed = {
  metadata: {locked_group:'Grupo Avanza',show_audit:false},
  months:[{id:'2026-10',label:'Oct 2026'},{id:'2026-11',label:'Nov 2026'}],
  group_options:[{id:'Grupo Avanza',label:'Grupo Avanza'}],
  details:{},concept_options:[],
};
const encode_col = n => {
  let text = '';
  for(n++;n;n=Math.floor((n-1)/26))text=String.fromCharCode(65+(n-1)%26)+text;
  return text;
};
const XLSX = {utils:{
  encode_col,
  encode_cell:({r,c})=>encode_col(c)+(r+1),
  aoa_to_sheet:aoa=>{
    const sheet={aoa};
    aoa.forEach((row,r)=>row.forEach((v,c)=>{if(v!==null&&v!==undefined)sheet[encode_col(c)+(r+1)]={v};}));
    return sheet;
  },
  book_new:()=>({sheets:{}}),
  book_append_sheet:(book,sheet,name)=>{book.sheets[name]=sheet;},
}};
const exported=[];
const context = vm.createContext({assert,fixture,XLSX,window:{XLSX},console:process.env.AVANZA_EXPORT_JSON?{log:console.error}:console});
const html = fixture.template.replace('__DATA__',()=>JSON.stringify(seed));
const start=html.indexOf('<script>')+8,end=html.lastIndexOf('</script>');
const code=html.slice(start,end);
new vm.Script(code);
vm.runInContext(code.slice(0,code.indexOf("document.querySelectorAll('.tab').forEach(t=>t.addEventListener")),context);
XLSX.writeFile=(book,name)=>{context.download={book,name};exported.push({book,name});};
vm.runInContext(`
  const data=buildAvanzaUploadData(fixture.aoa,'Libro octubre noviembre.xlsx');
  const textData=buildAvanzaUploadData(fixture.aoa.map(r=>r.map(v=>v==null?'':String(v))),'Texto.xlsx');
  const rutKey=x=>String(x).replace(/[^0-9k]/gi,'').toUpperCase();
  const id=r=>JSON.stringify([r.empresa,rutKey(r.rut)]);
  for(const month of ['2026-10','2026-11']){
    const rows=data.details[month],expected=fixture.pythonRows[month];
    assert.equal(new Set(rows.map(id)).size,rows.length);
    assert.equal(rows.length,expected.length);
    assert.equal(textData.details[month].length,rows.length);
    assert.equal(textData.kpis[month].total_haberes,data.kpis[month].total_haberes);
    const byPerson=new Map(expected.map(r=>[id(r),r]));
    for(const row of rows){
      const other=byPerson.get(id(row));
      for(const key of ['sueldo_base','total_haberes','sueldo_liquido','dias','sede','area']){
        assert.equal(row[key],other[key],month+' '+key+' '+row.rut);
      }
      assert.equal(JSON.stringify(row.raw.slice(0,25)),JSON.stringify(other.raw.slice(0,25)));
      assert.equal(normalizeAvanzaPeriod(row.raw[25]),normalizeAvanzaPeriod(other.raw[25]));
      assert.equal(row.raw[26],other.raw[26]);
    }
    assert.equal(rows.reduce((s,r)=>s+r.total_haberes,0),fixture.totals[month].haberes);
    assert.equal(rows.reduce((s,r)=>s+r.sueldo_liquido,0),fixture.totals[month].liquido);
    fixture.aoa[5].forEach((header,col)=>{
      const key=headerKey(header);
      if(!/^(haberes |descuentos |aportes patronales |informativos - costo |liquidacion - total )/.test(key)&&key!=='liquidacion - sueldo liquido')return;
      const expectedTotal=fixture.aoa.slice(6).filter(r=>normalizeAvanzaPeriod(r[25])===month).reduce((s,r)=>s+parseNumber(r[col]),0);
      assert.equal(rows.reduce((s,r)=>s+parseNumber(r.raw[col]),0),expectedTotal,month+' '+header);
    });
    const sousa=rows.find(r=>rutKey(r.rut)==='269726253');
    assert.equal(sousa.dias,30);
    assert.equal(sousa.total_haberes,month==='2026-10'?2579115:2779115);
    assert.equal(sousa.sede,'OUTSOURCING');
  }
  DATA=data;byId=Object.fromEntries(data.months.map(m=>[m.id,m]));
  state.grupo='Grupo Avanza';
  for(const report of ['avanza_general','avanza_outsourcing']){
    downloadRevision(report);
    const sheets=download.book.sheets;
    assert.equal(Object.keys(sheets).length,6);
    const allDetails=sheets['Detalle individual'].aoa.slice(2);
    const sousa=allDetails.filter(r=>rutKey(r[3])==='269726253');
    assert.equal(sousa.length,1);
    assert.deepEqual(sousa[0].slice(6,9),[30,30,0]);
    assert.deepEqual(sousa[0].slice(12,15),[2579115,2779115,200000]);
    for(const [month,name] of [['2026-10','Oct 2026'],['2026-11','Nov 2026']]){
      const expected=data.details[month].filter(r=>report==='avanza_general'||isOutsourcingDownloadRow(r));
      const monthRows=sheets[name].aoa;
      assert.equal(monthRows[0].length,fixture.aoa[5].length);
      assert.equal(monthRows.length,expected.length+1);
      const keys=monthRows.slice(1).map(r=>JSON.stringify([r[26],rutKey(r[0])]));
      assert.equal(new Set(keys).size,expected.length);
      const header=headerIndex(monthRows[0]);
      for(const metric of ['Liquidación - Total Haberes','Liquidación - Sueldo Líquido']){
        const col=idxOf(header,metric),total=monthRows.slice(1).reduce((s,r)=>s+parseNumber(r[col]),0);
        const rawTotal=expected.reduce((s,r)=>s+(metric.includes('Total')?r.total_haberes:r.sueldo_liquido),0);
        assert.equal(total,rawTotal);
      }
      assert.equal(sheets['Resumen por empresa'].aoa.at(-1)[month==='2026-10'?2:3],expected.length);
    }
    assert.equal(sheets['Resumen por empresa'].aoa.at(-1)[11],allDetails.reduce((s,r)=>s+r[12],0));
    assert.equal(sheets['Resumen por empresa'].aoa.at(-1)[12],allDetails.reduce((s,r)=>s+r[13],0));
  }
  const original=data.details['2026-11'].find(r=>rutKey(r.rut)==='269726253');
  const crossCompany={...original,empresa:'EMPRESA DE SERVICIOS TRANSITORIOS G.A. SPA'};
  const comparison=detailSheetAoA([original,crossCompany],[original,crossCompany],'2026-10','2026-11');
  assert.equal(comparison.length,4);
  const source=fixture.aoa.slice(6).find(r=>rutKey(r[0])==='269726253');
  const extra=[...source];
  extra[0]='26.972.625-3';extra[2]='Goldfield';extra[3]=30;extra[4]=0;extra[8]=0;extra[12]=123;
  const principal=[...source];principal[0]='26972625-3';principal[2]='Goldfield';
  const merged=consolidateAvanzaAccumulated(fixture.aoa[5],[extra,principal]);
  assert.equal(merged.length,1);
  assert.equal(merged[0][3],30);
  assert.equal(merged[0][12],parseNumber(principal[12])+123);
  console.log('Passed: consolidation, Python/JS parity, totals, days, both exports, unique people, cross-company identity.');
`,context);
if(process.env.AVANZA_EXPORT_JSON)process.stdout.write(JSON.stringify(exported));
