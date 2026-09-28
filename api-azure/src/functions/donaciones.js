import { app } from '@azure/functions';
import * as XLSX from 'xlsx';

const EXPECTED_HEADERS = [
  'Nro.', 'Fondo / Programa / Proyecto', 'Beneficiario',
  'Descripción de bienes a donar', 'Valor estimado (USD)', 'Fecha de intención',
  'Respuesta del beneficiario', 'Fecha de acta entrega-recepción', 'Estado actual',
  'Avance %', 'Próxima acción requerida', 'Responsable / área', 'Expediente digital',
  'Detalle del trámite', 'Observaciones / alertas', 'Días en gestión', 'Semáforo'
];

const FIELD_ALIASES = {
  nro:['nro','nro.','numero','número','id'],
  programa:['fondo / programa / proyecto','fondo programa proyecto','programa','proyecto','fondo'],
  beneficiario:['beneficiario','beneficiario final'],
  descripcion:['descripción de bienes a donar','descripcion de bienes a donar','descripcion','descripción','bienes'],
  valor:['valor estimado (usd)','valor estimado usd','valor','monto','monto usd'],
  fechaIntencion:['fecha de intención','fecha de intencion','fecha intención','fecha intencion'],
  respuesta:['respuesta del beneficiario','respuesta','fecha respuesta','respuesta beneficiario'],
  fechaActa:['fecha de acta entrega-recepción','fecha de acta entrega recepcion','fecha acta','fecha de acta'],
  estado:['estado actual','estado','situacion','situación'],
  avance:['avance %','avance','porcentaje de avance','avance porcentaje'],
  proximaAccion:['próxima acción requerida','proxima accion requerida','próxima acción','proxima accion','accion requerida','acción requerida'],
  responsable:['responsable / área','responsable area','responsable / area','responsable','área','area'],
  expediente:['expediente digital','expediente','link','enlace','url'],
  detalle:['detalle del trámite','detalle del tramite','detalle','trámite','tramite'],
  observaciones:['observaciones / alertas','observaciones alertas','observaciones','alertas'],
  diasGestion:['días en gestión','dias en gestion','días','dias'],
  semaforo:['semáforo','semaforo','alerta','semaforo alerta']
};

function normalizeText(value){
  return String(value ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim().toLowerCase();
}
function normalizeKey(value){
  return normalizeText(value).replace(/[%()$]/g,'').replace(/[^a-z0-9]+/g,'');
}
const NORMALIZED_ALIASES = Object.fromEntries(
  Object.entries(FIELD_ALIASES).map(([k,v]) => [k, v.map(normalizeKey)])
);
function valueByAliases(row, field){
  const aliases = NORMALIZED_ALIASES[field] || [];
  for(const [key,value] of Object.entries(row)){
    if(aliases.includes(normalizeKey(key))) return value;
  }
  return '';
}
function parseNumber(value){
  if(value === null || value === undefined || value === '') return 0;
  if(typeof value === 'number') return Number.isFinite(value) ? value : 0;
  let s = String(value).trim().replace(/[^\d,.-]/g,'');
  if(s.includes(',') && s.includes('.')) s = s.replace(/\./g,'').replace(',','.');
  else if(s.includes(',')) s = s.replace(',','.');
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}
function parseProgress(value){
  if(value === null || value === undefined || value === '') return 0;
  if(typeof value === 'number') return value > 1 ? value / 100 : value;
  const raw = String(value).trim();
  const n = parseNumber(raw);
  return raw.includes('%') ? n / 100 : (n > 1 ? n / 100 : n);
}
function excelDateToISO(value){
  if(!value) return '';
  if(value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0,10);
  if(typeof value === 'number'){
    const p = XLSX.SSF.parse_date_code(value);
    if(!p) return '';
    return `${String(p.y).padStart(4,'0')}-${String(p.m).padStart(2,'0')}-${String(p.d).padStart(2,'0')}`;
  }
  const text = String(value).trim();
  let m = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if(m) return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;
  m = text.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  if(m) return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
  return text;
}
function isoToUtcDate(iso){
  if(!iso || !/^\d{4}-\d{2}-\d{2}$/.test(iso)) return null;
  const d = new Date(`${iso}T00:00:00Z`);
  return Number.isNaN(d.getTime()) ? null : d;
}
function daysBetween(startIso, endIso){
  const a = isoToUtcDate(startIso), b = isoToUtcDate(endIso);
  if(!a || !b) return null;
  return Math.max(0, Math.floor((b-a)/86400000));
}
function deriveManagement(fechaIntencion, fechaActa, estado){
  const finalizado = normalizeText(estado) === 'finalizado';
  let dias = null;
  if(finalizado){
    if(fechaIntencion && fechaActa) dias = daysBetween(fechaIntencion, fechaActa);
    else if(fechaActa) dias = 0;
  } else {
    const inicio = fechaIntencion || fechaActa;
    if(inicio) dias = daysBetween(inicio, new Date().toISOString().slice(0,10));
  }
  let semaforo = 'Sin iniciar';
  if(finalizado) semaforo = 'Finalizado';
  else if(dias !== null) semaforo = dias > 90 ? 'Crítico' : (dias > 45 ? 'Atención' : 'Normal');
  return {dias, semaforo};
}
function normalizeImported(row){
  const fechaIntencion = excelDateToISO(valueByAliases(row,'fechaIntencion'));
  const fechaActa = excelDateToISO(valueByAliases(row,'fechaActa'));
  const estado = String(valueByAliases(row,'estado') ?? '').trim();
  const derived = deriveManagement(fechaIntencion, fechaActa, estado);
  return {
    nro:valueByAliases(row,'nro'),
    programa:String(valueByAliases(row,'programa') ?? '').trim(),
    beneficiario:String(valueByAliases(row,'beneficiario') ?? '').trim(),
    descripcion:String(valueByAliases(row,'descripcion') ?? '').trim(),
    valor:parseNumber(valueByAliases(row,'valor')),
    fechaIntencion,
    respuesta:String(valueByAliases(row,'respuesta') ?? '').trim(),
    fechaActa,
    estado,
    avance:parseProgress(valueByAliases(row,'avance')),
    proximaAccion:String(valueByAliases(row,'proximaAccion') ?? '').trim(),
    responsable:String(valueByAliases(row,'responsable') ?? '').trim(),
    expediente:String(valueByAliases(row,'expediente') ?? '').trim(),
    detalle:String(valueByAliases(row,'detalle') ?? '').trim(),
    observaciones:String(valueByAliases(row,'observaciones') ?? '').trim(),
    diasGestion:derived.dias,
    semaforo:derived.semaforo
  };
}
function findHeaderRow(matrix){
  return matrix.findIndex(row => {
    if(!Array.isArray(row)) return false;
    const cells = row.map(normalizeKey);
    const hasNro = cells.some(c => ['nro','no','numero','id'].includes(c));
    const hasCore = cells.some(c => c.includes('programa') || c.includes('proyecto') || c.includes('fondo') || c.includes('beneficiario'));
    return hasNro && hasCore;
  });
}
function parseWorkbook(buffer){
  const workbook = XLSX.read(buffer,{type:'buffer',cellDates:true});
  const candidates = ['Control', ...workbook.SheetNames.filter(n => n !== 'Control')];
  let selected = null;
  for(const name of candidates){
    const sheet = workbook.Sheets[name];
    if(!sheet) continue;
    const matrix = XLSX.utils.sheet_to_json(sheet,{header:1,raw:true,defval:null});
    const headerIndex = findHeaderRow(matrix);
    if(headerIndex >= 0){ selected = {name,matrix,headerIndex}; break; }
  }
  if(!selected){
    throw new Error(`No se encontró la tabla de control. Encabezados esperados: ${EXPECTED_HEADERS.join(' | ')}`);
  }
  const headers = selected.matrix[selected.headerIndex].map(h => String(h ?? '').trim());
  const rows = selected.matrix.slice(selected.headerIndex+1)
    .filter(row => Array.isArray(row) && row.some(v => v !== null && v !== undefined && String(v).trim() !== ''))
    .filter(row => normalizeText(row[0]) !== 'total general')
    .map(row => {
      const obj = {};
      headers.forEach((h,i) => { if(h) obj[h] = row[i]; });
      return normalizeImported(obj);
    })
    .filter(r => r.nro !== null && r.nro !== undefined && String(r.nro).trim() !== '');
  if(!rows.length) throw new Error('El Excel no contiene registros válidos.');
  const ids = rows.map(r => String(r.nro).trim());
  const duplicates = [...new Set(ids.filter((id,i) => ids.indexOf(id) !== i))];
  if(duplicates.length) throw new Error(`Existen Nro. duplicados: ${duplicates.join(', ')}`);
  const warnings = [];
  rows.forEach(r => {
    if(!r.programa) warnings.push(`Nro. ${r.nro}: sin Fondo / Programa / Proyecto.`);
    if(!r.beneficiario) warnings.push(`Nro. ${r.nro}: sin Beneficiario.`);
  });
  return {rows, sheetName:selected.name, headerRowIndex:selected.headerIndex, headers, warnings};
}
function isXlsx(buffer){ return buffer.length > 4 && buffer[0] === 0x50 && buffer[1] === 0x4b; }
function isHtml(buffer, contentType=''){
  const sample = buffer.subarray(0,700).toString('utf8').trim().toLowerCase();
  return contentType.toLowerCase().includes('text/html') || sample.startsWith('<!doctype html') || sample.startsWith('<html') || sample.includes('<html');
}
async function getGraphToken(){
  const tenant = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const secret = process.env.CLIENT_SECRET;
  if(!tenant || !clientId || !secret) throw new Error('Faltan TENANT_ID, CLIENT_ID o CLIENT_SECRET.');
  const body = new URLSearchParams({
    client_id:clientId,
    client_secret:secret,
    scope:'https://graph.microsoft.com/.default',
    grant_type:'client_credentials'
  });
  const r = await fetch(`https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/token`,{
    method:'POST', headers:{'Content-Type':'application/x-www-form-urlencoded'}, body
  });
  const p = await r.json();
  if(!r.ok || !p.access_token) throw new Error(p.error_description || 'No se pudo obtener token de Microsoft Graph.');
  return p.access_token;
}
async function downloadFromGraph(){
  const driveId = process.env.GRAPH_DRIVE_ID;
  const itemId = process.env.GRAPH_ITEM_ID;
  if(!driveId || !itemId) throw new Error('Faltan GRAPH_DRIVE_ID o GRAPH_ITEM_ID.');
  const token = await getGraphToken();
  const base = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}`;
  const auth = {Authorization:`Bearer ${token}`};
  let meta = {};
  const mr = await fetch(`${base}?$select=id,name,eTag,lastModifiedDateTime,size`,{headers:auth,cache:'no-store'});
  if(mr.ok) meta = await mr.json();
  const r = await fetch(`${base}/content`,{headers:{...auth,Accept:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'},redirect:'follow',cache:'no-store'});
  const buffer = Buffer.from(await r.arrayBuffer());
  const ct = r.headers.get('content-type') || '';
  if(!r.ok || !isXlsx(buffer) || isHtml(buffer,ct)) throw new Error(`Microsoft Graph no devolvió un XLSX válido (HTTP ${r.status}).`);
  return {buffer, sourceName:meta.name ? `SharePoint · ${meta.name}` : 'SharePoint · Excel origen', sourceModifiedAt:meta.lastModifiedDateTime || null, etag:meta.eTag || null, bytes:buffer.length, sourceMode:'graph'};
}
async function downloadFromPublicShare(){
  const configured = process.env.SHAREPOINT_DOWNLOAD_URL;
  if(!configured) throw new Error('No se configuró SHAREPOINT_DOWNLOAD_URL.');
  const sep = configured.includes('?') ? '&' : '?';
  const url = `${configured}${sep}_=${Date.now()}`;
  const r = await fetch(url,{
    method:'GET', redirect:'follow', cache:'no-store',
    headers:{'User-Agent':'FIAS-Donaciones-LiveAPI','Accept':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*'}
  });
  const buffer = Buffer.from(await r.arrayBuffer());
  const ct = r.headers.get('content-type') || '';
  if(!r.ok || !isXlsx(buffer) || isHtml(buffer,ct)) throw new Error(`SharePoint no devolvió un XLSX válido (HTTP ${r.status}).`);
  return {buffer, sourceName:'SharePoint · Excel institucional', sourceModifiedAt:r.headers.get('last-modified') || null, etag:r.headers.get('etag') || null, bytes:buffer.length, sourceMode:'public-share'};
}
async function downloadSource(){
  const graphReady = process.env.TENANT_ID && process.env.CLIENT_ID && process.env.CLIENT_SECRET && process.env.GRAPH_DRIVE_ID && process.env.GRAPH_ITEM_ID;
  return graphReady ? downloadFromGraph() : downloadFromPublicShare();
}
function corsHeaders(request){
  const origin = request.headers.get('origin') || '';
  const configured = String(process.env.ALLOWED_ORIGINS || '*').split(',').map(x => x.trim()).filter(Boolean);
  const allow = configured.includes('*') ? '*' : (configured.includes(origin) ? origin : configured[0] || 'null');
  return {
    'Access-Control-Allow-Origin':allow,
    'Access-Control-Allow-Methods':'GET,OPTIONS',
    'Access-Control-Allow-Headers':'Content-Type,Accept,Cache-Control',
    'Vary':'Origin',
    'Cache-Control':'no-store, no-cache, must-revalidate, max-age=0',
    'Pragma':'no-cache',
    'Content-Type':'application/json; charset=utf-8'
  };
}

app.http('donaciones',{
  methods:['GET','OPTIONS'],
  authLevel:'anonymous',
  route:'donaciones',
  handler:async (request,context) => {
    const headers = corsHeaders(request);
    if(request.method === 'OPTIONS') return {status:204,headers};
    try{
      const source = await downloadSource();
      const parsed = parseWorkbook(source.buffer);
      return {
        status:200,
        headers,
        jsonBody:{
          data:parsed.rows,
          meta:{
            fetchedAt:new Date().toISOString(),
            sourceName:source.sourceName,
            sourceModifiedAt:source.sourceModifiedAt,
            sourceMode:source.sourceMode,
            etag:source.etag,
            bytes:source.bytes,
            records:parsed.rows.length,
            sheetName:parsed.sheetName,
            headerRowIndex:parsed.headerRowIndex,
            warnings:parsed.warnings
          }
        }
      };
    }catch(error){
      context.error('Error consultando Excel origen',error);
      return {status:502,headers,jsonBody:{error:error?.message || 'No se pudo consultar el Excel origen.'}};
    }
  }
});
