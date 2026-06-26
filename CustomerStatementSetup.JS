// ════════════════════════════════════════════════════════
// CustomerStatementSetup.JS  v2
// SharePoint-hosted — NO inline onclick.
// All events wired via wireEvents() + DOMContentLoaded.
// Mirrors CashFlow.JS v4 patterns exactly.
// v2 fixes: absolute SP URL, non-blocking aging load,
//           page digest fallback, robust error messages.
// ════════════════════════════════════════════════════════

// ── CONFIG ───────────────────────────────────────────────
var SP_BASE   = '/sites/ARCollection';
var SP_FOLDER = 'Shared Documents/Collection';
var SETUP_FILE = 'SetupCustomerSchedule.xlsx';
var AGING_FILE = 'Aging With Customer Notes.xlsx';

// Build absolute origin prefix once (works inside SP modern page iframes)
function spOrigin() {
  // Use window.location.origin; fall back to protocol+host for IE
  return (window.location.origin || (window.location.protocol + '//' + window.location.host));
}

// ── STATE ────────────────────────────────────────────────
var CSS_ST = {
  customers:    [],
  agingRows:    [],
  selectedCols: [],
  toTags:       [],
  ccTags:       [],
  // Multi-value filter tags per field (arrays of strings)
  fTags: { endcust:[], bu:[], material:[], cocode:[], type:[] },
  editingId:    null,
  filterSearch: '',
  colFilters:   {},
  custPage:     0,
  prevPage:     0,
  prevSearch:   '',
  isAdmin:      false,
  agingLoaded:  false
};

// Field → filter tag state key mapping
var FILTER_FIELDS = ['endcust','bu','material','cocode','type'];

var ALL_COLS = [
  'Company Name','SO Number','Customer Name','Cust No.','Customer PO','Pull Number',
  'Doc','Invoice','Inv. List','Delivery','Inv Date','Due Date','Terms',
  'Current','Over Due  1','Over Due 31','Over Due 61','Over Due 91','Over Due 121',
  'Inv Balance','Comments','Header Text','Business Unit','END Customer',
  'Customer Material','TYPE','Delivery Num'
];
var DEFAULT_COLS = [
  'Company Name','SO Number','Customer Name','Cust No.','Customer PO','Pull Number',
  'Doc','Invoice','Inv. List','Delivery','Inv Date','Due Date','Terms',
  'Current','Over Due  1','Over Due 31','Over Due 61','Over Due 91','Over Due 121',
  'Inv Balance','Comments','Header Text','Business Unit','END Customer',
  'Customer Material','TYPE','Delivery Num'
];
var COL_KEY_MAP = {
  'Company Name':'CompanyName','SO Number':'SONumber','Customer Name':'CustomerName',
  'Cust No.':'CustNo','Customer PO':'CustomerPO','Pull Number':'PullNumber',
  'Doc':'Doc','Invoice':'Invoice','Inv. List':'InvList','Delivery':'Delivery',
  'Inv Date':'InvDate','Due Date':'DueDate','Terms':'Terms',
  'Current':'Current','Over Due  1':'OverDue1','Over Due 31':'OverDue31',
  'Over Due 61':'OverDue61','Over Due 91':'OverDue91','Over Due 121':'OverDue121',
  'Inv Balance':'InvBalance','Comments':'Comments','Header Text':'HeaderText',
  'Business Unit':'BusinessUnit','END Customer':'ENDCustomer',
  'Customer Material':'CustomerMaterial','TYPE':'TYPE','Delivery Num':'DeliveryNum'
};

// ── DOM HELPERS ──────────────────────────────────────────
function G(id)       { return document.getElementById(id); }
function txt(id, v)  { var el=G(id); if(el) el.textContent=v; }
function show(id)    { var el=G(id); if(el) el.style.display=''; }
function hide(id)    { var el=G(id); if(el) el.style.display='none'; }

function showErr(msg) {
  var b=G('err-bar'); if(!b) return;
  b.style.display='flex'; G('err-msg').textContent=msg; hide('ok-bar');
}
function showOk(msg) {
  var b=G('ok-bar'); if(!b) return;
  b.style.display='flex'; G('ok-msg').textContent=msg; hide('err-bar');
  setTimeout(function(){ b.style.display='none'; }, 5000);
}
function hideAlerts() { hide('err-bar'); hide('ok-bar'); }

function stamp() {
  txt('s-time','Updated '+new Date().toLocaleTimeString());
}
function stampSetupFileTime(isoDate) {
  var el = G('setup-file-time'); if(!el) return;
  if(!isoDate) { el.textContent=''; return; }
  var d = new Date(isoDate);
  el.textContent = '📄 Setup created: '+d.toLocaleDateString()+' '+d.toLocaleTimeString();
}

function stampAgingFileTime(rawDate) {
  var el = G('aging-file-time'); if(!el) return;
  if(!rawDate) { el.textContent=''; return; }
  // Accept both Date objects and ISO date strings
  var d = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
  if(isNaN(d.getTime())) { el.textContent=''; return; }
  el._createdDate = d.toISOString();
  // SP returns UTC; display in Israel timezone (Asia/Jerusalem)
  var opts = {timeZone:'Asia/Jerusalem',year:'numeric',month:'numeric',day:'numeric',hour:'2-digit',minute:'2-digit'};
  el.textContent = '| 📊 Aging created: '+d.toLocaleString('en-IL', opts);
  // Auto-set statement date to aging file creation date
  var prevDate = G('prev-date');
  if(prevDate) {
    // Use Israel timezone for the date
    var ilDate = new Date(d.toLocaleString('en-US', {timeZone:'Asia/Jerusalem'}));
    prevDate.value = ilDate.getFullYear()+'-'
      +String(ilDate.getMonth()+1).padStart(2,'0')+'-'
      +String(ilDate.getDate()).padStart(2,'0');
  }
}

function fetchAgingFileTime() {
  // ListItemAllFields 'Created' = "Date created" shown in SP properties panel
  var url = spOrigin()+SP_BASE
    +"/_api/web/GetFileByServerRelativePath(decodedurl='"
    +SP_BASE+'/'+SP_FOLDER+'/'+AGING_FILE
    +"')/ListItemAllFields?$select=Created";
  fetch(url, {
    headers: {'Accept':'application/json;odata=verbose'},
    credentials: 'include'
  }).then(function(r){ return r.json(); })
  .then(function(data){
    var created = data && data.d && data.d.Created;
    if(created) stampAgingFileTime(created);
  }).catch(function(e){
    console.warn('[CSS] Could not fetch aging Created date:', e.message);
  });
}
function parseNum(v) {
  if(v===null||v===undefined||v==='') return 0;
  if(typeof v==='number') return v;
  var n=parseFloat(String(v).replace(/[,$\s]/g,''));
  return isNaN(n)?0:n;
}
function fmtUSD(v) { return '$'+Math.round(Number(v)||0).toLocaleString('en-US'); }

// Robust row value getter — tries the exact column name, COL_KEY_MAP alias,
// and trimmed variants to handle any whitespace differences in SP-served Excel files
function getRowVal(r, colName, aliasKey) {
  // Try exact match first
  if(r[colName] !== undefined && r[colName] !== null) return r[colName];
  // Try trimmed key (handles trailing spaces in header cells)
  var trimKey = colName.trim();
  if(trimKey !== colName && r[trimKey] !== undefined && r[trimKey] !== null) return r[trimKey];
  // Try COL_KEY_MAP alias (e.g. 'InvBalance')
  if(aliasKey && r[aliasKey] !== undefined && r[aliasKey] !== null) return r[aliasKey];
  // Numeric columns: if value exists but is 0, return 0 (don't miss falsy zeros)
  if(r[colName] === 0) return 0;
  if(aliasKey && r[aliasKey] === 0) return 0;
  // Scan all keys case-insensitively as last resort
  var lc = colName.toLowerCase().replace(/\s+/g,' ').trim();
  var keys = Object.keys(r);
  for(var i=0;i<keys.length;i++){
    if(keys[i].toLowerCase().replace(/\s+/g,' ').trim() === lc) {
      return r[keys[i]]; // may be 0, null, etc
    }
  }
  return '';
}
function genId()   { return Date.now()+Math.floor(Math.random()*9999); }

// ── SHAREPOINT FETCH ─────────────────────────────────────
// Uses absolute URL (origin + relative path) so it works inside
// SharePoint modern page script-editor iframes where relative
// paths resolve against the CDN, not the SP tenant.
function spFetch(filename) {
  var path = SP_BASE+'/'+SP_FOLDER+'/'+filename;
  var url  = spOrigin()+SP_BASE+"/_api/web/GetFileByServerRelativePath(decodedurl='"+path+"')/$value";
  return fetch(url, {
    headers: { 'Accept': 'application/octet-stream' },
    credentials: 'include'
  }).then(function(res) {
    if(!res.ok) throw new Error('HTTP '+res.status+' for "'+filename+'" — check the file path and SharePoint permissions.');
    return res.arrayBuffer();
  });
}

// ── SHAREPOINT SAVE ──────────────────────────────────────
// 1. Try the page's hidden __REQUESTDIGEST field (fastest — no extra round-trip).
// 2. Fall back to POST /_api/contextinfo if the field is missing or stale.
function getDigest() {
  // SP injects this into classic pages; modern pages may not have it — that's OK,
  // we fall back to contextinfo which always works.
  var el = document.getElementById('__REQUESTDIGEST');
  if (el && el.value && el.value.indexOf('0x') === -1) {
    return Promise.resolve(el.value);
  }
  return fetch(spOrigin()+SP_BASE+'/_api/contextinfo', {
    method: 'POST',
    headers: { 'Accept': 'application/json;odata=verbose', 'Content-Length': '0' },
    credentials: 'include'
  }).then(function(res) {
    if(!res.ok) throw new Error('contextinfo HTTP '+res.status);
    return res.json();
  }).then(function(info) {
    return info.d.GetContextWebInformation.FormDigestValue;
  });
}

function spSave(filename, arrayBuffer) {
  var path = SP_BASE+'/'+SP_FOLDER+'/'+filename;
  var uploadUrl = spOrigin()+SP_BASE+"/_api/web/GetFileByServerRelativePath(decodedurl='"+path+"')/$value";
  return getDigest().then(function(digest) {
    return fetch(uploadUrl, {
      method: 'PUT',
      headers: {
        'Accept': 'application/json;odata=verbose',
        'X-RequestDigest': digest,
        'Content-Type': 'application/octet-stream'
      },
      credentials: 'include',
      body: arrayBuffer
    });
  }).then(function(res) {
    if(!res.ok) throw new Error('Upload HTTP '+res.status+
      ' — the file may be checked out or you may not have edit permissions.');
    return res;
  });
}

// ── XLSX HELPERS ─────────────────────────────────────────
function xlsxToRows(buf) {
  var wb = XLSX.read(buf, { type:'array', cellDates:true });
  var wsName = wb.SheetNames[0];
  var ws = wb.Sheets[wsName];
  var rows = XLSX.utils.sheet_to_json(ws, { defval:null, raw:true });
  return { rows:rows, wb:wb, wsName:wsName };
}

function findAgingSheet(wb) {
  var prefer = ['AR_AGING','ARAGING','AR AGING'];
  for(var p=0; p<prefer.length; p++) {
    for(var s=0; s<wb.SheetNames.length; s++) {
      if(wb.SheetNames[s].toUpperCase()===prefer[p]) return wb.SheetNames[s];
    }
  }
  return wb.SheetNames[0];
}

// ── NORMALISE SETUP ROW ───────────────────────────────────
function normaliseSetupRow(r, idx) {
  function g(keys) {
    for(var i=0; i<keys.length; i++) {
      if(r[keys[i]]!==undefined&&r[keys[i]]!==null&&r[keys[i]]!=='') return r[keys[i]];
    }
    return '';
  }
  var toRaw  = String(g(['Email','To Email','ToEmail'])||'').trim();
  var toArr  = toRaw ? toRaw.split(/[;,]+/).map(function(e){return e.trim();}).filter(Boolean) : [];
  var ccRaw  = String(g(['Collector Email','CollectorEmail','CC'])||'').trim();
  var ccArr  = ccRaw ? ccRaw.split(/[;,]+/).map(function(e){return e.trim();}).filter(Boolean) : [];
  var sched  = String(g(['Schedule'])||'None').trim();
  var schedV = parseNum(g(['Schedule Time','ScheduleTime']))||1;
  return {
    id:           genId()+idx,
    custno:       String(g(['Cust No.','CustNo','Customer No.'])||'').replace(/^'+/,'').trim(),
    name:         String(g(['Customer Name','CustomerName'])||'').replace(/^'+/,'').trim(),
    to:           toArr,
    cc:           ccArr,
    emailtype:    String(g(['Email Type','EmailType'])||'Statement').trim(),
    collector:    String(g(['Collector Email','CollectorEmail'])||'').trim(),
    endcust:      String(g(['END Customer','ENDCustomer'])||'').trim(),
    bu:           String(g(['Business Unit','BusinessUnit'])||'').trim(),
    material:     String(g(['Customer Material','CustomerMaterial'])||'').trim(),
    cocode:       String(g(['CoCode'])||'').trim(),
    type:         String(g(['TYPE','Type'])||'').trim(),
    consolidated: String(g(['Consolidated Customer','ConsolidatedCustomer'])||'').trim(),
    arma:         String(g(['Include ARMA','IncludeARMA'])||'Y').trim().toUpperCase()==='Y',
    credit:       String(g(['Include Credit','IncludeCredit'])||'Y').trim().toUpperCase()==='Y',
    promised:     String(g(['Promised Date','PromisedDate'])||'').trim().toUpperCase()==='Y',
    schedule:     sched,
    schedtime:    schedV,
    time:         '08:00',
    group:        String(g(['Group','group'])||'').trim(),
    _src:         r
  };
}

// ── BUILD SETUP XLSX ──────────────────────────────────────
function customersToRows() {
  return CSS_ST.customers.map(function(c) {
    return {
      'Cust No.':             c.custno,
      'Customer Name':        c.name,
      'Email':                c.to.join('; '),
      'END Customer':         c.endcust,
      'Business Unit':        c.bu,
      'Customer Material':    c.material,
      'CoCode':               c.cocode,
      'TYPE':                 c.type,
      'Consolidated Customer':c.consolidated,
      'Collector Email':      c.cc.join('; '),
      'Email Type':           c.emailtype,
      'Promised Date':        c.promised ? 'Y' : '',
      'Include ARMA':         c.arma    ? 'Y' : '',
      'Include Credit':       c.credit  ? 'Y' : '',
      'Group':                c.group||'',
      'Schedule':             c.schedule,
      'Schedule Time':        c.schedtime
    };
  });
}

function buildSetupXlsx() {
  var rows = customersToRows();
  var ws   = XLSX.utils.json_to_sheet(rows);
  var colWidths = Object.keys(rows[0]||{}).map(function(k){ return {wch: Math.max(k.length+4,12)}; });
  ws['!cols'] = colWidths;
  // Add named table 'CustomerSchedule' covering all data (required for Power Automate)
  var range = XLSX.utils.decode_range(ws['!ref']||'A1');
  ws['!autofilter'] = { ref: ws['!ref'] };
  ws['!tables'] = [{
    name: 'CustomerSchedule',
    ref: ws['!ref'],
    headerRow: true,
    totalsRow: false,
    styleInfo: { themeIndex: 2, showFirstColumn: false, showLastColumn: false,
                 showRowStripes: true, showColumnStripes: false }
  }];
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  return XLSX.write(wb, { bookType:'xlsx', type:'array' });
}

// ── LOAD FROM SHAREPOINT ─────────────────────────────────
// Two-phase: load Setup first (fast, small), then Aging in background (large).
function loadFromSharePoint(forceRefresh) {
  hideAlerts();
  var btn = G('btn-reload');
  if(btn){ btn.disabled=true; btn.textContent='Loading setup...'; }

  spFetch(SETUP_FILE).then(function(buf) {
    var res = xlsxToRows(buf);
    CSS_ST.customers = res.rows
      .map(normaliseSetupRow)
      .filter(function(c){ return c.custno||c.name; });

    stamp();
    renderCustomers();
    renderScheduleList();
    populatePreviewDropdown();
    // Fetch file last-modified time
    try {
      var metaUrl = spOrigin()+SP_BASE+"/_api/web/GetFileByServerRelativePath(decodedurl='"+SP_BASE+"/"+SP_FOLDER+"/"+SETUP_FILE+"')/ListItemAllFields?$select=Created";
      fetch(metaUrl,{headers:{'Accept':'application/json;odata=verbose'},credentials:'include'})
        .then(function(r){return r.json();})
        .then(function(d){ stampSetupFileTime(d&&d.d&&d.d.TimeLastModified); })
        .catch(function(){});
    } catch(e){}
    if(btn){ btn.disabled=false; btn.textContent='\u21BA Reload from SharePoint'; }
    showOk('Loaded '+CSS_ST.customers.length+' customer(s) from SharePoint. Loading aging data...');

    // Phase 2: load aging (from cache if available, SharePoint if not)
    loadAgingBackground(forceRefresh||false);

  }).catch(function(err) {
    if(btn){ btn.disabled=false; btn.textContent='\u21BA Reload from SharePoint'; }
    showErr('Could not load '+SETUP_FILE+': '+err.message);
    console.error('[CSS] spFetch setup error:', err);
  });
}

// ── AGING CACHE ──────────────────────────────────────────
// Cache key includes today's date so data auto-refreshes daily.
// Cache stored in sessionStorage (cleared when browser tab closes).
// Force-refresh: hold Shift when clicking "Reload from SharePoint".
var AGING_CACHE_KEY = 'css_aging_' + new Date().toISOString().slice(0,10); // e.g. css_aging_2026-06-17

function loadAgingBackground(forceRefresh) {
  var agingStatus = G('aging-status');

  // 1. Try sessionStorage cache first (instant — no network)
  if(!forceRefresh) {
    try {
      var cached = sessionStorage.getItem(AGING_CACHE_KEY);
      if(cached) {
        var parsed = JSON.parse(cached);
        if(parsed && parsed.rows && parsed.rows.length) {
          CSS_ST.agingRows  = parsed.rows;
          CSS_ST.agingLoaded = true;
          buildAgingLookups();
          if(agingStatus) agingStatus.textContent = CSS_ST.agingRows.length+' aging rows loaded (cached)';
          console.log('[CSS] Aging loaded from cache:', CSS_ST.agingRows.length, 'rows');
          // Restore creation date from cache if stored
          // Restore from cache or fetch async
          if(parsed.createdDate) { stampAgingFileTime(parsed.createdDate); } else { setTimeout(fetchAgingFileTime, 200); }
          return; // done — no network call needed
        }
      }
    } catch(e) {
      console.warn('[CSS] Cache read failed, fetching from SharePoint:', e.message);
    }
  }

  // 2. Cache miss or force-refresh — fetch from SharePoint
  if(agingStatus) agingStatus.textContent = 'Loading aging data from SharePoint...';

  spFetch(AGING_FILE).then(function(buf) {
    var wb    = XLSX.read(buf, { type:'array', cellDates:true });
    var sheet = findAgingSheet(wb);
    CSS_ST.agingRows   = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { defval:null, raw:true });
    // Fetch creation date from SharePoint (called separately)
    fetchAgingFileTime();
    CSS_ST.agingLoaded = true;
    buildAgingLookups();

    // 3. Save to sessionStorage for instant reuse this session
    // sessionStorage limit is ~5MB per origin — JSON of 48K rows is ~15MB so
    // we store only the columns we actually need (reduces to ~3–4MB)
    try {
      var slim = CSS_ST.agingRows.map(function(r) {
        // Keep only columns used for filtering, preview, dropdowns, and statement
        return {
          'Cust No.':         r['Cust No.'],
          'Customer Name':    r['Customer Name'],
          'ARMA':             r['ARMA'],
          'Promised Date':    r['Promised Date'],
          'Inv Balance':      r['Inv Balance'],
          'CoCode':           r['CoCode'],
          'SO Number':        r['SO Number'],
          'Invoice':          r['Invoice'],
          'Customer PO':      r['Customer PO'],
          'Pull Number':      r['Pull Number'],
          'Doc':              r['Doc'],
          'Inv. List':        r['Inv. List'],
          'Delivery':         r['Delivery'],
          'Inv Date':         r['Inv Date'],
          'Due Date':         r['Due Date'],
          'Terms':            r['Terms'],
          'Current':          r['Current'],
          'Over Due  1':      r['Over Due  1'],
          'Over Due 31':      r['Over Due 31'],
          'Over Due 61':      r['Over Due 61'],
          'Over Due 91':      r['Over Due 91'],
          'Over Due 121':     r['Over Due 121'],
          'Comments':         r['Comments'],
          'Header Text':      r['Header Text'],
          'Business Unit':    r['Business Unit'],
          'END Customer':     r['END Customer'],
          'Customer Material':r['Customer Material'],
          'TYPE':             r['TYPE'],
          'Delivery Num':     r['Delivery Num'],
          'Collector':        r['Collector']||r['Collector Name'],
          'Company Name':     r['Company Name']
        };
      });
      var json = JSON.stringify({ rows: slim, ts: Date.now(), createdDate: (G('aging-file-time')&&G('aging-file-time')._createdDate)||null });
      sessionStorage.setItem(AGING_CACHE_KEY, json);
      console.log('[CSS] Aging cached in sessionStorage:', Math.round(json.length/1024)+'KB');
    } catch(e) {
      // sessionStorage quota exceeded — store a smaller subset (just filter columns)
      console.warn('[CSS] Cache failed ('+e.message+'), trying slim-only cache...');
      try {
        Object.keys(sessionStorage).forEach(function(k){
          if(k.indexOf('css_aging_')===0||k.indexOf('msal.')===0) sessionStorage.removeItem(k);
        });
        // Store only rows needed for filter matching — not all display columns
        var micro = CSS_ST.agingRows.map(function(r){
          return {
            'Cust No.':        r['Cust No.'],
            'ARMA':            r['ARMA'],
            'Promised Date':   r['Promised Date'],
            'Inv Balance':     r['Inv Balance'],
            'CoCode':          r['CoCode'],
            'Business Unit':   r['Business Unit'],
            'END Customer':    r['END Customer'],
            'Customer Material':r['Customer Material'],
            'TYPE':            r['TYPE'],
            'Collector':       r['Collector']||r['Collector Name'],
            'Customer Name':   r['Customer Name']
          };
        });
        var j2 = JSON.stringify({ rows: micro, ts: Date.now(), slim: true });
        sessionStorage.setItem(AGING_CACHE_KEY, j2);
        console.log('[CSS] Micro-cache stored:', Math.round(j2.length/1024)+'KB');
      } catch(e2) {
        console.warn('[CSS] Even micro-cache failed — running without cache');
      }
    }

    if(agingStatus) agingStatus.textContent = CSS_ST.agingRows.length+' aging rows loaded';
    console.log('[CSS] Aging loaded from SharePoint:', CSS_ST.agingRows.length, 'rows from sheet', sheet);
    fetchAgingFileTime();

  }).catch(function(err) {
    CSS_ST.agingLoaded = false;
    if(agingStatus) agingStatus.textContent = 'Aging data unavailable — reload to retry';
    console.warn('[CSS] Aging load failed:', err.message);
  });
}

// ── AGING LOOKUP INDEX ───────────────────────────────────
// Built once after aging rows load; used to drive modal dropdowns.
// CSS_ST.agingIndex = { custno: { name, collectors:[], cocodes:[] } }
// CSS_ST.agingCustList = [ {custno, name} ... ] sorted by name
function buildAgingLookups() {
  var idx = {};
  CSS_ST.agingRows.forEach(function(r) {
    var no   = String(r['Cust No.']||r['CustNo']||'').replace(/^'+/,'').trim();
    var nm   = String(r['Customer Name']||r['CustomerName']||'').trim();
    var col  = String(r['Collector']||r['Collector Name']||r['CollectorName']||'').trim();
    var cocd = String(r['CoCode']||'').trim();
    if(!no && !nm) return;
    if(!idx[no]) idx[no] = { custno:no, name:nm, collectors:{}, cocodes:{} };
    if(!idx[no].name && nm) idx[no].name = nm;
    if(col)  idx[no].collectors[col]  = 1;
    if(cocd) idx[no].cocodes[cocd]    = 1;
  });
  // Convert sets → sorted arrays
  CSS_ST.agingIndex = {};
  Object.keys(idx).forEach(function(k) {
    CSS_ST.agingIndex[k] = {
      custno:     idx[k].custno,
      name:       idx[k].name,
      collectors: Object.keys(idx[k].collectors).sort(),
      cocodes:    Object.keys(idx[k].cocodes).sort()
    };
  });
  // Sorted list for dropdown
  CSS_ST.agingCustList = Object.values(CSS_ST.agingIndex)
    .filter(function(c){ return c.name||c.custno; })
    .sort(function(a,b){ return (a.name||'').localeCompare(b.name||''); });
  console.log('[CSS] Aging index built:', CSS_ST.agingCustList.length, 'unique customers');
  // Refresh modal dropdowns if modal is already open
  if(G('modal-overlay').className.indexOf('open')>-1) populateModalDropdowns();
}

// Populate the custno datalist and name datalist from agingCustList
function populateModalDropdowns() {
  var list = CSS_ST.agingCustList || [];
  // custno datalist
  var dlNo = G('dl-custno'); if(dlNo) {
    dlNo.innerHTML = list.map(function(c){
      return '<option value="'+esc(c.custno)+'">'+(c.name?esc(c.name):'')+'</option>';
    }).join('');
  }
  // name datalist
  var dlNm = G('dl-custname'); if(dlNm) {
    dlNm.innerHTML = list.map(function(c){
      return '<option value="'+esc(c.name)+'">'+(c.custno?'('+esc(c.custno)+')':'')+'</option>';
    }).join('');
  }
}

// Called when user picks a cust no → auto-fill name + dropdowns
function onCustNoChange() {
  var inp  = G('m-custno'); if(!inp) return;
  var val  = inp.value.trim();
  var rec  = CSS_ST.agingIndex && CSS_ST.agingIndex[val];
  if(rec) {
    var nm = G('m-custname'); if(nm && !nm.value) nm.value = rec.name;
    populateCollectorSelect(rec.collectors);
    populateCocodeSelect(rec.cocodes);
  }
}

// Called when user picks a name → auto-fill cust no + dropdowns
function onCustNameChange() {
  var inp  = G('m-custname'); if(!inp) return;
  var val  = inp.value.trim();
  var rec  = null;
  var list = CSS_ST.agingCustList || [];
  for(var i=0;i<list.length;i++){
    if(list[i].name.toLowerCase()===val.toLowerCase()){ rec=list[i]; break; }
  }
  if(rec) {
    var no = G('m-custno'); if(no && !no.value) no.value = rec.custno;
    populateCollectorSelect(rec.collectors);
    populateCocodeSelect(rec.cocodes);
  }
}

function populateCollectorSelect(collectors, preselectVal) {
  var sel = G('m-collector'); if(!sel) return;
  var useVal = (preselectVal !== undefined) ? preselectVal : sel.value;
  // Merge in the preselectVal if it's not already in the list
  var opts = collectors.slice();
  if(useVal && opts.indexOf(useVal) === -1) opts.unshift(useVal);
  sel.innerHTML = '<option value="">-- select --</option>'
    + opts.map(function(c){ return '<option value="'+esc(c)+'"'+(c===useVal?' selected':'')+'>'+esc(c)+'</option>'; }).join('');
  sel.value = useVal || '';
}

function populateCocodeSelect(cocodes, preselectVal) {
  // CoCode is now a tag input — add a datalist for suggestions
  var inp = G('m-cocode-inp'); if(!inp) return;
  // Build or update datalist
  var dlId = 'dl-cocode-suggest';
  var dl = G(dlId);
  if(!dl){ dl=document.createElement('datalist'); dl.id=dlId; document.body.appendChild(dl); }
  dl.innerHTML = cocodes.map(function(c){ return '<option value="'+esc(c)+'">'; }).join('');
  inp.setAttribute('list', dlId);
  // If editing and preselectVal provided, seed the tag array
  if(preselectVal) {
    strToFTags(preselectVal).forEach(function(v){
      if(CSS_ST.fTags.cocode.indexOf(v)===-1) CSS_ST.fTags.cocode.push(v);
    });
    renderFilterTags('cocode');
  }
}

// ── RENDER: CUSTOMERS TABLE ──────────────────────────────
function renderCustomers() {
  var q   = CSS_ST.filterSearch.toLowerCase();
  var cf  = CSS_ST.colFilters || {};
  var rows = CSS_ST.customers.filter(function(c) {
    // Global search (all fields)
    if(q) {
      var all = [c.custno,c.name,c.to.join(' '),c.cc.join(' '),c.emailtype,c.schedule,c.group||''].join(' ').toLowerCase();
      if(all.indexOf(q) === -1) return false;
    }
    // Column filters (Enter-triggered, per-column)
    if(cf.custno    && c.custno.toLowerCase().indexOf(cf.custno.toLowerCase())   ===-1) return false;
    if(cf.name      && c.name.toLowerCase().indexOf(cf.name.toLowerCase())       ===-1) return false;
    if(cf.to        && c.to.join(' ').toLowerCase().indexOf(cf.to.toLowerCase())  ===-1) return false;
    if(cf.cc        && c.cc.join(' ').toLowerCase().indexOf(cf.cc.toLowerCase())  ===-1) return false;
    if(cf.emailtype && (c.emailtype||'').toLowerCase().indexOf(cf.emailtype.toLowerCase()) ===-1) return false;
    if(cf.schedule  && c.schedule.toLowerCase().indexOf(cf.schedule.toLowerCase())===-1) return false;
    if(cf.group     && (c.group||'').toLowerCase().indexOf(cf.group.toLowerCase())===-1) return false;
    return true;
  });

  var tbody = G('cust-tbody'); if(!tbody) return;
  txt('cust-count', rows.length+' of '+CSS_ST.customers.length+' customer(s)');

  if(!rows.length) {
    tbody.innerHTML='<tr><td colspan="8" class="empty">No customers match the current filters.</td></tr>';
    wireCustTableEvents();
    return;
  }

  tbody.innerHTML = rows.map(function(c) {
    var schedBdg = c.schedule==='None'?'bdg-gray':c.schedule==='Daily'?'bdg-green':c.schedule==='Weekly'?'bdg-blue':'bdg-amber';
    var toDisp   = c.to.length
      ? (c.to[0]+(c.to.length>1?' <span class="bdg bdg-gray">+'+(c.to.length-1)+'</span>':''))
      : '<span style="color:var(--gray);font-style:italic">none set</span>';
    var opts = [];
    if(c.arma)    opts.push('<span class="bdg bdg-gray">ARMA</span>');
    if(c.credit)  opts.push('<span class="bdg bdg-gray">Credit</span>');
    if(c.promised)opts.push('<span class="bdg bdg-gray">Prom.</span>');

    return '<tr>'
      +'<td><span class="bdg bdg-gray">'+esc(c.custno||'—')+'</span></td>'
      +'<td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="'+esc(c.name)+'">'+esc(c.name)+'</td>'
      +'<td style="font-size:11px;max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">'+toDisp+'</td>'
      +'<td style="font-size:11px">'+esc(c.cc.length?c.cc.join(', '):'—')+'</td>'
      +'<td><span class="bdg bdg-blue">'+esc(c.emailtype||'Statement')+'</span></td>'
      +'<td><span class="bdg '+schedBdg+'">'+esc(c.schedule==='None'?'None':(c.schedule+(c.schedule!=='Daily'?' /'+c.schedtime:'')))+'</span>'+(c.group?' <span class="bdg bdg-gray" style="font-size:10px">'+esc(c.group)+'</span>':'')+'</td>'
      +'<td style="font-size:11px">'+(opts.length?opts.join(' '):'—')+'</td>'
      +'<td style="text-align:center;white-space:nowrap">'
        +'<button class="btn btn-outline btn-sm css-edit-btn" data-cid="'+c.id+'" style="margin-right:4px">Edit</button>'
        +'<button class="btn btn-danger btn-sm css-del-btn" data-cid="'+c.id+'">Del</button>'
      +'</td>'
      +'</tr>';
  }).join('');

  wireCustTableEvents();
}

// Download Customers tab as Excel
function downloadCustomersExcel() {
  var rows = CSS_ST.customers.map(function(c){
    return {
      'Cust No.':         c.custno,
      'Customer Name':    c.name,
      'To Email(s)':      c.to.join('; '),
      'CC':               c.cc.join('; '),
      'Email Type':       c.emailtype,
      'Collector':        c.collector,
      'Group':            c.group||'',
      'END Customer':     c.endcust||'',
      'Business Unit':    c.bu||'',
      'Customer Material':c.material||'',
      'CoCode':           c.cocode||'',
      'TYPE':             c.type||'',
      'Consolidated':     c.consolidated||'',
      'Include ARMA':     c.arma?'Y':'N',
      'Include Credit':   c.credit?'Y':'N',
      'Promised Date':    c.promised?'Y':'N',
      'Schedule':         c.schedule,
      'Schedule Time':    c.schedtime||''
    };
  });
  var ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = Object.keys(rows[0]||{}).map(function(k){return {wch:Math.max(k.length+2,14)};});
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Customers');
  var today = new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb, 'CustomerSetup_'+today+'.xlsx');
}

// Download Schedule tab as Excel
function downloadScheduleExcel() {
  var rows = CSS_ST.customers.map(function(c){
    return {
      'Cust No.':      c.custno,
      'Customer Name': c.name,
      'Schedule':      c.schedule,
      'Day Value':     c.schedtime||'',
      'Group':         c.group||'',
      'Collector':     c.collector||'',
      'To Email(s)':   c.to.join('; ')
    };
  });
  var ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = Object.keys(rows[0]||{}).map(function(k){return {wch:Math.max(k.length+2,14)};});
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Schedule');
  var today = new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb, 'CustomerSchedule_'+today+'.xlsx');
}

// Check if current user is a SharePoint site admin
function checkAdminStatus() {
  var url = spOrigin()+SP_BASE+'/_api/web/currentUser/isSiteAdmin';
  return fetch(url, {
    headers:{'Accept':'application/json;odata=verbose'},
    credentials:'include'
  }).then(function(r){ return r.json(); })
  .then(function(d){
    CSS_ST.isAdmin = !!(d && d.d && d.d.IsSiteAdmin);
    applyAdminUI();
  }).catch(function(){ CSS_ST.isAdmin = false; applyAdminUI(); });
}

function applyAdminUI() {
  var banner   = G('filters-readonly-banner');
  var controls = G('filters-admin-controls');
  if(CSS_ST.isAdmin) {
    if(banner)   banner.style.display='none';
    if(controls) { controls.disabled=false; controls.style.opacity=''; controls.style.pointerEvents=''; }
  } else {
    if(banner)   banner.style.display='';
    if(controls) { controls.disabled=true; controls.style.opacity='0.6'; controls.style.pointerEvents='none'; }
  }
}

function esc(s) {
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function wireCustTableEvents() {
  var tbody=G('cust-tbody'); if(!tbody||tbody._cssWired) return;
  tbody._cssWired=true;
  tbody.addEventListener('click', function(e) {
    var btn=e.target;
    while(btn&&btn!==tbody){
      if(btn.classList&&(btn.classList.contains('css-edit-btn')||btn.classList.contains('css-del-btn'))) break;
      btn=btn.parentNode;
    }
    if(!btn||btn===tbody) return;
    var cid=Number(btn.getAttribute('data-cid'));
    if(btn.classList.contains('css-edit-btn'))     openModal(cid);
    else if(btn.classList.contains('css-del-btn')) deleteCustomer(cid);
  });
}

// ── RENDER: COLUMN GRID ──────────────────────────────────
function renderColGrid() {
  var el=G('col-grid'); if(!el) return;
  el.innerHTML = ALL_COLS.map(function(c) {
    var sel = CSS_ST.selectedCols.indexOf(c)>-1;
    return '<div class="col-item'+(sel?' sel':'')+'" data-col="'+esc(c)+'">'
      +'<span class="chk">'+(sel?'[x]':'[ ]')+'</span> '+esc(c)
      +'</div>';
  }).join('');
  if(!el._colWired) {
    el._colWired=true;
    el.addEventListener('click', function(e) {
      var item=e.target;
      while(item&&item!==el){ if(item.classList&&item.classList.contains('col-item')) break; item=item.parentNode; }
      if(!item||item===el) return;
      var c=item.getAttribute('data-col'); if(!c) return;
      var idx=CSS_ST.selectedCols.indexOf(c);
      if(idx>-1) CSS_ST.selectedCols.splice(idx,1); else CSS_ST.selectedCols.push(c);
      renderColGrid();
    });
  }
}

// ── RENDER: SCHEDULE LIST ────────────────────────────────
function renderScheduleList() {
  var el=G('sched-list'); if(!el) return;
  if(!CSS_ST.customers.length) { el.innerHTML='<div class="empty">No customers loaded yet.</div>'; return; }

  el.innerHTML = CSS_ST.customers.map(function(c,idx) {
    var isNone    = c.schedule==='None';
    var isDaily   = c.schedule==='Daily';
    var dayLabel  = c.schedule==='Weekly'?'Day (1-7)':c.schedule==='Monthly'?'Day of month':'';
    return '<div class="sched-row">'
      +'<div><div style="font-size:13px;font-weight:600">'+esc(c.name)+'</div>'
        +'<div style="font-size:11px;color:var(--gray)">#'+esc(c.custno)+'</div></div>'
      +'<div class="field" style="margin:0"><label>Frequency</label>'
        +'<select class="sched-freq-sel" data-idx="'+idx+'">'
          +['None','Daily','Weekly','Monthly'].map(function(s){
            return '<option value="'+s+'"'+(c.schedule===s?' selected':'')+'>'+s+'</option>';
          }).join('')
        +'</select></div>'
      +'<div class="field" style="margin:0">'
        +((!isNone && !isDaily)
          ? '<label>'+dayLabel+'</label>'
            +'<input type="number" class="sched-day-inp" data-idx="'+idx+'"'
              +' min="1" max="'+(c.schedule==='Weekly'?'7':'31')+'" value="'+c.schedtime+'"'
              +' style="padding:7px 10px;font-size:13px;border:1px solid var(--border);border-radius:6px;">'
          : '<label>&nbsp;</label><div style="color:var(--gray);font-size:12px;padding-top:8px">'
              +(isNone?'Manual only':'Every business day')+'</div>')
      +'</div>'
      +'<div style="padding-bottom:2px;align-self:flex-end">'
        +'<button class="btn btn-outline btn-sm css-sched-prev" data-cid="'+c.id+'">Preview</button>'
      +'</div>'
      +'</div>';
  }).join('');

  wireSchedEvents();
}

function wireSchedEvents() {
  var el=G('sched-list'); if(!el||el._schedWired) return;
  el._schedWired=true;
  el.addEventListener('change', function(e) {
    var t=e.target; if(!t.getAttribute) return;
    var idx=Number(t.getAttribute('data-idx')); if(isNaN(idx)) return;
    var c=CSS_ST.customers[idx]; if(!c) return;
    if(t.classList.contains('sched-freq-sel')){ c.schedule=t.value; renderScheduleList(); }
    else if(t.classList.contains('sched-day-inp')){ c.schedtime=parseInt(t.value,10)||1; }
  });
  el.addEventListener('click', function(e) {
    var btn=e.target;
    while(btn&&btn!==el){ if(btn.classList&&btn.classList.contains('css-sched-prev')) break; btn=btn.parentNode; }
    if(!btn||btn===el) return;
    var cid=Number(btn.getAttribute('data-cid'));
    switchTab('preview');
    renderPreviewForCust(cid);
  });
}

// ── RENDER: PREVIEW ──────────────────────────────────────
function populatePreviewDropdown() {
  var sel=G('prev-cust'); if(!sel) return;
  var prevSelected = getSelectedPrevIds();
  sel.innerHTML = CSS_ST.customers.map(function(c){
    var wasSelected = prevSelected.indexOf(c.id) > -1;
    return '<option value="'+c.id+'"'+(wasSelected?' selected':'')+'>'+esc(c.name)+' ('+esc(c.custno)+')</option>';
  }).join('');
  // Ensure checklist selection state is initialised
  if(!CSS_ST._checklistSel) CSS_ST._checklistSel = {};
}

function getSelectedPrevIds() {
  var sel=G('prev-cust'); if(!sel) return [];
  var ids=[];
  for(var i=0;i<sel.options.length;i++){
    if(sel.options[i].selected) ids.push(Number(sel.options[i].value));
  }
  return ids;
}

function getSelectedCustomers() {
  // Primary: checklist checkboxes (Step 1)
  var fromChecklist = getChecklistSelected();
  if(fromChecklist.length) return fromChecklist;
  // Fallback: prev-cust multi-select (legacy)
  var ids = getSelectedPrevIds();
  return CSS_ST.customers.filter(function(c){ return ids.indexOf(c.id)>-1; });
}

// ── PREVIEW: multi-customer ────────────────────────────────────────────────

function renderPreview() {
  var area=G('prev-area'); if(!area) return;
  var custs=getSelectedCustomers();
  if(!custs.length){ area.innerHTML='<div class="empty">Select customers above and click "Preview selected".</div>'; txt('prev-summary',''); return; }
  if(!CSS_ST.agingLoaded){ area.innerHTML='<div style="background:var(--amber-l);border:1px solid var(--amber-b);border-radius:7px;padding:10px 14px;font-size:13px;color:var(--amber)">Aging data still loading — please wait and try again.</div>'; return; }
  var html=[]; var totalRows=0; var totalBal=0;
  custs.forEach(function(cust){
    var filtered=filterAgingForCustomer(cust);
    totalRows+=filtered.length;
    filtered.forEach(function(r){ totalBal+=parseNum(getRowVal(r,'Inv Balance','InvBalance')); });
    html.push(renderCustomerPanel(cust,filtered));
  });
  area.innerHTML=html.join('');
  wirePreviewPanelToggles();
  txt('prev-summary', custs.length+' customer(s) \u00b7 '+totalRows+' row(s) \u00b7 Total: '+fmtUSD(totalBal));
}

// Returns the set of Cust No. values this customer's statement covers.
// If consolidated field is filled (pipe-separated), includes all those numbers.
// Always includes the customer's own custno.
function getConsolidatedCustNos(cust) {
  var nos = {};
  if(cust.custno) nos[cust.custno.trim()] = 1;
  if(cust.consolidated && cust.consolidated.trim()) {
    cust.consolidated.split('|').forEach(function(n){
      var t=n.trim(); if(t) nos[t]=1;
    });
  }
  return Object.keys(nos);
}

function filterAgingForCustomer(cust) {
  var nos = getConsolidatedCustNos(cust);

  // Parse pipe-separated filter values into lowercase arrays for fast lookup
  function mkSet(str){ return (str||'').split('|').map(function(v){return v.trim().toLowerCase();}).filter(Boolean); }
  var fEndCust  = mkSet(cust.endcust);
  var fBU       = mkSet(cust.bu);
  var fMaterial = mkSet(cust.material);
  var fCoCode   = mkSet(cust.cocode);
  var fType     = mkSet(cust.type);

  return CSS_ST.agingRows.filter(function(r){
    // ── Cust No. match (consolidated-aware) ──────────────────────────────────
    var custNo = String(r['Cust No.']||r['CustNo']||'').replace(/^'+/,'').trim();
    if(nos.length && nos.indexOf(custNo) === -1) return false;

    // ── Customer-specific field filters — match ANY listed value (OR logic) ──
    if(fEndCust.length) {
      var rEnd = String(getRowVal(r,'END Customer','ENDCustomer')||'').trim().toLowerCase();
      if(fEndCust.indexOf(rEnd) === -1) return false;
    }
    if(fBU.length) {
      var rBU = String(getRowVal(r,'Business Unit','BusinessUnit')||'').trim().toLowerCase();
      if(fBU.indexOf(rBU) === -1) return false;
    }
    if(fMaterial.length) {
      var rMat = String(getRowVal(r,'Customer Material','CustomerMaterial')||'').trim().toLowerCase();
      if(fMaterial.indexOf(rMat) === -1) return false;
    }
    if(fCoCode.length) {
      var rCC = String(getRowVal(r,'CoCode','')||'').trim().toLowerCase();
      if(fCoCode.indexOf(rCC) === -1) return false;
    }
    if(fType.length) {
      var rType = String(getRowVal(r,'TYPE','Type')||'').trim().toLowerCase();
      if(fType.indexOf(rType) === -1) return false;
    }

    // ── Include options ───────────────────────────────────────────────────────
    if(!cust.arma && String(r['ARMA']||'').trim().toUpperCase()==='ARMA') return false;
    if(!cust.credit && parseNum(getRowVal(r,'Inv Balance','InvBalance'))<0) return false;
    if(cust.promised) {
      var pd = String(r['Promised Date']||r['PromisedDate']||'').trim();
      if(pd && pd!=='null' && pd!=='') return false;
    }
    return true;
  });
}

function renderCustomerPanel(cust, filtered) {
  var cols=CSS_ST.selectedCols.length?CSS_ST.selectedCols:DEFAULT_COLS;
  var SUM_COLS=['Inv Balance'];
  var NUM_COLS=['Current','Over Due  1','Over Due 31','Over Due 61','Over Due 91','Over Due 121','Inv Balance'];
  var total=0; filtered.forEach(function(r){ total+=parseNum(getRowVal(r,'Inv Balance','InvBalance')); });
  var panelId='panel-'+cust.id;
  var nos=getConsolidatedCustNos(cust);
  var isConsolidated=nos.length>1;

  // Active filter badges
  var filterBadges = '';
  function mkBadge(lbl, val) {
    if(!val) return '';
    var parts = val.split('|').filter(Boolean);
    if(!parts.length) return '';
    var display = parts.length===1 ? parts[0] : parts.length+' values';
    return '<span class="bdg bdg-gray" title="'+esc(lbl)+': '+esc(parts.join(', '))+'">'+lbl+': '+esc(display)+'</span> ';
  }
  filterBadges += mkBadge('END', cust.endcust);
  filterBadges += mkBadge('BU',  cust.bu);
  filterBadges += mkBadge('Mat', cust.material);
  filterBadges += mkBadge('CC',  cust.cocode);
  filterBadges += mkBadge('TYPE',cust.type);
  var bannerColor=filtered.length?'var(--green)':'var(--amber)';
  var bannerBg=filtered.length?'var(--green-l)':'var(--amber-l)';
  var bannerBorder=filtered.length?'var(--green-b)':'var(--amber-b)';

  // Build sub-customer summary chips for consolidated header
  var subSummary = '';
  if(isConsolidated) {
    var subTotals = {};
    nos.forEach(function(n){ subTotals[n]={rows:0,bal:0}; });
    filtered.forEach(function(r){
      var cn=String(r['Cust No.']||r['CustNo']||'').replace(/^'+/,'').trim();
      if(subTotals[cn]){ subTotals[cn].rows++; subTotals[cn].bal+=parseNum(getRowVal(r,'Inv Balance','InvBalance')); }
    });
    subSummary = '<div style="display:flex;flex-wrap:wrap;gap:5px;margin-top:7px">'
      + nos.map(function(n){
          var st=subTotals[n]||{rows:0,bal:0};
          return '<span class="bdg '+(st.rows>0?'bdg-blue':'bdg-gray')+'" title="Cust No. '+n+'">'
            +'#'+esc(n)+': '+st.rows+' rows / '+fmtUSD(st.bal)+'</span>';
        }).join('')
      +'</div>';
  }

  var thead='<tr>'+cols.map(function(c){ return '<th'+(NUM_COLS.indexOf(c)>-1?' style="text-align:right"':'')+'>'+esc(c)+'</th>'; }).join('')+'</tr>';
  var tbody2=filtered.length
    ?filtered.slice(0,200).map(function(r){
        return '<tr>'+cols.map(function(c){
          var key=COL_KEY_MAP[c]||c;
          var val=getRowVal(r,c,key);
          var disp=String(val===null||val===undefined?'':val);
          var isNum=NUM_COLS.indexOf(c)>-1;
          if(isNum){
            var n=parseNum(val);
            disp = (n===0 && c!=='Inv Balance') ? '\u2014' : fmtUSD(n);
          }
          return '<td'+(isNum?' style="text-align:right"':'')+'>'+esc(disp)+'</td>';
        }).join('')+'</tr>';
      }).join('')
    :'<tr><td colspan="'+cols.length+'" style="text-align:center;color:var(--gray);padding:14px;font-style:italic">No open items</td></tr>';
  var tfoot2=filtered.length?'<tfoot><tr>'+cols.map(function(c){
    return SUM_COLS.indexOf(c)>-1
      ?'<td style="text-align:right;font-weight:700;background:var(--green-l);color:var(--green)">'+fmtUSD(total)+'</td>'
      :'<td></td>';
  }).join('')+'</tr></tfoot>':'';

  return '<div style="border:1px solid var(--border);border-radius:var(--r);margin-bottom:10px;overflow:hidden">'
    +'<div style="background:'+bannerBg+';border-bottom:1px solid '+bannerBorder+';padding:10px 14px" id="'+panelId+'-hdr">'
      +'<div style="display:flex;align-items:center;justify-content:space-between">'
        +'<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap">'
          +'<span style="font-weight:600;font-size:14px;color:'+bannerColor+'">'+esc(cust.name)+'</span>'
          +'<span class="bdg bdg-gray">#'+esc(cust.custno)+'</span>'
          +(isConsolidated?'<span class="bdg bdg-blue" title="'+esc(nos.join(', '))+'">Consolidated ('+nos.length+' customers)</span>':'')
          +(filtered.length?'<span class="bdg bdg-green">'+filtered.length+' rows &bull; '+fmtUSD(total)+'</span>':'<span class="bdg bdg-amber">No open items</span>')
          +(filterBadges?'<span style="font-size:10px;color:var(--gray)">|</span>'+filterBadges:'')
        +'</div>'
        +'<button class="btn btn-outline btn-sm" id="'+panelId+'-tog" data-panel="'+panelId+'" style="flex-shrink:0">Collapse</button>'
      +'</div>'
      +subSummary
    +'</div>'
    +'<div id="'+panelId+'" style="overflow-x:auto;max-height:360px;overflow-y:auto">'
      +'<table class="prev-tbl"><thead>'+thead+'</thead><tbody>'+tbody2+'</tbody>'+tfoot2+'</table>'
    +'</div>'
    +(filtered.length>200?'<div style="font-size:11px;color:var(--gray);padding:4px 14px">First 200 rows shown.</div>':'')
    +'</div>';
}

function wirePreviewPanelToggles() {
  var area=G('prev-area'); if(!area||area._togWired) return;
  area._togWired=true;
  area.addEventListener('click', function(e){
    var btn=e.target;
    while(btn&&btn!==area){ if(btn.getAttribute&&btn.getAttribute('data-panel')) break; btn=btn.parentNode; }
    if(!btn||btn===area) return;
    var pid=btn.getAttribute('data-panel');
    var el=G(pid); if(!el) return;
    if(el.style.display==='none'){ el.style.display=''; btn.textContent='Collapse'; }
    else{ el.style.display='none'; btn.textContent='Expand'; }
  });
}

function buildStatementHTML(cust, filtered) {
  var cols=CSS_ST.selectedCols.length?CSS_ST.selectedCols:DEFAULT_COLS;
  var NUM_COLS=['Current','Over Due  1','Over Due 31','Over Due 61','Over Due 91','Over Due 121','Inv Balance'];
  var nos = getConsolidatedCustNos(cust);
  var isConsolidated = nos.length > 1;
  var total=0; filtered.forEach(function(r){ total+=parseNum(getRowVal(r,'Inv Balance','InvBalance')); });
  var today=new Date().toLocaleDateString();

  // Build per-sub-customer subtotal rows for consolidated statements
  var subtotalRows = '';
  if(isConsolidated) {
    var subTotals = {};
    nos.forEach(function(n){ subTotals[n]={rows:0,bal:0,name:''}; });
    filtered.forEach(function(r){
      var cn=String(r['Cust No.']||r['CustNo']||'').replace(/^'+/,'').trim();
      var nm=String(r['Customer Name']||r['CustomerName']||'').trim();
      if(subTotals[cn]){
        subTotals[cn].rows++;
        subTotals[cn].bal+=parseNum(getRowVal(r,'Inv Balance','InvBalance'));
        if(nm && !subTotals[cn].name) subTotals[cn].name=nm;
      }
    });
    subtotalRows = '<table style="width:100%;border-collapse:collapse;margin-bottom:16px">'
      +'<thead><tr>'
        +'<th style="background:#374151;color:#fff;padding:6px 10px;font-size:11px;text-align:left">Cust No.</th>'
        +'<th style="background:#374151;color:#fff;padding:6px 10px;font-size:11px;text-align:left">Customer Name</th>'
        +'<th style="background:#374151;color:#fff;padding:6px 10px;font-size:11px;text-align:right">Rows</th>'
        +'<th style="background:#374151;color:#fff;padding:6px 10px;font-size:11px;text-align:right">Inv Balance</th>'
      +'</tr></thead><tbody>'
      +nos.map(function(n){
          var st=subTotals[n]||{rows:0,bal:0,name:''};
          return '<tr>'
            +'<td style="padding:5px 10px;border-bottom:1px solid #e5e7eb;font-size:11px">'+esc(n)+'</td>'
            +'<td style="padding:5px 10px;border-bottom:1px solid #e5e7eb;font-size:11px">'+esc(st.name)+'</td>'
            +'<td style="padding:5px 10px;border-bottom:1px solid #e5e7eb;font-size:11px;text-align:right">'+st.rows+'</td>'
            +'<td style="padding:5px 10px;border-bottom:1px solid #e5e7eb;font-size:11px;text-align:right;font-weight:600">'+fmtUSD(st.bal)+'</td>'
            +'</tr>';
        }).join('')
      +'<tr><td colspan="3" style="padding:6px 10px;font-weight:700;font-size:11px;text-align:right">Grand Total</td>'
        +'<td style="padding:6px 10px;font-weight:700;font-size:11px;text-align:right;background:#def7ec;color:#057a55">'+fmtUSD(total)+'</td></tr>'
      +'</tbody></table>';
  }

  var thead='<tr>'+cols.map(function(c){
    return '<th align="'+(NUM_COLS.indexOf(c)>-1?'right':'left')+'" style="background:#1a56db;color:#fff;padding:6px 8px;font-size:10px;white-space:nowrap">'+esc(c)+'</th>';
  }).join('')+'</tr>';
  var tbody=filtered.map(function(r){
    return '<tr>'+cols.map(function(c){
      var key=COL_KEY_MAP[c]||c;
      var val=getRowVal(r,c,key);
      var disp=String(val===null||val===undefined?'':val);
      var isNum=NUM_COLS.indexOf(c)>-1;
      if(isNum){
        var n=parseNum(val);
        disp = (n===0 && c!=='Inv Balance') ? '' : fmtUSD(n);
      }
      return '<td align="'+(isNum?'right':'left')+'" style="padding:5px 8px;border-bottom:1px solid #e5e7eb;font-size:10px">'+esc(disp)+'</td>';
    }).join('')+'</tr>';
  }).join('');
  var tfoot='<tr>'+cols.map(function(c){
    return c==='Inv Balance'
      ?'<td align="right" style="padding:6px 8px;font-weight:700;background:#def7ec;color:#057a55;font-size:10px">'+fmtUSD(total)+'</td>'
      :'<td></td>';
  }).join('')+'</tr>';

  var consolidatedNote = isConsolidated
    ? '<div style="background:#eff6ff;border:1px solid #bfdbfe;border-radius:6px;padding:8px 12px;margin-bottom:12px;font-size:11px;color:#1e40af">'
        +'<b>Consolidated Statement</b> — this statement covers '+nos.length+' customer accounts: '+nos.map(function(n){return '#'+esc(n);}).join(', ')
      +'</div>'
    : '';

  return '<html><head><meta charset="UTF-8"></head>'
    +'<body style="font-family:Arial,sans-serif;font-size:12px;color:#111;margin:0;padding:16px">'
    +'<div style="background:#1a56db;color:#fff;padding:14px 18px;margin-bottom:12px;border-radius:4px">'
      +'<b style="font-size:15px">Account Statement &#8212; '+esc(cust.name)+'</b>'
      +(isConsolidated?'<span style="opacity:.8;font-size:11px;margin-left:8px">(Consolidated)</span>':'')
      +'<div style="font-size:10px;margin-top:4px;opacity:.85">Cust No: '+esc(cust.custno)+' &bull; Date: '+today+'</div>'
    +'</div>'
    +'<p style="margin-bottom:12px;font-size:13px">Dear '+esc(cust.name)+',</p>'
    +'<p style="margin-bottom:12px;font-size:13px">Please find your account statement as of '+today+' below. Kindly review and contact us with any discrepancies.</p>'
    +consolidatedNote
    +(isConsolidated?'<p style="font-weight:600;font-size:12px;margin-bottom:6px">Summary by customer account:</p>'+subtotalRows:'')
    +(isConsolidated?'<p style="font-weight:600;font-size:12px;margin-bottom:6px">Detail:</p>':'')
    +'<table style="width:100%;border-collapse:collapse;margin-bottom:14px"><thead>'+thead+'</thead><tbody>'+tbody+'</tbody><tfoot>'+tfoot+'</tfoot></table>'
    +'<p style="font-size:10px;color:#6b7280;border-top:1px solid #e5e7eb;padding-top:10px">Automated statement &#8212; NVIDIA AR Collections. NVARCollection@nvidia.com</p>'
    +'</body></html>';
}

function renderPreviewForCust(cid) {
  var sel=G('prev-cust'); if(!sel) return;
  for(var i=0;i<sel.options.length;i++) sel.options[i].selected=(Number(sel.options[i].value)===cid);
  renderPreview();
}

// Build XLS for a customer and trigger browser download
function downloadStatementXLS(cust, filtered) {
  var cols=CSS_ST.selectedCols.length?CSS_ST.selectedCols:DEFAULT_COLS;
  // Build rows array matching selected columns
  var rows=filtered.map(function(r){
    var obj={};
    cols.forEach(function(c){
      var key=COL_KEY_MAP[c]||c;
      var val=r[c]!==undefined?r[c]:(r[key]!==undefined?r[key]:'');
      obj[c]=val===null||val===undefined?'':val;
    });
    return obj;
  });
  var ws=XLSX.utils.json_to_sheet(rows);
  ws['!cols']=cols.map(function(c){return {wch:Math.max(c.length+2,12)};});
  var wb=XLSX.utils.book_new();
  var sheetName=String(cust.name).replace(/[:\\/?*\[\]]/g,'').substring(0,28);
  XLSX.utils.book_append_sheet(wb,ws,sheetName||'Statement');
  var today=new Date().toISOString().split('T')[0];
  var fname='Statement_'+cust.custno+'_'+today+'.xlsx';
  var buf=XLSX.write(wb,{bookType:'xlsx',type:'array'});
  var blob=new Blob([new Uint8Array(buf)],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  var url=URL.createObjectURL(blob);
  var a=document.createElement('a'); a.href=url; a.download=fname;
  document.body.appendChild(a); a.click();
  document.body.removeChild(a);
  setTimeout(function(){URL.revokeObjectURL(url);},2000);
  return fname;
}

// ════════════════════════════════════════════════════════════════════════════
// MICROSOFT GRAPH API — Token via SharePoint's built-in OAuth endpoint
//
// Strategy: SharePoint itself is registered in the tenant and pre-authorized
// for Graph. We call SP's /_api/SP.OAuth.Token/Acquire with the page's digest
// to get a delegated Graph access token — no custom app registration needed,
// no popup, no MSAL, works silently using the user's existing SP session.
// ════════════════════════════════════════════════════════════════════════════

var GRAPH_TOKEN     = null;
var GRAPH_TOKEN_EXP = 0;
var GRAPH_USER      = {};

function graphGetToken() {
  if(GRAPH_TOKEN && Date.now() < GRAPH_TOKEN_EXP - 300000) return Promise.resolve(GRAPH_TOKEN);

  // Try multiple approaches in order:
  // 1. SP.OAuth.Token/Acquire with form-encoded body (correct format)
  // 2. SP.OAuth.Token/Acquire with resource as query param
  // 3. _api/SP.AppContextSite token endpoint
  // 4. ADAL token from window context if available

  // Check if SharePoint ADAL token is available in page context
  if(window._spPageContextInfo && window._spPageContextInfo.aadTokenProviderFactory) {
    try {
      var factory = window._spPageContextInfo.aadTokenProviderFactory;
      if(factory && factory.getTokenProvider) {
        return factory.getTokenProvider().then(function(provider){
          return provider.getToken('https://graph.microsoft.com');
        }).then(function(token){
          GRAPH_TOKEN = token;
          GRAPH_TOKEN_EXP = Date.now() + 3000000;
          return GRAPH_TOKEN;
        });
      }
    } catch(e) {}
  }

  // Try SP.OAuth.Token/Acquire with form-encoded body
  var url = spOrigin() + SP_BASE + '/_api/SP.OAuth.Token/Acquire';
  return getDigest().then(function(digest){
    return fetch(url, {
      method: 'POST',
      credentials: 'include',
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/x-www-form-urlencoded',
        'X-RequestDigest': digest
      },
      body: 'resource=https%3A%2F%2Fgraph.microsoft.com'
    }).then(function(res){
      if(res.ok) return res.json();
      // Try alternate: resource as query param with empty body
      return fetch(url+'?resource='+encodeURIComponent('https://graph.microsoft.com'), {
        method: 'POST',
        credentials: 'include',
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'X-RequestDigest': digest,
          'Content-Length': '0'
        }
      }).then(function(r){
        if(!r.ok) throw new Error('SP OAuth HTTP '+r.status);
        return r.json();
      });
    });
  }).then(function(data){
    var token = data && data.d && data.d.BearerToken;
    if(!token) throw new Error('SP OAuth: no token returned — site may block delegated Graph access');
    GRAPH_TOKEN     = token;
    GRAPH_TOKEN_EXP = Date.now() + 3500000;
    return GRAPH_TOKEN;
  });
}

function graphGetMe() {
  return graphGetToken().then(function(token){
    return fetch('https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName', {
      headers: { 'Authorization': 'Bearer '+token, 'Accept': 'application/json' }
    });
  }).then(function(r){
    if(!r.ok) throw new Error('Graph /me HTTP '+r.status);
    return r.json();
  }).then(function(me){
    GRAPH_USER = { displayName: me.displayName||'', mail: me.mail||me.userPrincipalName||'' };
    return GRAPH_USER;
  });
}

function graphConnect() {
  var btn = G('btn-graph-auth');
  if(btn){ btn.disabled=true; btn.textContent='Connecting...'; }
  graphGetToken().then(function(){
    return graphGetMe();
  }).then(function(me){
    showGraphConnected(me);
    if(btn){ btn.disabled=false; btn.textContent='Connect Outlook'; }
  }).catch(function(err){
    if(btn){ btn.disabled=false; btn.textContent='Connect Outlook'; }
    var msg = err.message||String(err);
    // SP OAuth 400/403 = tenant policy blocks delegated Graph access from SP pages
    // This is common in enterprise tenants with strict Conditional Access policies
    if(msg.indexOf('400')>-1 || msg.indexOf('403')>-1 || msg.indexOf('no token')>-1) {
      showErr('Direct Graph access is blocked by NVIDIA IT policy. '
        +'Use ✉ Option B (Open in Outlook) which opens drafts in your local Outlook, '
        +'or ▶ Option A (Power Automate) which sends from NVARCollection@nvidia.com in the background.');
    } else {
      showErr('Graph auth failed: '+msg+'. Use Option A or Option B instead.');
    }
  });
}

function showGraphConnected(me) {
  var banner = G('graph-auth-banner'); if(banner) banner.style.display='none';
  var ok = G('graph-auth-ok');
  if(ok){ ok.style.display='flex'; ok.style.alignItems='center'; }
  txt('graph-user-name', (me.displayName||me.mail)+' ('+(me.mail||'')+')');
}

// Silent connect — used by sendDirectly() before sending

// ── POWER AUTOMATE HTTP TRIGGER ──────────────────────────────────────────────
var PA_TRIGGER_URL = 'https://default43083d15727340c1b7db39efd9ccc1.7a.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/d408235839a7419cab763c1ca9272514/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=GeDlqzY-Rneaq-Wbeq8a41GxSJ7T9pU_7b9OYREoeD4';

function runViaPowerAutomate() {
  var custs = getSelectedCustomers();
  if(!custs.length){ alert('Select at least one customer first.'); return; }
  if(!CSS_ST.agingLoaded){ alert('Aging data still loading — please wait.'); return; }

  var urlInp = G('inp-pa-url');
  var url = (urlInp && urlInp.value.trim()) || PA_TRIGGER_URL;
  if(!url){ showErr('Power Automate URL is not configured.'); return; }

  var progDiv = G('send-progress'); if(progDiv) progDiv.style.display='';
  var progBar = G('send-prog-bar');  if(progBar) progBar.style.width='30%';
  var progTxt = G('send-prog-text'); if(progTxt) progTxt.textContent='Calling Power Automate...';
  var results = G('send-results');   if(results) results.innerHTML='';

  var custNos = custs.map(function(c){ return c.custno; });
  var payload = {
    triggeredBy: 'CustomerStatementSetup',
    triggerTime: new Date().toISOString(),
    customers: custs.map(function(c){
      var filtered = filterAgingForCustomer(c);
      var total = 0;
      filtered.forEach(function(r){ total += parseNum(getRowVal(r,'Inv Balance','InvBalance')); });
      return { custno:c.custno, name:c.name, to:c.to, cc:c.cc,
               collector:c.collector, schedule:c.schedule,
               rowCount:filtered.length, totalBal:total };
    }),
    selectedCustNos: custNos
  };

  fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  }).then(function(res){
    if(progBar) progBar.style.width='100%';
    if(res.status===200||res.status===202||res.status===204){
      if(progTxt) progTxt.textContent='Power Automate triggered successfully!';
      if(results) results.innerHTML=
        '<div style="background:var(--green-l);border:1px solid var(--green-b);border-radius:7px;padding:10px 14px;font-size:13px;color:var(--green)">'
          +'&#10003; <strong>Flow triggered</strong> — Power Automate will send '
          +custs.length+' statement(s) from NVARCollection@nvidia.com.'
        +'</div>';
    } else {
      return res.text().then(function(t){ throw new Error('HTTP '+res.status+': '+t.slice(0,200)); });
    }
  }).catch(function(err){
    if(progBar) progBar.style.width='100%';
    if(progTxt) progTxt.textContent='Error';
    if(results) results.innerHTML=
      '<div style="background:var(--red-l);border:1px solid var(--red-b);border-radius:7px;padding:10px 14px;font-size:13px;color:var(--red)">'
        +'&#10007; PA trigger failed: '+esc(err.message)
      +'</div>';
  });
}

// ── GRAPH BATCH SEND ─────────────────────────────────────────────────────────
function runEmailBatch(mode) {
  var custs = getSelectedCustomers();
  if(!custs.length){ alert('Select at least one customer first.'); return; }
  if(!CSS_ST.agingLoaded){ alert('Aging data not loaded yet.'); return; }

  // EML mode needs no auth — only send/draft need Graph token
  if((mode==='send'||mode==='draft') && !GRAPH_TOKEN) {
    showErr('Connect Outlook first before sending directly.');
    return;
  }

  var progDiv = G('send-progress'); if(progDiv) progDiv.style.display='';
  var progBar = G('send-prog-bar');
  var progTxt = G('send-prog-text');
  var results = G('send-results');  if(results) results.innerHTML='';

  var total = custs.length, done = 0, successes = [], failures = [];

  function updateProgress() {
    var pct = Math.round((done/total)*100);
    if(progBar) progBar.style.width = pct+'%';
    if(progTxt) progTxt.textContent = done+' / '+total+' processed...';
  }

  function renderResults() {
    if(!results) return;
    var html = '';
    if(successes.length) {
      html += '<div style="margin-bottom:8px"><div style="font-size:12px;font-weight:600;color:var(--green);margin-bottom:4px">'
        +'&#10003; '+(mode==='draft'?'Saved to Drafts':'Sent')+': '+successes.length+'</div>'
        +successes.map(function(c){ return '<div style="font-size:12px;padding:4px 0;border-bottom:1px solid var(--border)">'
          +'<span class="bdg bdg-green">&#10003;</span> '+esc(c.name)+'</div>'; }).join('')+'</div>';
    }
    if(failures.length) {
      html += '<div><div style="font-size:12px;font-weight:600;color:var(--red);margin-bottom:4px">&#10007; Failed: '+failures.length+'</div>'
        +failures.map(function(f){ return '<div style="font-size:12px;padding:4px 0;border-bottom:1px solid var(--border)">'
          +'&#10007; '+esc(f.cust.name)+' — '+esc(f.err)+'</div>'; }).join('')+'</div>';
    }
    results.innerHTML = html;
  }

  function processNext(idx) {
    if(idx >= custs.length) {
      updateProgress(); renderResults();
      if(progTxt) progTxt.textContent = 'Done — '+successes.length+' '+(mode==='draft'?'draft(s)':'sent')
        +(failures.length ? ', '+failures.length+' failed' : '');
      return;
    }
    var cust = custs[idx];
    var filtered = filterAgingForCustomer(cust);
    if(progTxt) progTxt.textContent = 'Processing '+esc(cust.name)+'...';

    var promise;
    if(mode==='eml') {
      try { buildAndDownloadEML(cust, filtered); promise = Promise.resolve({ok:true}); }
      catch(e) { promise = Promise.reject(new Error(e.message)); }
    } else if(mode==='send') {
      promise = graphSendOne(cust, filtered);
    } else {
      promise = graphSaveDraftOne(cust, filtered).then(function(){ return {ok:true}; });
    }

    promise.then(function(){ successes.push(cust); })
    .catch(function(err){ failures.push({cust:cust, err:err.message||String(err)}); })
    .then(function(){ done++; updateProgress(); setTimeout(function(){ processNext(idx+1); }, mode==='eml'?800:500); });
  }
  updateProgress();
  processNext(0);
}

// ── GRAPH SEND/DRAFT ─────────────────────────────────────────────────────────
function buildGraphMessage(cust, filtered) {
  var today = new Date().toISOString().split('T')[0];
  var todayFmt = new Date().toLocaleDateString();
  var tplSubj = ((G('tpl-subject')||{}).value||'Account Statement - {CustomerName} - {Date}').trim();
  var tplBody = ((G('tpl-body')||{}).value||'Dear {CustomerName},\n\nPlease find attached your account statement as of {Date}.\n\nBest regards,\nNVIDIA AR Collections').trim();
  var collector = cust.collector||'';
  var subj = tplSubj.replace(/{CustomerName}/g,cust.name).replace(/{Date}/g,todayFmt).replace(/{Collector}/g,collector);
  var body = tplBody.replace(/{CustomerName}/g,cust.name).replace(/{Date}/g,todayFmt).replace(/{Collector}/g,collector);
  var toArr = (cust.to||[]).filter(function(e){return e&&e.trim();});
  var ccArr = (cust.cc||[]).filter(function(e){return e&&e.trim();});
  if(collector&&collector.trim()&&ccArr.indexOf(collector.trim())===-1) ccArr.push(collector.trim());
  var xlsBuf = buildStatementXLSBuf(cust, filtered);
  var xlsB64 = arrayBufferToBase64(xlsBuf);
  var xlsFname = 'Statement_'+cust.custno+'_'+today+'.xlsx';
  return {
    subject: subj, importance:'normal',
    body: { contentType:'text', content:body },
    toRecipients: toArr.map(function(e){return {emailAddress:{address:e}};}),
    ccRecipients: ccArr.map(function(e){return {emailAddress:{address:e}};}),
    attachments: [{ '@odata.type':'#microsoft.graph.fileAttachment',
      name:xlsFname,
      contentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      contentBytes:xlsB64 }]
  };
}

function graphSendOne(cust, filtered) {
  return graphGetToken().then(function(token){
    var msg = buildGraphMessage(cust, filtered);
    return fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
      method:'POST',
      headers:{'Authorization':'Bearer '+token,'Content-Type':'application/json'},
      body: JSON.stringify({message:msg, saveToSentItems:true})
    });
  }).then(function(res){
    if(res.status===202||res.status===200) return {ok:true};
    return res.json().then(function(e){ throw new Error((e.error&&e.error.message)||'HTTP '+res.status); });
  });
}

function graphSaveDraftOne(cust, filtered) {
  return graphGetToken().then(function(token){
    var msg = buildGraphMessage(cust, filtered);
    return fetch('https://graph.microsoft.com/v1.0/me/messages', {
      method:'POST',
      headers:{'Authorization':'Bearer '+token,'Content-Type':'application/json'},
      body: JSON.stringify(msg)
    });
  }).then(function(res){
    if(res.status===201||res.status===200) return res.json();
    return res.json().then(function(e){ throw new Error((e.error&&e.error.message)||'HTTP '+res.status); });
  });
}

// ── PREVIEW CHECKLIST ─────────────────────────────────────────────────────────
var PAGE_SIZE = 20;

function populateCollectorFilter() {
  var sel = G('filter-collector'); if(!sel) return;
  var collectors = {};
  CSS_ST.customers.forEach(function(c){
    var col = c.collector||(c.cc&&c.cc[0])||'';
    if(col) collectors[col]=1;
  });
  var prev = sel.value;
  sel.innerHTML = '<option value="">— All collectors —</option>'
    +Object.keys(collectors).sort().map(function(c){
      return '<option value="'+esc(c)+'"'+(c===prev?' selected':'')+'>'+esc(c)+'</option>';
    }).join('');

  var gsel = G('filter-group'); if(!gsel) return;
  var groups = {};
  CSS_ST.customers.forEach(function(c){ if(c.group) groups[c.group]=1; });
  var gprev = gsel.value;
  gsel.innerHTML = '<option value="">— All groups —</option>'
    +Object.keys(groups).sort().map(function(g){
      return '<option value="'+esc(g)+'"'+(g===gprev?' selected':'')+'>'+esc(g)+'</option>';
    }).join('');
}

function renderChecklist() {
  var el = G('cust-checklist'); if(!el) return;
  var col    = (G('filter-collector')||{}).value||'';
  var grpFlt = (G('filter-group')||{}).value||'';
  var srch   = ((G('prev-cust-search')||{}).value||'').toLowerCase().trim();
  var custs  = CSS_ST.customers.filter(function(c){
    if(col && c.collector!==col && !(c.cc&&c.cc.indexOf(col)>-1)) return false;
    if(grpFlt && (c.group||'')!==grpFlt) return false;
    if(srch) {
      var all=[c.custno,c.name,c.to.join(' '),c.group||'',c.collector||''].join(' ').toLowerCase();
      if(all.indexOf(srch)===-1) return false;
    }
    return true;
  });

  var totalPages = Math.max(1, Math.ceil(custs.length / PAGE_SIZE));
  if(CSS_ST.prevPage >= totalPages) CSS_ST.prevPage = totalPages-1;
  var page = CSS_ST.prevPage;
  var paged = custs.slice(page*PAGE_SIZE, (page+1)*PAGE_SIZE);

  var pageInfo = G('page-info');
  var btnPrev  = G('btn-prev-page');
  var btnNext  = G('btn-next-page');
  if(pageInfo) pageInfo.textContent = 'Page '+(page+1)+' / '+totalPages+' ('+custs.length+')';
  if(btnPrev)  btnPrev.disabled  = page <= 0;
  if(btnNext)  btnNext.disabled  = page >= totalPages-1;
  var pagDiv = G('checklist-pagination');
  if(pagDiv) pagDiv.style.display = totalPages <= 1 ? 'none' : 'flex';

  var groups={};
  paged.forEach(function(c){
    var g = c.group||c.collector||(c.cc&&c.cc[0])||'(no group)';
    if(!groups[g]) groups[g]=[];
    groups[g].push(c);
  });

  if(!paged.length){
    el.innerHTML='<div class="empty">No customers match the selected filters.</div>';
    txt('checklist-count','');
    return;
  }

  var html='';
  Object.keys(groups).sort().forEach(function(grp){
    html+='<div style="padding:6px 8px 2px;background:#f3f4f6;font-size:11px;font-weight:600;color:var(--gray);'
      +'position:sticky;top:0;display:flex;align-items:center;gap:8px;border-bottom:1px solid var(--border)">'
      +'<label style="display:flex;align-items:center;gap:6px;cursor:pointer">'
        +'<input type="checkbox" class="grp-chk" data-grp="'+esc(grp)+'" style="cursor:pointer"> '
        +esc(grp)
      +'</label>'
      +'<span style="font-weight:400;color:var(--gray)">('+groups[grp].length+')</span>'
    +'</div>';
    groups[grp].forEach(function(c){
      var selected = CSS_ST._checklistSel&&CSS_ST._checklistSel[c.id];
      html+='<div style="display:flex;align-items:center;gap:10px;padding:6px 12px;border-bottom:1px solid #f9fafb" data-cid="'+c.id+'">'
        +'<input type="checkbox" class="cust-chk" data-cid="'+c.id+'"'+(selected?' checked':'')+' style="cursor:pointer;flex-shrink:0">'
        +'<div style="flex:1;min-width:0">'
          +'<div style="font-size:13px;font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">'+esc(c.name)+'</div>'
          +'<div style="font-size:11px;color:var(--gray)">#'+esc(c.custno)
            +(c.to&&c.to.length?' &bull; '+esc(c.to[0]):'')
            +' <span class="bdg bdg-gray">'+esc(c.schedule)+'</span>'
          +'</div>'
        +'</div>'
      +'</div>';
    });
  });

  el.innerHTML = html;
  wireChecklist();
  updateChecklistCount();
}

function wireChecklist() {
  var el=G('cust-checklist'); if(!el||el._clWired) return;
  el._clWired=true;
  el.addEventListener('change', function(e){
    var t=e.target;
    if(!CSS_ST._checklistSel) CSS_ST._checklistSel={};
    if(t.classList.contains('cust-chk')){
      var cid=Number(t.getAttribute('data-cid'));
      CSS_ST._checklistSel[cid]=t.checked;
      updateChecklistCount();
    } else if(t.classList.contains('grp-chk')){
      var grp=t.getAttribute('data-grp');
      var all=Array.from(el.children);
      var capture=false;
      all.forEach(function(div){
        var ghk=div.querySelector('.grp-chk');
        if(ghk){ capture=(ghk.getAttribute('data-grp')===grp); return; }
        if(capture){
          var chk=div.querySelector('.cust-chk');
          if(chk){ chk.checked=t.checked; CSS_ST._checklistSel[Number(chk.getAttribute('data-cid'))]=t.checked; }
        }
      });
      updateChecklistCount();
    }
  });
}

function updateChecklistCount() {
  var sel=getChecklistSelected();
  txt('checklist-count', sel.length+' customer(s) selected');
  populatePreviewDropdownFromChecklist(sel);
}

function getChecklistSelected() {
  if(!CSS_ST._checklistSel) return [];
  return CSS_ST.customers.filter(function(c){ return CSS_ST._checklistSel[c.id]; });
}

function populatePreviewDropdownFromChecklist(sel) {
  var selEl=G('prev-cust'); if(!selEl) return;
  for(var i=0;i<selEl.options.length;i++){
    var cid=Number(selEl.options[i].value);
    selEl.options[i].selected = !!CSS_ST._checklistSel[cid];
  }
}

// ── Build EML file: template body text + XLS & HTML as attachments ───────────

// ── Send directly via Graph (reuses runEmailBatch) ────────────────────────────
function sendDirectly(mode) {
  // mode: 'send' | 'draft'
  if(!GRAPH_TOKEN) {
    showErr('Connect your Outlook account first — click the "Connect Outlook" button in Option C below.');
    var banner = G('graph-auth-banner');
    if(banner) banner.scrollIntoView({behavior:'smooth', block:'nearest'});
    return;
  }
  runEmailBatch(mode);
}

function graphConnect_silent() {
  // Attempt silent token acquisition without showing UI
  return graphGetToken().then(function(){
    return graphGetMe();
  }).then(function(me){
    showGraphConnected(me);
  });
}

// ── Build EML file: template body text + XLS & HTML as attachments ───────────
// Body = plain text from Email Template fields only (no statement data in body)
// Attachments: Statement.xlsx + Statement.html (full formatted statement)
function buildAndDownloadEML(cust, filtered) {
  var today   = new Date().toISOString().split('T')[0];
  var todayFmt= new Date().toLocaleDateString();
  var nos     = getConsolidatedCustNos(cust);

  var toArr=(cust.to||[]).filter(function(e){return e&&e.trim();});
  var ccArr=(cust.cc||[]).filter(function(e){return e&&e.trim();});
  if(cust.collector&&cust.collector.trim()&&ccArr.indexOf(cust.collector.trim())===-1)
    ccArr.push(cust.collector.trim());
  var collector=cust.collector||ccArr[0]||'';

  // Use Email Template fields from Filters & Columns tab
  var tplSubj=((G('tpl-subject')||{}).value||'Account Statement - {CustomerName} - {Date}').trim();
  var tplBody=((G('tpl-body')||{}).value||'Dear {CustomerName},\n\nPlease find attached your account statement as of {Date}.\n\nBest regards,\nNVIDIA AR Collections').trim();

  var subj=tplSubj
    .replace(/{CustomerName}/g,cust.name)
    .replace(/{Date}/g,todayFmt)
    .replace(/{Collector}/g,collector);
  var bodyText=tplBody
    .replace(/{CustomerName}/g,cust.name)
    .replace(/{Date}/g,todayFmt)
    .replace(/{Collector}/g,collector);

  // Attachments
  var xlsBuf  =buildStatementXLSBuf(cust,filtered);
  var xlsB64  =arrayBufferToBase64(xlsBuf);
  var xlsFname='Statement_'+cust.custno+'_'+today+'.xlsx';

  var htmlContent=buildStatementHTML(cust,filtered);
  var htmlBytes  =new TextEncoder().encode(htmlContent);
  var htmlB64    =arrayBufferToBase64(htmlBytes.buffer);
  var htmlFname  ='Statement_'+cust.custno+'_'+today+'.html';

  var boundary='NvidiaAR_'+Date.now()+'_bound';

  var emlLines=[
    'MIME-Version: 1.0',
    'From: NVARCollection@nvidia.com',
    'To: '+toArr.join(', '),
    ccArr.length?'Cc: '+ccArr.join(', '):'',
    'Subject: =?UTF-8?Q?'+encodeQPHeader(subj)+'?=',
    'X-Unsent: 1',
    'Content-Type: multipart/mixed; boundary="'+boundary+'"',
    '',
    // Part 1: plain text body from template
    '--'+boundary,
    'Content-Type: text/plain; charset=utf-8',
    'Content-Transfer-Encoding: quoted-printable',
    '',
    encodeQP(bodyText),
    '',
    // Part 2: Excel attachment
    '--'+boundary,
    'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'Content-Transfer-Encoding: base64',
    'Content-Disposition: attachment; filename="'+xlsFname+'"',
    '',
    xlsB64.match(/.{1,76}/g).join('\r\n'),
    '',
    // Part 3: HTML statement attachment
    '--'+boundary,
    'Content-Type: text/html; charset=utf-8',
    'Content-Transfer-Encoding: base64',
    'Content-Disposition: attachment; filename="'+htmlFname+'"',
    '',
    htmlB64.match(/.{1,76}/g).join('\r\n'),
    '',
    '--'+boundary+'--'
  ].filter(function(l){return l!==false&&l!==undefined;}).join('\r\n');

  var blob=new Blob([emlLines],{type:'message/rfc822'});
  var url=URL.createObjectURL(blob);
  var a=document.createElement('a');
  a.href=url; a.download='Statement_'+cust.custno+'_'+today+'.eml';
  document.body.appendChild(a); a.click();
  document.body.removeChild(a);
  setTimeout(function(){URL.revokeObjectURL(url);},3000);
  return a.download;
}

function encodeQPHeader(str){
  var out='';
  for(var i=0;i<str.length;i++){
    var code=str.charCodeAt(i);
    out+=(code>127||code===61)?'='+('0'+code.toString(16).toUpperCase()).slice(-2):str[i];
  }
  return out;
}

// Build XLS ArrayBuffer (reused by both EML attachment and direct download)
function buildStatementXLSBuf(cust, filtered) {
  var cols=CSS_ST.selectedCols.length?CSS_ST.selectedCols:DEFAULT_COLS;
  var rows=filtered.map(function(r){
    var obj={};
    cols.forEach(function(c){
      var key=COL_KEY_MAP[c]||c;
      var val=getRowVal(r,c,key);
      obj[c]=val===null||val===undefined?'':val;
    });
    return obj;
  });
  var ws=XLSX.utils.json_to_sheet(rows);
  ws['!cols']=cols.map(function(c){return {wch:Math.max(c.length+2,12)};});
  var wb=XLSX.utils.book_new();
  var sname=String(cust.name).replace(/[:\\/\?\*\[\]]/g,'').substring(0,28)||'Statement';
  XLSX.utils.book_append_sheet(wb,ws,sname);
  return XLSX.write(wb,{bookType:'xlsx',type:'array'});
}

function downloadStatementXLS(cust, filtered) {
  var buf    = buildStatementXLSBuf(cust,filtered);
  var today  = new Date().toISOString().split('T')[0];
  var fname  = 'Statement_'+cust.custno+'_'+today+'.xlsx';
  var blob   = new Blob([new Uint8Array(buf)],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  var url    = URL.createObjectURL(blob);
  var a      = document.createElement('a'); a.href=url; a.download=fname;
  document.body.appendChild(a); a.click();
  document.body.removeChild(a);
  setTimeout(function(){URL.revokeObjectURL(url);},2000);
  return fname;
}

// Convert ArrayBuffer to base64 string
function arrayBufferToBase64(buf) {
  var bytes = new Uint8Array(buf);
  var binary='';
  var chunk=8192;
  for(var i=0;i<bytes.length;i+=chunk){
    binary+=String.fromCharCode.apply(null, bytes.subarray(i,Math.min(i+chunk,bytes.length)));
  }
  return btoa(binary);
}

// Quoted-printable encoder for HTML body (keeps ASCII safe, encodes high chars)
function encodeQP(str) {
  // Simple version: encode non-ASCII and special chars, wrap at 76 chars
  var out='';
  var line='';
  for(var i=0;i<str.length;i++){
    var code=str.charCodeAt(i);
    var ch=str[i];
    var encoded;
    if(code===13||code===10){
      // CRLF
      out+=line.replace(/ $/,'=20')+'\r\n';
      line='';
      if(code===13&&str.charCodeAt(i+1)===10) i++;
      continue;
    } else if(code<32||code>126||code===61){
      encoded='='+('0'+code.toString(16).toUpperCase()).slice(-2);
    } else {
      encoded=ch;
    }
    if(line.length+encoded.length>75){
      out+=line+'=\r\n';
      line='';
    }
    line+=encoded;
  }
  if(line) out+=line;
  return out;
}

function sendViaOutlook() {
  var custs=getSelectedCustomers();
  if(!custs.length){ alert('Select at least one customer first.'); return; }
  if(!CSS_ST.agingLoaded){ alert('Aging data not loaded yet.'); return; }

  var delay=0;
  var links=[];
  var statusEl=G('send-outlook-status');
  var linksEl=G('send-outlook-links');

  custs.forEach(function(cust){
    var filtered=filterAgingForCustomer(cust);
    var total=0; filtered.forEach(function(r){ total+=parseNum(getRowVal(r,'Inv Balance','InvBalance')); });
    var toArr=(cust.to||[]).filter(function(e){return e&&e.trim();});
    var ccArr=(cust.cc||[]).filter(function(e){return e&&e.trim();});
    if(cust.collector&&cust.collector.trim()&&ccArr.indexOf(cust.collector.trim())===-1) ccArr.push(cust.collector.trim());
    var nos=getConsolidatedCustNos(cust);

    (function(c,f,ta,ca,tot,n){
      setTimeout(function(){
        // Download .eml file — opens in Outlook with HTML body + XLS attached
        var emlFname = buildAndDownloadEML(c,f);
        // Show instruction banner
        if(statusEl) statusEl.textContent='Opening — double-click the downloaded .eml file to open in Outlook';
      }, delay);
    })(cust,filtered,toArr,ccArr,total,nos);

    delay+=2000;

    links.push(
      '<div style="background:var(--green-l);border:1px solid var(--green-b);border-radius:7px;padding:10px 14px;margin-bottom:8px">'
        +'<div style="display:flex;align-items:center;gap:8px;margin-bottom:6px">'
          +'<span class="bdg bdg-green">&#10003;</span>'
          +'<span style="font-size:13px;font-weight:600">'+esc(cust.name)+'</span>'
          +(nos.length>1?'<span class="bdg bdg-blue">Consolidated</span>':'')
        +'</div>'
        +'<div style="font-size:12px;display:grid;grid-template-columns:55px 1fr;gap:3px;color:var(--gray)">'
          +'<span>To:</span><span style="color:#111">'+esc(toArr.join(', ')||'(none set)')+'</span>'
          +(ccArr.length?'<span>CC:</span><span style="color:#111">'+esc(ccArr.join(', '))+'</span>':'')
          +'<span>Rows:</span><span style="color:#111">'+filtered.length+' &bull; '+fmtUSD(total)+'</span>'
        +'</div>'
      +'</div>'
    );
  });

  if(linksEl) linksEl.innerHTML=
    '<div style="background:var(--blue-l);border:1px solid var(--blue-b);border-radius:7px;padding:9px 14px;margin-bottom:10px;font-size:12px;color:#1e3a8a">'
      +'<b>&#128206; EML files downloaded</b> — double-click each <code>.eml</code> file in your Downloads folder to open in Outlook. '
      +'The email is pre-filled (To, CC, Subject, HTML body) with the Excel statement already attached. Just click Send.'
    +'</div>'
    +links.join('');
  if(statusEl) setTimeout(function(){ statusEl.textContent=custs.length+' EML file(s) generated'; },delay+500);
}


// ── MODAL ────────────────────────────────────────────────
function openModal(custId) {
  CSS_ST.toTags=[]; CSS_ST.ccTags=[]; CSS_ST.editingId=custId||null;
  var isEdit = (custId!==null&&custId!==undefined);
  txt('modal-title', isEdit?'Edit Customer':'Add Customer');

  // Clear all fields
  ['m-custno','m-custname','m-type','m-consolidated','m-group']
    .forEach(function(id){ var el=G(id); if(el) el.value=''; });
  // Clear filter tags
  FILTER_FIELDS.forEach(function(f){ CSS_ST.fTags[f]=[]; renderFilterTags(f); });
  G('m-emailtype').value='Statement';
  G('m-arma').checked=true; G('m-credit').checked=true; G('m-promised').checked=false;
  G('m-schedule').value='None'; G('m-schedtime').value=1; // time field removed

  if(isEdit) {
    var c=null;
    for(var i=0;i<CSS_ST.customers.length;i++){ if(CSS_ST.customers[i].id===custId){c=CSS_ST.customers[i];break;} }
    if(c) {
      G('m-custno').value    = c.custno;
      G('m-custname').value  = c.name;
      G('m-emailtype').value = c.emailtype||'Statement';
      // collector + cocode set AFTER populateCollectorSelect/populateCocodeSelect below
      // Restore filter multi-tags from stored | values
      CSS_ST.fTags.endcust  = strToFTags(c.endcust);
      CSS_ST.fTags.bu       = strToFTags(c.bu);
      CSS_ST.fTags.material = strToFTags(c.material);
      CSS_ST.fTags.cocode   = strToFTags(c.cocode);
      CSS_ST.fTags.type     = strToFTags(c.type);
      FILTER_FIELDS.forEach(function(f){ renderFilterTags(f); });
      G('m-type').value      = c.type||'';
      G('m-consolidated').value = c.consolidated||'';
      G('m-group').value     = c.group||'';
      G('m-arma').checked    = c.arma;
      G('m-credit').checked  = c.credit;
      G('m-promised').checked= c.promised;
      G('m-schedule').value  = c.schedule||'Weekly';
      G('m-schedtime').value = c.schedtime||1;
      // time field removed
      CSS_ST.toTags = c.to.slice();
      CSS_ST.ccTags = c.cc.slice();
    }
  }
  renderTagWrap('to'); renderTagWrap('cc'); updateSchedHint();
  // Populate aging-driven dropdowns
  populateModalDropdowns();
  if(isEdit && c) {
    // Restore collector and cocode from aging data for this customer
    var rec = CSS_ST.agingIndex && (CSS_ST.agingIndex[c.custno] || null);
    if(rec) {
      // Populate options then pre-select the saved value
      populateCollectorSelect(rec.collectors, c.collector);
      populateCocodeSelect(rec.cocodes, c.cocode);
    } else {
      // Customer not in aging index — show saved value as only option
      var colSel=G('m-collector');
      if(colSel) {
        colSel.innerHTML='<option value="">-- none --</option>'
          +(c.collector?'<option value="'+esc(c.collector)+'" selected>'+esc(c.collector)+'</option>':'');
        colSel.value=c.collector||'';
      }
      var cdSel=G('m-cocode');
      if(cdSel) {
        cdSel.innerHTML='<option value="">-- none --</option>'
          +(c.cocode?'<option value="'+esc(c.cocode)+'" selected>'+esc(c.cocode)+'</option>':'');
        cdSel.value=c.cocode||'';
      }
    }
  }
  G('modal-overlay').className = 'open';
}

function closeModal() {
  var ov=G('modal-overlay'); if(ov) ov.className='';
  CSS_ST.editingId=null; CSS_ST.toTags=[]; CSS_ST.ccTags=[];
}

function saveModalCustomer() {
  var custno=(G('m-custno').value||'').trim();
  var name  =(G('m-custname').value||'').trim();
  if(!custno||!name){ alert('Customer No. and Customer Name are required.'); return; }

  var c = {
    id:           CSS_ST.editingId||genId(),
    custno:       custno,
    name:         name,
    to:           CSS_ST.toTags.slice(),
    cc:           CSS_ST.ccTags.slice(),
    emailtype:    G('m-emailtype').value,
    collector:    (G('m-collector').value||'').trim(),
    endcust:      fTagsToStr('endcust'),
    bu:           fTagsToStr('bu'),
    material:     fTagsToStr('material'),
    cocode:       fTagsToStr('cocode'),
    type:         fTagsToStr('type'),
    consolidated: (G('m-consolidated').value||'').trim(),
    group:        (G('m-group')&&G('m-group').value||'').trim(),
    arma:         G('m-arma').checked,
    credit:       G('m-credit').checked,
    promised:     G('m-promised').checked,
    schedule:     G('m-schedule').value,
    schedtime:    parseInt(G('m-schedtime').value,10)||1,
    time:         '08:00'
  };

  if(CSS_ST.editingId) {
    var idx=-1;
    for(var i=0;i<CSS_ST.customers.length;i++){ if(CSS_ST.customers[i].id===CSS_ST.editingId){idx=i;break;} }
    if(idx>-1) CSS_ST.customers[idx]=c; else CSS_ST.customers.push(c);
  } else {
    CSS_ST.customers.push(c);
  }
  // Update UI immediately
  renderCustomers(); renderScheduleList(); populatePreviewDropdown();
  // Auto-save to SharePoint; keep modal open with Saving... until write confirmed
  var modalSaveBtn = G('btn-modal-save');
  pushToSharePoint('modal', modalSaveBtn, function() { closeModal(); });
}

function deleteCustomer(cid) {
  var c=null;
  for(var i=0;i<CSS_ST.customers.length;i++){ if(CSS_ST.customers[i].id===cid){c=CSS_ST.customers[i];break;} }
  if(!c) return;
  if(!confirm('Remove "'+c.name+'" from the statement list?')) return;
  CSS_ST.customers = CSS_ST.customers.filter(function(x){ return x.id!==cid; });
  renderCustomers(); renderScheduleList(); populatePreviewDropdown();
  // Auto-save to SharePoint after delete
  pushToSharePoint('delete', null, null);
}

// ── TAG INPUT ────────────────────────────────────────────
function renderTagWrap(type) {
  var wrap=G(type+'-wrap'); var inp=G(type+'-input');
  if(!wrap||!inp) return;
  var tags = type==='to'?CSS_ST.toTags:CSS_ST.ccTags;
  // Remove existing chips
  var chips=wrap.querySelectorAll('.tag-chip');
  for(var i=0;i<chips.length;i++) chips[i].parentNode.removeChild(chips[i]);
  tags.forEach(function(tag,i){
    var chip=document.createElement('span');
    chip.className='tag-chip';
    chip.setAttribute('data-idx',i);
    chip.setAttribute('data-type',type);
    chip.innerHTML=esc(tag)+' <button type="button" class="css-tag-remove" data-idx="'+i+'" data-type="'+type+'">&times;</button>';
    wrap.insertBefore(chip,inp);
  });
}

function addTag(type,val) {
  val=val.trim().replace(/,+$/,'');
  if(!val) return;
  if(type==='to') CSS_ST.toTags.push(val); else CSS_ST.ccTags.push(val);
  renderTagWrap(type);
}

function removeTag(type,idx) {
  if(type==='to') CSS_ST.toTags.splice(idx,1); else CSS_ST.ccTags.splice(idx,1);
  renderTagWrap(type);
}

// ── FILTER MULTI-TAG HELPERS ─────────────────────────────────────────────────
// Renders tag chips inside a .tag-wrap div for a filter field
function renderFilterTags(field) {
  var wrapId = 'm-'+field+'-wrap';
  var wrap = G(wrapId); if(!wrap) return;
  var tags = CSS_ST.fTags[field] || [];
  // Remove existing chips
  var existing = wrap.querySelectorAll('.tag-chip');
  for(var i=0;i<existing.length;i++) existing[i].parentNode.removeChild(existing[i]);
  // Get the input (last child)
  var inp = wrap.querySelector('input');
  tags.forEach(function(tag,i){
    var chip = document.createElement('span');
    chip.className = 'tag-chip';
    chip.setAttribute('data-fidx', i);
    chip.setAttribute('data-fld', field);
    chip.innerHTML = esc(tag)
      +' <button type="button" class="css-ftag-remove" data-fidx="'+i+'" data-fld="'+field+'">&times;</button>';
    wrap.insertBefore(chip, inp);
  });
}

function addFilterTag(field, val) {
  // Split on | or , to allow pasting multiple values at once
  var parts = val.split(/[|,]+/).map(function(v){return v.trim();}).filter(Boolean);
  parts.forEach(function(v){
    if(CSS_ST.fTags[field].indexOf(v) === -1) CSS_ST.fTags[field].push(v);
  });
  renderFilterTags(field);
}

function removeFilterTag(field, idx) {
  CSS_ST.fTags[field].splice(idx, 1);
  renderFilterTags(field);
}

// Wire a filter tag input (keydown + blur)
function wireFilterTagInput(inputId, field) {
  var inp = G(inputId); if(!inp) return;
  inp.addEventListener('keydown', function(e){
    if(e.key==='Enter' || e.key==='|' || e.key===','){
      e.preventDefault();
      if(this.value.trim()) { addFilterTag(field, this.value); this.value=''; }
    }
  });
  inp.addEventListener('blur', function(){
    if(this.value.trim()) { addFilterTag(field, this.value); this.value=''; }
  });
}

// Wire remove buttons on all filter wrap divs via delegation
function wireFilterTagWraps() {
  FILTER_FIELDS.forEach(function(field){
    var wrap = G('m-'+field+'-wrap'); if(!wrap||wrap._fWired) return;
    wrap._fWired = true;
    wrap.addEventListener('click', function(e){
      var btn = e.target;
      while(btn&&btn!==wrap){
        if(btn.classList&&btn.classList.contains('css-ftag-remove')) break;
        btn = btn.parentNode;
      }
      if(!btn||btn===wrap) return;
      removeFilterTag(btn.getAttribute('data-fld'), parseInt(btn.getAttribute('data-fidx'),10));
    });
  });
}

// Convert filter tag arrays → pipe-separated string for storage
function fTagsToStr(field) {
  return (CSS_ST.fTags[field]||[]).join('|');
}
// Parse stored pipe-separated string → tag array
function strToFTags(str){ return (str||'').split('|').map(function(v){return v.trim();}).filter(Boolean); }

// ── FILTER MULTI-TAG HELPERS ─────────────────────────────────────────────────
function renderFilterTags(field){
  var wrap=G('m-'+field+'-wrap'); if(!wrap) return;
  var tags=CSS_ST.fTags[field]||[];
  var existing=wrap.querySelectorAll('.tag-chip');
  for(var i=0;i<existing.length;i++) existing[i].parentNode.removeChild(existing[i]);
  var inp=wrap.querySelector('input');
  tags.forEach(function(tag,i){
    var chip=document.createElement('span'); chip.className='tag-chip';
    chip.innerHTML=esc(tag)+' <button type="button" class="css-ftag-remove" data-fidx="'+i+'" data-fld="'+field+'">&times;</button>';
    wrap.insertBefore(chip,inp);
  });
}
function addFilterTag(field,val){
  val.split(/[|,]+/).map(function(v){return v.trim();}).filter(Boolean).forEach(function(v){
    if(CSS_ST.fTags[field].indexOf(v)===-1) CSS_ST.fTags[field].push(v);
  });
  renderFilterTags(field);
}
function removeFilterTag(field,idx){ CSS_ST.fTags[field].splice(idx,1); renderFilterTags(field); }
function wireFilterTagInput(inputId,field){
  var inp=G(inputId); if(!inp) return;
  inp.addEventListener('keydown',function(e){
    if(e.key==='Enter'||e.key==='|'||e.key===','){e.preventDefault(); if(this.value.trim()){addFilterTag(field,this.value);this.value='';}}
  });
  inp.addEventListener('blur',function(){ if(this.value.trim()){addFilterTag(field,this.value);this.value='';} });
}
function wireFilterTagWraps(){
  FILTER_FIELDS.forEach(function(field){
    var wrap=G('m-'+field+'-wrap'); if(!wrap||wrap._fWired) return;
    wrap._fWired=true;
    wrap.addEventListener('click',function(e){
      var btn=e.target;
      while(btn&&btn!==wrap){if(btn.classList&&btn.classList.contains('css-ftag-remove')) break; btn=btn.parentNode;}
      if(!btn||btn===wrap) return;
      removeFilterTag(btn.getAttribute('data-fld'),parseInt(btn.getAttribute('data-fidx'),10));
    });
  });
}
function fTagsToStr(field){ return (CSS_ST.fTags[field]||[]).join('|'); }

function updateSchedHint() {
  var sel=G('m-schedule'); if(!sel) return;
  var lbl=G('m-schedtime-lbl'); var hint=G('sched-hint');
  var inp=G('m-schedtime'); var dayFld=G('sched-day-field');

  // Customer-specific filters always visible for all frequency values
  if(sel.value==='None'){
    if(dayFld) dayFld.style.display='none';
    if(hint)   hint.textContent='None: will not be sent automatically — manual only';
  } else if(sel.value==='Daily'){
    if(dayFld) dayFld.style.display='none';
    if(hint)   hint.textContent='Daily: sends every business day (Mon-Fri)';
  } else if(sel.value==='Weekly'){
    if(dayFld) dayFld.style.display='';
    if(lbl)    lbl.textContent='Day of week';
    if(inp){   inp.disabled=false; inp.max=7; }
    if(hint)   hint.textContent='Weekly: 1=Mon, 2=Tue, 3=Wed, 4=Thu, 5=Fri';
  } else {
    if(dayFld) dayFld.style.display='';
    if(lbl)    lbl.textContent='Day of month';
    if(inp){   inp.disabled=false; inp.max=31; }
    if(hint)   hint.textContent='Monthly: enter 1-31 for the day of month to send';
  }
}

// ── CORE SAVE — called by every action that mutates customers ────────────
// saveTrigger: 'modal'|'delete'|'manual'  — shown in status message
// modalSaveBtn: optional button to restore after save (for modal context)
// onSuccess: optional callback after confirmed write
function pushToSharePoint(saveTrigger, modalSaveBtn, onSuccess) {
  if(modalSaveBtn){ modalSaveBtn.disabled=true; modalSaveBtn.textContent='Saving...'; }
  var spBtn=G('btn-save-sp');
  if(spBtn){ spBtn.disabled=true; spBtn.textContent='Saving...'; }

  function resetBtns() {
    if(modalSaveBtn){ modalSaveBtn.disabled=false; modalSaveBtn.textContent='Save'; }
    if(spBtn){ spBtn.disabled=false; spBtn.textContent='Save to SharePoint'; }
  }

  var filePath = SP_BASE+'/'+SP_FOLDER+'/'+SETUP_FILE;

  // ── SP file operations ─────────────────────────────────────────────────────
  function spFileOp(action) {
    // action: 'UndoCheckOut' | 'CheckOut' | 'CheckIn'
    var url = spOrigin()+SP_BASE
      +"/_api/web/GetFileByServerRelativePath(decodedurl='"+filePath+"')/"+action
      +(action==='CheckIn' ? "(comment='Auto-save from Customer Statement Setup',checkintype=1)" : '');
    return getDigest().then(function(digest){
      return fetch(url, {
        method: 'POST',
        headers: { 'Accept':'application/json;odata=verbose', 'X-RequestDigest': digest,
                   'Content-Length': '0' },
        credentials: 'include'
      });
    }).catch(function(){ return {ok:false}; });
  }

  // ── Main save sequence ─────────────────────────────────────────────────────
  // 1. UndoCheckOut (silently releases any existing checkout)
  // 2. CheckOut (gets an exclusive lock so PUT succeeds cleanly)
  // 3. PUT file content
  // 4. CheckIn (releases lock — file is now available to everyone)
  function doSave() {
    var buf, ab;
    try {
      buf = buildSetupXlsx();
      ab  = new Uint8Array(buf).buffer;
    } catch(e) {
      resetBtns();
      showErr('Build error: '+e.message);
      return;
    }

    if(spBtn) spBtn.textContent='Saving…';

    // Step 1: UndoCheckOut (ignore failure — file may not be checked out)
    spFileOp('UndoCheckOut').then(function() {
      // Step 2: CheckOut
      return spFileOp('CheckOut');
    }).then(function() {
      // Step 3: PUT file content
      return spSave(SETUP_FILE, ab);
    }).then(function() {
      // Step 4: CheckIn — CRITICAL — releases the lock
      return spFileOp('CheckIn').then(function(r) {
        if(r && !r.ok) {
          // CheckIn failed — try once more after short delay
          return new Promise(function(res){ setTimeout(res, 1000); })
            .then(function(){ return spFileOp('CheckIn'); });
        }
      });
    }).then(function() {
      stamp(); resetBtns();
      var action = saveTrigger==='delete'?'removed':saveTrigger==='modal'?'updated':'saved';
      showOk('SharePoint updated — SetupCustomerSchedule.xlsx '+action+' successfully.');
      if(onSuccess) onSuccess();
    }).catch(function(err) {
      resetBtns();
      var msg = (err && err.message) || String(err);
      if(msg.indexOf('423')>-1) {
        // Still locked — force undo then retry entire sequence once
        showOk('File locked — forcing release and retrying...');
        spFileOp('UndoCheckOut').then(function(){
          return new Promise(function(res){ setTimeout(res, 2000); });
        }).then(function(){
          return spFileOp('CheckOut');
        }).then(function(){
          var buf2 = buildSetupXlsx();
          return spSave(SETUP_FILE, new Uint8Array(buf2).buffer);
        }).then(function(){
          return spFileOp('CheckIn');
        }).then(function(){
          stamp(); resetBtns();
          showOk('SharePoint updated after releasing file lock.');
          if(onSuccess) onSuccess();
        }).catch(function(err2){
          resetBtns();
          // Last resort — try UndoCheckOut so at least file is not stuck
          spFileOp('UndoCheckOut').catch(function(){});
          showErr('Save failed: '+((err2&&err2.message)||String(err2))+
            '. Please close the file in Excel and try again.');
        });
      } else {
        // Non-lock error — still try to check in / undo to not leave file stuck
        spFileOp('CheckIn').catch(function(){ spFileOp('UndoCheckOut').catch(function(){}); });
        showErr('SharePoint write failed: '+msg);
      }
    });
  }

  // Always start with UndoCheckOut to clear any stale lock, then save
  spFileOp('UndoCheckOut').then(doSave).catch(doSave);
}

// Manual save button on Preview tab (same core)
function saveToSharePoint() {
  hideAlerts();
  pushToSharePoint('manual', null, null);
}

// ── DOWNLOAD AS EXCEL ────────────────────────────────────
function downloadSetupExcel() {
  try {
    var buf  = buildSetupXlsx();
    var blob = new Blob([new Uint8Array(buf)],
      {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    var url  = URL.createObjectURL(blob);
    var a    = document.createElement('a');
    a.href   = url; a.download='SetupCustomerSchedule.xlsx';
    document.body.appendChild(a); a.click();
    document.body.removeChild(a); URL.revokeObjectURL(url);
  } catch(e) { showErr('Excel download failed: '+e.message); }
}

// ── EXPORT POWER AUTOMATE FLOW JSON ─────────────────────
// Generates a Power Automate package JSON importable via
// make.powerautomate.com → My flows → Import Package (Legacy).
// Sends FROM NVARCollection@nvidia.com via shared mailbox.
function exportPowerAutomate() {
  var spSite = spOrigin()+SP_BASE;
  var spFile = '/Shared Documents/Collection/';

  var flow = {
    "$schema": "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#",
    "contentVersion": "1.0.0.0",
    "parameters": {
      "$connections": { "defaultValue": {}, "type": "Object" },
      "$authentication": { "defaultValue": {}, "type": "SecureObject" }
    },
    "triggers": {
      "Recurrence": {
        "type": "Recurrence",
        "recurrence": {
          "frequency": "Day",
          "interval": 1,
          "schedule": { "hours": ["7"], "minutes": [0] },
          "timeZone": "Pacific Standard Time"
        }
      }
    },
    "actions": {
      "Initialize_TodayDate": {
        "type": "InitializeVariable",
        "inputs": { "variables": [{ "name": "TodayDate", "type": "string",
          "value": "@{formatDateTime(utcNow(),'yyyy-MM-dd')}" }] },
        "runAfter": {}
      },
      "Initialize_DayOfWeek": {
        "type": "InitializeVariable",
        "inputs": { "variables": [{ "name": "DayOfWeek", "type": "integer",
          "value": "@dayOfWeek(utcNow())" }] },
        "runAfter": { "Initialize_TodayDate": ["Succeeded"] }
      },
      "Initialize_DayOfMonth": {
        "type": "InitializeVariable",
        "inputs": { "variables": [{ "name": "DayOfMonth", "type": "integer",
          "value": "@dayOfMonth(utcNow())" }] },
        "runAfter": { "Initialize_DayOfWeek": ["Succeeded"] }
      },
      "Initialize_SentCount": {
        "type": "InitializeVariable",
        "inputs": { "variables": [{ "name": "SentCount", "type": "integer", "value": 0 }] },
        "runAfter": { "Initialize_DayOfMonth": ["Succeeded"] }
      },
      "Initialize_ShouldSend": {
        "type": "InitializeVariable",
        "inputs": { "variables": [{ "name": "ShouldSend", "type": "boolean", "value": false }] },
        "runAfter": { "Initialize_SentCount": ["Succeeded"] }
      },
      "Initialize_FilteredRows": {
        "type": "InitializeVariable",
        "inputs": { "variables": [{ "name": "FilteredRows", "type": "array", "value": [] }] },
        "runAfter": { "Initialize_ShouldSend": ["Succeeded"] }
      },
      "Get_Aging_Rows": {
        "type": "ApiConnection",
        "inputs": {
          "host": { "connection": { "name": "@parameters('$connections')['sharepointonline']['connectionId']" } },
          "method": "get",
          "path": "/datasets/@{encodeURIComponent(encodeURIComponent('"+spSite+"'))}/tables/@{encodeURIComponent(encodeURIComponent('"+spFile+AGING_FILE+"'))}/items",
          "queries": { "$top": 100000 }
        },
        "runAfter": { "Initialize_FilteredRows": ["Succeeded"] }
      },
      "For_Each_Customer": {
        "type": "Foreach",
        "foreach": "@variables('CustomerList')",
        "runtimeConfiguration": { "concurrency": { "repetitions": 1 } },
        "runAfter": { "Get_Aging_Rows": ["Succeeded"] },
        "actions": {
          "Check_Schedule": {
            "type": "Switch",
            "expression": "@items('For_Each_Customer')?['schedule']?['frequency']",
            "cases": {
              "Daily": { "case": "Daily", "actions": {
                "Set_Daily": { "type": "SetVariable", "inputs": { "name": "ShouldSend",
                  "value": "@and(greaterOrEquals(variables('DayOfWeek'),1),lessOrEquals(variables('DayOfWeek'),5))" }}
              }},
              "Weekly": { "case": "Weekly", "actions": {
                "Set_Weekly": { "type": "SetVariable", "inputs": { "name": "ShouldSend",
                  "value": "@equals(variables('DayOfWeek'),int(items('For_Each_Customer')?['schedule']?['day']))" }}
              }},
              "Monthly": { "case": "Monthly", "actions": {
                "Set_Monthly": { "type": "SetVariable", "inputs": { "name": "ShouldSend",
                  "value": "@equals(variables('DayOfMonth'),int(items('For_Each_Customer')?['schedule']?['day']))" }}
              }}
            },
            "default": { "actions": { "Set_Default": { "type": "SetVariable",
              "inputs": { "name": "ShouldSend", "value": false }}}},
            "runAfter": {}
          },
          "If_ShouldSend": {
            "type": "If",
            "expression": { "equals": ["@variables('ShouldSend')", true] },
            "runAfter": { "Check_Schedule": ["Succeeded"] },
            "actions": {
              "Filter_By_CustNo": {
                "type": "Query",
                "inputs": { "from": "@body('Get_Aging_Rows')?['value']",
                  "where": "@equals(string(item()?['Cust No.']),string(items('For_Each_Customer')?['custno']))" }
              },
              "Set_Filtered": { "type": "SetVariable",
                "inputs": { "name": "FilteredRows", "value": "@body('Filter_By_CustNo')" },
                "runAfter": { "Filter_By_CustNo": ["Succeeded"] }
              },
              "If_Has_Rows": {
                "type": "If",
                "expression": { "greater": ["@length(variables('FilteredRows'))", 0] },
                "runAfter": { "Set_Filtered": ["Succeeded"] },
                "actions": {
                  "Build_Table_Rows": {
                    "type": "Select",
                    "inputs": {
                      "from": "@variables('FilteredRows')",
                      "select": "@concat('<tr><td>',item()?['SO Number'],'</td><td>',item()?['Invoice'],'</td><td>',item()?['Customer PO'],'</td><td>',item()?['Inv Date'],'</td><td>',item()?['Due Date'],'</td><td>',item()?['Terms'],'</td><td style=\"text-align:right\">',item()?['Current'],'</td><td style=\"text-align:right\">',item()?['Over Due  1'],'</td><td style=\"text-align:right\">',item()?['Over Due 31'],'</td><td style=\"text-align:right\">',item()?['Over Due 61'],'</td><td style=\"text-align:right\">',item()?['Over Due 91'],'</td><td style=\"text-align:right\">',item()?['Over Due 121'],'</td><td style=\"text-align:right;font-weight:600\">',item()?['Inv Balance'],'</td><td>',item()?['Comments'],'</td></tr>')"
                    }
                  },
                  "Compose_Email_HTML": {
                    "type": "Compose",
                    "inputs": "@concat('<html><head><style>body{font-family:Arial,sans-serif;font-size:12px}table{width:100%;border-collapse:collapse}th{background:#1a56db;color:#fff;padding:7px 9px;font-size:11px;white-space:nowrap;text-align:left}th.r{text-align:right}td{padding:6px 9px;border-bottom:1px solid #e5e7eb;font-size:11px}tr:nth-child(even) td{background:#f9fafb}</style></head><body><div style=\"background:#1a56db;color:#fff;padding:14px 18px;margin-bottom:16px\"><strong style=\"font-size:15px\">Account Statement — ',items('For_Each_Customer')?['name'],'</strong><div style=\"font-size:11px;opacity:.85;margin-top:4px\">Cust No: ',items('For_Each_Customer')?['custno'],' | Date: ',variables('TodayDate'),' | Collector: ',items('For_Each_Customer')?['collector'],'</div></div><table><thead><tr><th>SO Number</th><th>Invoice</th><th>Customer PO</th><th>Inv Date</th><th>Due Date</th><th>Terms</th><th class=\"r\">Current</th><th class=\"r\">OD 1</th><th class=\"r\">OD 31</th><th class=\"r\">OD 61</th><th class=\"r\">OD 91</th><th class=\"r\">OD 121</th><th class=\"r\">Inv Balance</th><th>Comments</th></tr></thead><tbody>',join(body('Build_Table_Rows'),''),'</tbody></table><p style=\"font-size:11px;color:#6b7280;margin-top:16px\">Automated statement from NVIDIA AR Collections. Contact ',items('For_Each_Customer')?['collector'],' with queries.</p></body></html>')",
                    "runAfter": { "Build_Table_Rows": ["Succeeded"] }
                  },
                  "Send_Statement_Email": {
                    "type": "ApiConnection",
                    "inputs": {
                      "host": { "connection": { "name": "@parameters('$connections')['office365']['connectionId']" } },
                      "method": "post",
                      "path": "/v2/Mail",
                      "body": {
                        "From": "NVARCollection@nvidia.com",
                        "To": "@{join(items('For_Each_Customer')?['to'],';')}",
                        "Cc": "@{join(items('For_Each_Customer')?['cc'],';')}",
                        "Subject": "@concat('Account Statement — ',items('For_Each_Customer')?['name'],' — ',variables('TodayDate'))",
                        "Body": "@{outputs('Compose_Email_HTML')}",
                        "Importance": "Normal",
                        "IsHtml": true
                      }
                    },
                    "runAfter": { "Compose_Email_HTML": ["Succeeded"] }
                  },
                  "Increment_Sent": { "type": "IncrementVariable",
                    "inputs": { "name": "SentCount", "value": 1 },
                    "runAfter": { "Send_Statement_Email": ["Succeeded"] }
                  }
                },
                "else": { "actions": { "Log_No_Items": { "type": "Compose",
                  "inputs": "@concat('SKIP no open items: ',items('For_Each_Customer')?['name'])" }}}
              }
            },
            "else": { "actions": { "Log_Not_Today": { "type": "Compose",
              "inputs": "@concat('SKIP not scheduled today: ',items('For_Each_Customer')?['name'],' (',items('For_Each_Customer')?['schedule']?['frequency'],')')" }}}
          }
        }
      },
      "Send_Daily_Summary": {
        "type": "ApiConnection",
        "inputs": {
          "host": { "connection": { "name": "@parameters('$connections')['office365']['connectionId']" } },
          "method": "post",
          "path": "/v2/Mail",
          "body": {
            "From": "NVARCollection@nvidia.com",
            "To": "NVARCollection@nvidia.com",
            "Subject": "@concat('[AR Statements] ',variables('SentCount'),' sent — ',variables('TodayDate'))",
            "Body": "@concat('<p><strong>Daily AR Statement run complete.</strong></p><p>Date: ',variables('TodayDate'),'<br>Statements sent: <strong>',variables('SentCount'),'</strong></p><p>See Power Automate run history for full detail.</p>')",
            "IsHtml": true
          }
        },
        "runAfter": { "For_Each_Customer": ["Succeeded","Failed","Skipped"] }
      }
    }
  };

  // Embed the current customer list as a variable initialisation step
  // so the flow uses the same config as the setup page
  var custList = CSS_ST.customers.map(function(c) {
    return {
      custno: c.custno, name: c.name,
      to: c.to, cc: c.cc,
      collector: c.collector, emailtype: c.emailtype,
      schedule: { frequency: c.schedule, day: c.schedtime, time: c.time },
      filters: { arma: c.arma, credit: c.credit, promised: c.promised,
                 endcust: c.endcust, bu: c.bu, cocode: c.cocode,
                 material: c.material, type: c.type, consolidated: c.consolidated }
    };
  });
  // Insert as an Initialize_CustomerList action before For_Each_Customer
  flow.actions['Initialize_CustomerList'] = {
    "type": "InitializeVariable",
    "inputs": { "variables": [{ "name": "CustomerList", "type": "array",
      "value": custList }] },
    "runAfter": { "Initialize_FilteredRows": ["Succeeded"] }
  };
  // Fix For_Each_Customer to run after CustomerList is ready
  flow.actions['Get_Aging_Rows'].runAfter = { "Initialize_CustomerList": ["Succeeded"] };

  var json = JSON.stringify(flow, null, 2);
  var blob = new Blob([json], {type:'application/json'});
  var url  = URL.createObjectURL(blob);
  var a    = document.createElement('a');
  a.href   = url; a.download='NVARCollection_StatementFlow.json';
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
  showOk('Power Automate flow exported — import into make.powerautomate.com via My flows → Import Package (Legacy).');
}

// ── EXPORT PYTHON SCRIPT ─────────────────────────────────
function exportPython() {
  var cols = CSS_ST.selectedCols.length?CSS_ST.selectedCols:DEFAULT_COLS;
  var subj = (G('tpl-subject')||{}).value||'Account Statement -- {CustomerName} -- {Date}';
  var body = (G('tpl-body')||{}).value||'Dear {CustomerName},\n\nPlease find your statement attached.\n\nBest regards,\n{Collector}';

  var py = '#!/usr/bin/env python3\n'
    +'"""\nNVIDIA AR - Customer Statement Email Automation\nGenerated: '+new Date().toISOString()+'\n'
    +'Reads SetupCustomerSchedule.xlsx and Aging With Customer Notes.xlsx.\n"""\n\n'
    +'import pandas as pd, smtplib, os, io\n'
    +'from datetime import date\n'
    +'from email.mime.multipart import MIMEMultipart\n'
    +'from email.mime.text import MIMEText\n'
    +'from email.mime.base import MIMEBase\n'
    +'from email import encoders\n\n'
    +'SETUP_FILE   = r"SetupCustomerSchedule.xlsx"\n'
    +'AGING_FILE   = r"Aging With Customer Notes.xlsx"\n'
    +'SMTP_HOST    = "smtp.nvidia.com"\n'
    +'SMTP_PORT    = 587\n'
    +'SMTP_USER    = os.environ.get("SMTP_USER", "ar-collections@nvidia.com")\n'
    +'SMTP_PASS    = os.environ.get("SMTP_PASS", "")\n\n'
    +'STATEMENT_COLS = '+JSON.stringify(cols,null,4)+'\n\n'
    +'SUBJECT_TPL = "'+subj.replace(/"/g,'\\"')+'"\n'
    +'BODY_TPL    = """'+body+'"""\n\n'
    +'def should_send(schedule, day_val):\n'
    +'    today = date.today()\n'
    +'    if schedule == "Daily":   return today.weekday() < 5\n'
    +'    elif schedule == "Weekly": return today.weekday() == (int(day_val)-1)\n'
    +'    elif schedule == "Monthly": return today.day == int(day_val)\n'
    +'    return False\n\n'
    +'def filter_aging(df, cust):\n'
    +'    d = df.copy()\n'
    +'    if str(cust.get("Cust No.","")).strip():\n'
    +'        d = d[d["Cust No."].astype(str).str.strip() == str(cust["Cust No."]).strip()]\n'
    +'    if cust.get("Include ARMA","Y") != "Y":\n'
    +'        d = d[d.get("ARMA", pd.Series(dtype=str)).astype(str).str.upper() != "ARMA"]\n'
    +'    if cust.get("Include Credit","Y") != "Y":\n'
    +'        d = d[pd.to_numeric(d.get("Inv Balance",0), errors="coerce").fillna(0) >= 0]\n'
    +'    if cust.get("Promised Date","") == "Y":\n'
    +'        d = d[d.get("Promised Date", pd.Series(dtype=str)).astype(str).str.strip().isin(["","nan","None"])]\n'
    +'    avail = [c for c in STATEMENT_COLS if c in d.columns]\n'
    +'    return d[avail] if avail else d\n\n'
    +'def send_statement(cust, stmt_df):\n'
    +'    if stmt_df.empty:\n'
    +'        print(f"  Skip {cust[\'Customer Name\']} -- no open items"); return\n'
    +'    today_s = date.today().strftime("%Y-%m-%d")\n'
    +'    subj = SUBJECT_TPL.replace("{CustomerName}", str(cust["Customer Name"])).replace("{Date}", today_s)\n'
    +'    body = BODY_TPL.replace("{CustomerName}", str(cust["Customer Name"])).replace("{Date}", today_s).replace("{Collector}", str(cust.get("Collector Email","")))\n'
    +'    buf  = io.BytesIO(); stmt_df.to_excel(buf, index=False, engine="openpyxl"); buf.seek(0)\n'
    +'    msg  = MIMEMultipart()\n'
    +'    msg["From"]    = f"NVIDIA AR Collections <{SMTP_USER}>"\n'
    +'    msg["To"]      = str(cust["Email"])\n'
    +'    if cust.get("Collector Email"): msg["Cc"] = str(cust["Collector Email"])\n'
    +'    msg["Subject"] = subj\n'
    +'    msg.attach(MIMEText(body, "plain"))\n'
    +'    part = MIMEBase("application","octet-stream"); part.set_payload(buf.read()); encoders.encode_base64(part)\n'
    +'    part.add_header("Content-Disposition", f\'attachment; filename="Statement_{cust["Cust No."]}_{today_s}.xlsx"\')\n'
    +'    msg.attach(part)\n'
    +'    rcpt = [str(cust["Email"])] + [e.strip() for e in str(cust.get("Collector Email","")).split(";") if e.strip()]\n'
    +'    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as srv:\n'
    +'        srv.starttls(); srv.login(SMTP_USER, SMTP_PASS); srv.sendmail(SMTP_USER, rcpt, msg.as_string())\n'
    +'    print(f"  Sent -> {cust[\'Customer Name\']} -> {cust[\'Email\']}")\n\n'
    +'def main():\n'
    +'    setup_df = pd.read_excel(SETUP_FILE, dtype=str).fillna("")\n'
    +'    aging_df = pd.read_excel(AGING_FILE, sheet_name="AR_AGING", dtype=str).fillna("")\n'
    +'    print(f"Loaded {len(setup_df)} customers, {len(aging_df)} aging rows")\n'
    +'    for _, cust in setup_df.iterrows():\n'
    +'        if not should_send(cust.get("Schedule","Weekly"), cust.get("Schedule Time",1)): continue\n'
    +'        print(f"Processing: {cust[\'Customer Name\']} ({cust[\'Cust No.\']})")\n'
    +'        send_statement(cust, filter_aging(aging_df, cust))\n\n'
    +'if __name__ == "__main__":\n    main()\n';

  var blob=new Blob([py],{type:'text/plain'});
  var url=URL.createObjectURL(blob);
  var a=document.createElement('a'); a.href=url; a.download='send_statements.py';
  document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
}

// ── TAB SWITCHING ────────────────────────────────────────
var TAB_IDS=['customers','filters','schedule','preview'];
function switchTab(name) {
  TAB_IDS.forEach(function(t){
    var el=G('t-'+t); if(el) el.style.display=(t===name?'':'none');
  });
  var tabs=document.querySelectorAll('.tab');
  for(var i=0;i<tabs.length;i++){
    tabs[i].className='tab'+(tabs[i].getAttribute('data-tab')===name?' on':'');
  }
  if(name==='schedule') renderScheduleList();
  if(name==='preview'){
    populatePreviewDropdown();
    populateCollectorFilter();
    CSS_ST._checklistSel = CSS_ST._checklistSel || {};
    renderChecklist();

  }
}

// ── INIT COLUMNS ─────────────────────────────────────────
function initCols() { CSS_ST.selectedCols=DEFAULT_COLS.slice(); }

// ── WIRE DOC TYPE CHIPS ──────────────────────────────────
function wireDocTypeChips() {
  var grp=G('flt-doctypes'); if(!grp||grp._dtWired) return;
  grp._dtWired=true;
  grp.addEventListener('click', function(e) {
    var chip=e.target;
    while(chip&&chip!==grp){ if(chip.className&&chip.className.indexOf('chip')>-1) break; chip=chip.parentNode; }
    if(!chip||chip===grp) return;
    var cls=chip.className;
    chip.className=cls.indexOf(' on')>-1?cls.replace(' on',''):cls+' on';
  });
}

// ── BOOT: WIRE ALL EVENTS ────────────────────────────────
function wireEvents() {
  var tabBar=G('tab-bar'); var btnAdd=G('btn-add-cust');
  if(!tabBar||!btnAdd) return false;

  // Tab bar — delegation, no onclick attrs
  tabBar.addEventListener('click', function(e) {
    var btn=e.target;
    while(btn&&btn!==tabBar){ if(btn.getAttribute&&btn.getAttribute('data-tab')) break; btn=btn.parentNode; }
    if(!btn||btn===tabBar) return;
    var tabId=btn.getAttribute('data-tab'); if(!tabId) return;
    switchTab(tabId);
  });

  // Customers tab
  btnAdd.addEventListener('click', function(){ openModal(null); });
  var reloadBtn=G('btn-reload');
  if(reloadBtn) reloadBtn.addEventListener('click', function(e) {
    var force = e.shiftKey; // Shift+click forces re-fetch from SharePoint
    if(force) {
      // Clear all aging cache keys
      try { Object.keys(sessionStorage).forEach(function(k){ if(k.indexOf('css_aging_')===0) sessionStorage.removeItem(k); }); } catch(ex){}
      showOk('Cache cleared — reloading fresh data from SharePoint...');
    }
    loadFromSharePoint(force);
  });

  var searchInp=G('cust-search');
  if(searchInp) {
    searchInp.addEventListener('keydown', function(e){
      if(e.key==='Enter'){ CSS_ST.filterSearch=this.value; CSS_ST.custPage=0; renderCustomers(); }
    });
    // Also filter on clear (backspace to empty)
    searchInp.addEventListener('input', function(){ if(!this.value){ CSS_ST.filterSearch=''; renderCustomers(); } });
  }

  // Column filters — submit on Enter
  document.addEventListener('keydown', function(e){
    if(e.key==='Enter' && e.target && e.target.classList && e.target.classList.contains('col-filter')){
      var col = e.target.getAttribute('data-col');
      if(col){ CSS_ST.colFilters[col]=e.target.value.trim(); renderCustomers(); }
    }
  });
  var btnClearCF = G('btn-clear-col-filters');
  if(btnClearCF) btnClearCF.addEventListener('click', function(){
    CSS_ST.colFilters={};
    document.querySelectorAll('.col-filter').forEach(function(inp){ inp.value=''; });
    renderCustomers();
  });

  // Customers Excel download
  var btnCustXL = G('btn-cust-download-xl');
  if(btnCustXL) btnCustXL.addEventListener('click', downloadCustomersExcel);

  // Schedule Excel download
  var btnSchedXL = G('btn-sched-download-xl');
  if(btnSchedXL) btnSchedXL.addEventListener('click', downloadScheduleExcel);

  // Preview search
  var prevSearch = G('prev-cust-search');
  if(prevSearch) prevSearch.addEventListener('input', function(){
    CSS_ST.prevPage=0; renderChecklist();
  });

  // Pagination
  var btnPP = G('btn-prev-page'), btnNP = G('btn-next-page');
  if(btnPP) btnPP.addEventListener('click', function(){ CSS_ST.prevPage=Math.max(0,CSS_ST.prevPage-1); renderChecklist(); });
  if(btnNP) btnNP.addEventListener('click', function(){ CSS_ST.prevPage++; renderChecklist(); });

  // Modal
  var closeBtn=G('btn-modal-close'), cancelBtn=G('btn-modal-cancel'), saveBtn=G('btn-modal-save');
  var overlay=G('modal-overlay');
  if(closeBtn)  closeBtn.addEventListener('click', closeModal);
  if(cancelBtn) cancelBtn.addEventListener('click', closeModal);
  if(saveBtn)   saveBtn.addEventListener('click', saveModalCustomer);
  if(overlay)   overlay.addEventListener('click', function(e){ if(e.target===overlay) closeModal(); });

  // Tag inputs
  function wireTag(inputId, type) {
    var inp=G(inputId); if(!inp) return;
    inp.addEventListener('keydown', function(e){
      if(e.key==='Enter'||e.key===','){
        e.preventDefault(); addTag(type,this.value); this.value='';
      }
    });
    inp.addEventListener('blur', function(){ if(this.value.trim()){ addTag(type,this.value); this.value=''; }});
  }
  wireTag('to-input','to'); wireTag('cc-input','cc');

  // Tag removal — delegated on wraps
  ['to-wrap','cc-wrap'].forEach(function(wid){
    var wrap=G(wid); if(!wrap) return;
    wrap.addEventListener('click', function(e){
      var btn=e.target;
      while(btn&&btn!==wrap){ if(btn.className&&btn.className.indexOf('css-tag-remove')>-1) break; btn=btn.parentNode; }
      if(!btn||btn===wrap) return;
      removeTag(btn.getAttribute('data-type'), parseInt(btn.getAttribute('data-idx'),10));
    });
  });

  // Cust no / name auto-fill from aging index
  var custNoInp=G('m-custno');
  if(custNoInp) {
    custNoInp.addEventListener('change', onCustNoChange);
    custNoInp.addEventListener('input',  onCustNoChange);
  }
  var custNmInp=G('m-custname');
  if(custNmInp) {
    custNmInp.addEventListener('change', onCustNameChange);
    custNmInp.addEventListener('input',  onCustNameChange);
  }

  // Filter field multi-tag inputs
  wireFilterTagInput('m-endcust',    'endcust');
  wireFilterTagInput('m-bu',         'bu');
  wireFilterTagInput('m-material',   'material');
  wireFilterTagInput('m-cocode-inp', 'cocode');
  wireFilterTagInput('m-type',       'type');
  wireFilterTagWraps();

  var schedSel=G('m-schedule');
  if(schedSel) schedSel.addEventListener('change', updateSchedHint);

  // Filters tab - column buttons
  var colsAll=G('btn-cols-all'),colsNone=G('btn-cols-none'),colsDef=G('btn-cols-default');
  if(colsAll)  colsAll.addEventListener('click',  function(){ CSS_ST.selectedCols=ALL_COLS.slice(); renderColGrid(); });
  if(colsNone) colsNone.addEventListener('click', function(){ CSS_ST.selectedCols=[]; renderColGrid(); });
  if(colsDef)  colsDef.addEventListener('click',  function(){ CSS_ST.selectedCols=DEFAULT_COLS.slice(); renderColGrid(); });

  wireDocTypeChips();

  // Preview tab
  var prevDate=G('prev-date');
  if(prevDate){
    // Default to today; will be overridden by aging file's modified date once loaded
    var t=new Date();
    prevDate.value=t.getFullYear()+'-'
      +String(t.getMonth()+1).padStart(2,'0')+'-'
      +String(t.getDate()).padStart(2,'0');
    prevDate.title='Auto-set to the date Aging With Customer Notes.xlsx was last modified';
  }

  var btnPrevRefresh=G('btn-preview-refresh');
  if(btnPrevRefresh) btnPrevRefresh.addEventListener('click', renderPreview);

  // Checklist controls
  var btnPrevAll=G('btn-prev-all');
  if(btnPrevAll) btnPrevAll.addEventListener('click', function(){
    if(!CSS_ST._checklistSel) CSS_ST._checklistSel={};
    CSS_ST.customers.forEach(function(c){ CSS_ST._checklistSel[c.id]=true; });
    var chks=document.querySelectorAll('.cust-chk');
    for(var i=0;i<chks.length;i++) chks[i].checked=true;
    updateChecklistCount();
  });
  var btnPrevNone=G('btn-prev-none');
  if(btnPrevNone) btnPrevNone.addEventListener('click', function(){
    CSS_ST._checklistSel={};
    var chks=document.querySelectorAll('.cust-chk');
    for(var i=0;i<chks.length;i++) chks[i].checked=false;
    updateChecklistCount();
    G('prev-area').innerHTML='<div class="empty">Select customers above and click "Preview selected".</div>';
    txt('prev-summary','');
  });
  var btnPrevInvert=G('btn-prev-invert');
  if(btnPrevInvert) btnPrevInvert.addEventListener('click', function(){
    if(!CSS_ST._checklistSel) CSS_ST._checklistSel={};
    var chks=document.querySelectorAll('.cust-chk');
    for(var i=0;i<chks.length;i++){
      var cid=Number(chks[i].getAttribute('data-cid'));
      chks[i].checked=!chks[i].checked;
      CSS_ST._checklistSel[cid]=chks[i].checked;
    }
    updateChecklistCount();
  });

  // Collector filter
  var colFilter=G('filter-collector');
  if(colFilter) colFilter.addEventListener('change', function(){
    var el=G('cust-checklist'); if(el) el._clWired=false;
    renderChecklist();
  });
  var grpFilter=G('filter-group');
  if(grpFilter) grpFilter.addEventListener('change', function(){
    var el=G('cust-checklist'); if(el) el._clWired=false;
    renderChecklist();
  });

  var btnSendPA=G('btn-send-pa');
  if(btnSendPA) btnSendPA.addEventListener('click', runViaPowerAutomate);

  // Option B (was C): Graph auth buttons
  var btnGraphAuth=G('btn-graph-auth');
  if(btnGraphAuth) btnGraphAuth.addEventListener('click', graphConnect);
  var btnGraphDisc=G('btn-graph-disconnect');
  if(btnGraphDisc) btnGraphDisc.addEventListener('click', function(){
    GRAPH_TOKEN=null; GRAPH_TOKEN_EXP=0; GRAPH_USER={};
    var ok=G('graph-auth-ok'); if(ok) ok.style.display='none';
    var ban=G('graph-auth-banner'); if(ban) ban.style.display='';
  });

  // Apply client ID from IT setup and immediately connect
  var btnApplyClientId=G('btn-apply-client-id');
  if(btnApplyClientId) btnApplyClientId.addEventListener('click', function(){
    var inp=G('inp-client-id');
    var val=inp ? inp.value.trim() : '';
    if(!val || val.length < 30) {
      alert('Please paste a valid Azure AD Client ID (36-character GUID).');
      return;
    }
    // Update the runtime client ID — no JS file upload needed for testing
    AAD_CLIENT_ID = val;
    GRAPH_TOKEN = null; GRAPH_TOKEN_EXP = 0; // clear cached token
    showOk('Client ID set to: '+val+'. Connecting...');
    graphConnect();
  });

  // Show redirect URI in the IT instructions
  var rdUri=G('redirect-uri-display');
  if(rdUri) rdUri.textContent = window.location.protocol+'//'+window.location.host+'/_layouts/15/blank.aspx';

  // Graph auth
  // Option B (Graph/Outlook) removed

  // Send / Draft / EML buttons
  // Graph send/draft buttons removed (Option B removed)
  var btnEmlDl=G('btn-eml-download');
  if(btnEmlDl) btnEmlDl.addEventListener('click', function(){ runEmailBatch('eml'); });


  // Option C: Send directly / Save draft via Graph
  var btnSendDirect=G('btn-send-direct');
  if(btnSendDirect) btnSendDirect.addEventListener('click', function(){ sendDirectly('send'); });
  var btnDraftDirect=G('btn-save-draft-direct');
  if(btnDraftDirect) btnDraftDirect.addEventListener('click', function(){ sendDirectly('draft'); });


  // Option C: Send directly via Graph
  var btnSendDirect=G('btn-send-direct');
  if(btnSendDirect) btnSendDirect.addEventListener('click', function(){ sendDirectly('send'); });
  var btnDraftDirect=G('btn-save-draft-direct');
  if(btnDraftDirect) btnDraftDirect.addEventListener('click', function(){ sendDirectly('draft'); });

  // Save / export buttons
  var btnSaveSP=G('btn-save-sp'), btnDlXl=G('btn-download-xl');
  var btnExPa=G('btn-export-pa'), btnExPy=G('btn-export-py');
  if(btnSaveSP) btnSaveSP.addEventListener('click', saveToSharePoint);
  if(btnDlXl)   btnDlXl.addEventListener('click',   downloadSetupExcel);
  if(btnExPa)   btnExPa.addEventListener('click',   exportPowerAutomate);
  if(btnExPy)   btnExPy.addEventListener('click',   exportPython);

  return true;
}

// ── BOOT ─────────────────────────────────────────────────
function boot() {
  initCols();
  renderColGrid();
  renderCustomers();
  loadFromSharePoint();
}

// Exact same boot pattern as CashFlow.JS line 198
if(!wireEvents()) {
  document.addEventListener('DOMContentLoaded', function(){
    if(!wireEvents()) window.addEventListener('load', wireEvents);
  });
}
if(document.readyState==='loading') {
  document.addEventListener('DOMContentLoaded', boot);
} else {
  boot();
}
