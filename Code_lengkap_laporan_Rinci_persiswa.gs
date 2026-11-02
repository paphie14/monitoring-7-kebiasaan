/**
 * G7KAIH WEB API — SD ISLAM CIKAL HARAPAN II
 * Tempelkan ke Apps Script yang TERIKAT pada spreadsheet database.
 *
 * Sheet sumber: MASTER dan (opsional) Form Responses 1
 * Sheet baru otomatis: MONITORING_WEB, AKSES_ORANG_TUA, AKSES_WALI
 *
 * Deploy sebagai Web App:
 * Execute as: Me
 * Who has access: Anyone
 *
 * Jangan pernah menaruh ADMIN_KEY di HTML. Atur melalui Script Properties.
 */

const CFG = {
  MASTER: 'MASTER',
  RESPONSES_OLD: 'Form Responses 1',
  RESPONSES_WEB: 'MONITORING_WEB',
  PARENT_ACCESS: 'AKSES_ORANG_TUA',
  TEACHER_ACCESS: 'AKSES_WALI',
  TZ: 'Asia/Jakarta',
  HABITS: [
    'Bangun Pagi',
    'Beribadah',
    'Berolahraga',
    'Makan Sehat dan Bergizi',
    'Gemar Belajar',
    'Bermasyarakat',
    'Tidur Cepat'
  ],
  WEB_HEADERS: ['Tanggal','Kelas','Nama Siswa','Jenis Kelamin',
    'Bangun Pagi','Beribadah','Berolahraga','Makan Sehat dan Bergizi',
    'Gemar Belajar','Bermasyarakat','Tidur Cepat','Waktu Kirim','Diubah Oleh']
};

function doGet(e) {
  const p = (e && e.parameter) || {};
  const callback = String(p.callback || '');
  const action = String(p.action || 'ping');
  let result;
  try {
    result = routeGet_(action, p);
  } catch (err) {
    result = {ok:false, message:err.message || String(err)};
  }
  const body = JSON.stringify(result);
  if (callback && /^[a-zA-Z_$][\w$\.]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + body + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body)
    .setMimeType(ContentService.MimeType.JSON);
}


function doPost(e) {
    let result;
    
    try {
        const p = (e && e.parameter) || {};
        const action = String(p.action || '');
        const payload = p.payload ? JSON.parse(p.payload) : p;
        
        result = routePost_(action, payload);
        
    } catch (err) {
        result = {
            ok: false,
            message: err.message || String(err)
        };
    }
    
    const safe = JSON.stringify(result).replace(/</g, '\\u003c');
    
    return HtmlService.createHtmlOutput(
        '<!doctype html>' +
    '<html><head><meta charset="utf-8"></head><body>' +
    '<script>' +
    'window.top.postMessage({' +
        'type:"G7KAIH_API",' +
    'result:' + safe +
    '}, "*");' +
    '</script>' +
    '</body></html>'
    ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function routeGet_(action, p) {
  if (action === 'ping') return {ok:true, service:'G7KAIH API', version:'1.0'};
  if (action === 'initial') return {ok:true, data:getInitialData_()};
  if (action === 'parentReport') {
    const student = verifyParent_(p.code);
    return {ok:true, student:student, records:getRecords_({studentKey:key_(student.kelas,student.nama)})};
  }
  if (action === 'staffInitial') {
    const staff = verifyStaff_(p.token);
    const data = getStaffInitial_(staff);
    return {ok:true, data:data, staff:{role:staff.role,kelas:staff.kelas || ''}};
  }
  if (action === 'staffReport') {
    const staff = verifyStaff_(p.token);
    const filters = {
      period:p.period || 'month', date:p.date || '', month:p.month || '',
      semester:p.semester || '', year:p.year || '', kelas:p.kelas || '',
      tingkat:p.tingkat || '', student:p.student || ''
    };
    if (staff.role === 'teacher') filters.allowedClass = staff.kelas;
    return {ok:true, staff:{role:staff.role,kelas:staff.kelas || ''},
      report:buildReport_(filters)};
  }
  return {ok:false,message:'Aksi GET tidak dikenal.'};
}

function routePost_(action, p) {
  if (action === 'parentSubmit') return submitParent_(p);
  if (action === 'staffLogin') return loginStaff_(p);
  if (action === 'adminUpdate') {
    const staff = verifyStaff_(p.token);
    if (staff.role !== 'admin') throw new Error('Hanya admin yang boleh mengubah data.');
    return updateRecord_(p);
  }
  return {ok:false,message:'Aksi POST tidak dikenal.'};
}

function getInitialData_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.MASTER);
  if (!sh) throw new Error('Sheet MASTER tidak ditemukan.');
  const values = sh.getDataRange().getValues();
  const students = [];
  const classes = new Set();
  for (let i=1;i<values.length;i++) {
    const row=values[i], kelas=String(row[0]||'').trim().toUpperCase();
    const nama=String(row[3]||'').trim();
    if (!kelas || !nama) continue;
    classes.add(kelas);
    students.push({kelas,tingkat:Number(row[1])||parseInt(kelas,10),rombel:String(row[2]||''),nama,jk:String(row[4]||'')});
  }
  return {classes:[...classes].sort(classSort_), studentsCount:students.length, habits:CFG.HABITS};
}


function getStaffInitial_(staff) {
  const ss=SpreadsheetApp.getActive();
  const rows=ss.getSheetByName(CFG.MASTER).getDataRange().getValues();
  const classes=new Set(), students=[];
  for(let i=1;i<rows.length;i++){
    const kelas=String(rows[i][0]||'').trim().toUpperCase();
    const nama=String(rows[i][3]||'').trim();
    if(!kelas||!nama) continue;
    if(staff.role==='teacher' && kelas!==staff.kelas) continue;
    classes.add(kelas);
    students.push({kelas,nama,jk:String(rows[i][4]||'')});
  }
  return {classes:[...classes].sort(classSort_),students};
}

function setupG7KAIHWeb() {
  const ss=SpreadsheetApp.getActive();
  const master=ss.getSheetByName(CFG.MASTER);
  if (!master) throw new Error('Sheet MASTER tidak ditemukan.');
  ensureSheet_(ss,CFG.RESPONSES_WEB,CFG.WEB_HEADERS);
  const parent=ensureSheet_(ss,CFG.PARENT_ACCESS,['Kelas','Nama Siswa','Kode Akses','Aktif']);
  const teacher=ensureSheet_(ss,CFG.TEACHER_ACCESS,['Nama Wali','Kode Akses','Kelas','Aktif']);
  if (parent.getLastRow() === 1) {
    const rows=master.getDataRange().getValues().slice(1)
      .filter(r=>r[0]&&r[3])
      .map(r=>[String(r[0]).trim().toUpperCase(),String(r[3]).trim(),makeCode_(),true]);
    if (rows.length) parent.getRange(2,1,rows.length,4).setValues(rows);
  }
  [parent,teacher].forEach(s=>s.setFrozenRows(1));
  [parent,teacher].forEach(s=>s.protect().setDescription('Akses terbatas admin G7KAIH'));
  const props=PropertiesService.getScriptProperties();
  if (!props.getProperty('ADMIN_KEY')) props.setProperty('ADMIN_KEY',makeCode_()+makeCode_());
  Logger.log('Setup selesai. Kode orang tua ada di sheet AKSES_ORANG_TUA.');
  Logger.log('Atur ADMIN_KEY melalui Project Settings > Script Properties bila ingin menggantinya.');
}

function ensureSheet_(ss,name,headers) {
  let sh=ss.getSheetByName(name);
  if (!sh) sh=ss.insertSheet(name);
  if (sh.getLastRow()===0) sh.appendRow(headers);
  else if (sh.getLastRow()===1 && sh.getRange(1,1,1,headers.length).getValues()[0].join('')==='') sh.getRange(1,1,1,headers.length).setValues([headers]);
  return sh;
}

function submitParent_(p) {
  const student=verifyParent_(p.code);
  const tanggal=String(p.tanggal||'');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(tanggal)) throw new Error('Tanggal pemantauan tidak valid.');
  const answers=p.answers || [];
  if (!Array.isArray(answers) || answers.length!==7 || answers.some(x=>!['Ya','Tidak'].includes(x))) {
    throw new Error('Semua tujuh kebiasaan wajib dijawab Ya atau Tidak.');
  }
  const ss=SpreadsheetApp.getActive();
  const sh=ensureSheet_(ss,CFG.RESPONSES_WEB,CFG.WEB_HEADERS);
  const key=key_(student.kelas,student.nama);
  const data=sh.getDataRange().getValues();
  let target=0;
  for(let i=1;i<data.length;i++){
    if(String(data[i][0])===tanggal && key_(data[i][1],data[i][2])===key){target=i+1;break;}
  }
  const row=[tanggal,student.kelas,student.nama,student.jk,...answers,new Date(),'Orang Tua'];
  if(target) sh.getRange(target,1,1,row.length).setValues([row]);
  else sh.appendRow(row);
  return {ok:true,message:'Pemantauan berhasil disimpan.',tanggal,kelas:student.kelas,nama:student.nama};
}

function loginStaff_(p) {
  const role=String(p.role||'');
  const code=String(p.code||'').trim();
  if(role==='admin') {
    const expected=PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
    if(!expected || code!==expected) throw new Error('Kode admin salah.');
    return issueToken_({role:'admin'});
  }
  if(role==='teacher') {
    const sh=SpreadsheetApp.getActive().getSheetByName(CFG.TEACHER_ACCESS);
    if(!sh) throw new Error('Sheet AKSES_WALI belum tersedia. Jalankan setupG7KAIHWeb().');
    const rows=sh.getDataRange().getValues();
    for(let i=1;i<rows.length;i++){
      if(String(rows[i][1])===code && rows[i][3]!==false && rows[i][3]!=='FALSE')
        return issueToken_({role:'teacher',kelas:String(rows[i][2]).trim().toUpperCase(),nama:String(rows[i][0])});
    }
    throw new Error('Kode wali kelas tidak valid atau tidak aktif.');
  }
  throw new Error('Jenis akun tidak valid.');
}

function issueToken_(data) {
  const token=Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,'');
  CacheService.getScriptCache().put('staff_'+token,JSON.stringify(data),21600);
  return {ok:true,token,role:data.role,kelas:data.kelas||'',expiresIn:21600};
}
function verifyStaff_(token) {
  if(!token) throw new Error('Sesi staf tidak ditemukan. Silakan login kembali.');
  const raw=CacheService.getScriptCache().get('staff_'+token);
  if(!raw) throw new Error('Sesi telah berakhir. Silakan login kembali.');
  return JSON.parse(raw);
}

function verifyParent_(code) {
  code=String(code||'').trim().toUpperCase();
  if(!code) throw new Error('Masukkan kode akses orang tua.');
  const sh=SpreadsheetApp.getActive().getSheetByName(CFG.PARENT_ACCESS);
  if(!sh) throw new Error('Akses orang tua belum disiapkan. Jalankan setupG7KAIHWeb().');
  const rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(String(rows[i][2]||'').trim().toUpperCase()===code && rows[i][3]!==false && rows[i][3]!=='FALSE')
      return {kelas:String(rows[i][0]).trim().toUpperCase(),nama:String(rows[i][1]).trim(),jk:findGender_(rows[i][0],rows[i][1])};
  }
  throw new Error('Kode akses tidak valid atau sudah dinonaktifkan.');
}

function findGender_(kelas,nama) {
  const sh=SpreadsheetApp.getActive().getSheetByName(CFG.MASTER);
  const rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++) if(key_(rows[i][0],rows[i][3])===key_(kelas,nama)) return String(rows[i][4]||'');
  return '';
}

function getRecords_(filter) {
  const ss = SpreadsheetApp.getActive();
  const out = [];
  filter = filter || {};

  // Peta jenis kelamin dari MASTER
  const masterSheet = ss.getSheetByName(CFG.MASTER);
  if (!masterSheet) throw new Error('Sheet MASTER tidak ditemukan.');
  const master = masterSheet.getDataRange().getValues();
  const gender = {};

  for (let i = 1; i < master.length; i++) {
    gender[key_(master[i][0], master[i][3])] = String(master[i][4] || '');
  }

  // Ambil data monitoring dari website
  const web = ss.getSheetByName(CFG.RESPONSES_WEB);
  if (web && web.getLastRow() > 1) {
    const rows = web.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      out.push(normalizeWeb_(rows[i]));
    }
  }

  // Ambil data dari Google Form lama (jika sheet tersedia)
  const old = ss.getSheetByName(CFG.RESPONSES_OLD);
  if (old && old.getLastRow() > 1) {
    const rows = old.getDataRange().getValues();

    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      const kelas = String(r[1] || '').trim().toUpperCase();
      if (!kelas) continue;

      let nama = '';
      for (let c = 2; c <= 21; c++) {
        if (r[c]) {
          nama = String(r[c]).trim();
          break;
        }
      }
      if (!nama) continue;

      const answers = r.slice(22, 29).map(v => String(v || '').trim());
      const ts = r[0] instanceof Date ? r[0] : new Date(r[0]);
      const date = isNaN(ts)
        ? ''
        : Utilities.formatDate(ts, CFG.TZ, 'yyyy-MM-dd');

      out.push({
        tanggal: date,
        kelas: kelas,
        nama: nama,
        jk: gender[key_(kelas, nama)] || String(r[29] || ''),
        answers: answers.map(v =>
          v === 'Ya' ? 'Ya' : v === 'Tidak' ? 'Tidak' : ''
        ),
        source: 'Form lama'
      });
    }
  }

  // PENTING: jika dipanggil oleh parentReport, kembalikan hanya
  // riwayat siswa yang cocok dengan kelas + nama dari kode akses login.
  return out.filter(r => {
    if (!r.tanggal) return false;

    if (filter.studentKey) {
      return key_(r.kelas, r.nama) === filter.studentKey;
    }

    // Tanpa studentKey, fungsi tetap mengembalikan semua data
    // untuk kebutuhan laporan guru/admin.
    return true;
  });
}

function normalizeWeb_(r) {
  return {tanggal:normalizeDate_(r[0]),kelas:String(r[1]||''),nama:String(r[2]||''),jk:String(r[3]||''),answers:r.slice(4,11).map(v=>String(v||'')),source:'Website'};
}

function buildReport_(f) {
  let rows=getRecords_({});
  if(f.allowedClass) rows=rows.filter(r=>r.kelas===f.allowedClass);
  if(f.kelas) rows=rows.filter(r=>String(r.kelas).toUpperCase()===String(f.kelas).toUpperCase());
  if(f.tingkat) rows=rows.filter(r=>parseInt(r.kelas,10)===Number(f.tingkat));
  if(f.student) { const wanted=String(f.student).trim().toUpperCase(); rows=rows.filter(r=>String(r.nama).trim().toUpperCase()===wanted); }
  rows=rows.filter(r=>inPeriod_(r.tanggal,f));
  const yes=Array(7).fill(0),no=Array(7).fill(0),byStudent={};
  rows.forEach(r=>{
    const k=key_(r.kelas,r.nama);
    if(!byStudent[k]) byStudent[k]={kelas:r.kelas,nama:r.nama,jk:r.jk,total:0,yes:Array(7).fill(0),no:Array(7).fill(0),daily:{}};
    const s=byStudent[k];s.total++;
    if(!s.daily[r.tanggal])s.daily[r.tanggal]={tanggal:r.tanggal,entries:0,yes:0,no:0};
    s.daily[r.tanggal].entries++;
    r.answers.forEach((a,i)=>{if(a==='Ya'){yes[i]++;s.yes[i]++;s.daily[r.tanggal].yes++;}else if(a==='Tidak'){no[i]++;s.no[i]++;s.daily[r.tanggal].no++;}});
  });
  const expected=expectedDays_(f);
  const students=Object.values(byStudent).map(s=>{
    const answered=s.yes.reduce((a,b)=>a+b,0)+s.no.reduce((a,b)=>a+b,0), yesTotal=s.yes.reduce((a,b)=>a+b,0);
    const habits=CFG.HABITS.map((name,i)=>({name,yes:s.yes[i],no:s.no[i],percent:s.yes[i]+s.no[i]?Math.round(s.yes[i]/(s.yes[i]+s.no[i])*100):0}));
    const daily=Object.values(s.daily).sort((a,b)=>a.tanggal.localeCompare(b.tanggal)).map(d=>({...d,percent:d.yes+d.no?Math.round(d.yes/(d.yes+d.no)*100):0}));
    return {kelas:s.kelas,nama:s.nama,jk:s.jk,total:s.total,yes:s.yes,no:s.no,habits,answeredCount:answered,
      completionPercent:expected?Math.min(100,Math.round(s.total/expected*100)):0,
      percent:answered?Math.round(yesTotal/answered*100):0,daily};
  }).sort((a,b)=>a.kelas.localeCompare(b.kelas)||a.nama.localeCompare(b.nama,'id'));
  const totalAnswers=yes.reduce((a,b)=>a+b,0)+no.reduce((a,b)=>a+b,0);
  return {period:f,expectedDays:expected,entries:rows.length,students,
    habits:CFG.HABITS.map((name,i)=>({name,yes:yes[i],no:no[i],percent:yes[i]+no[i]?Math.round(yes[i]/(yes[i]+no[i])*100):0})),
    overallPercent:totalAnswers?Math.round(yes.reduce((a,b)=>a+b,0)/totalAnswers*100):0};
}
function expectedDays_(f) {
  const now=new Date(),y=Number(f.year)||Number(Utilities.formatDate(now,CFG.TZ,'yyyy'));
  if(f.period==='day')return 1;
  if(f.period==='month'){
    const ym=f.month||Utilities.formatDate(now,CFG.TZ,'yyyy-MM'),p=ym.split('-').map(Number);
    const days=new Date(p[0],p[1],0).getDate();
    const current=Utilities.formatDate(now,CFG.TZ,'yyyy-MM');
    return ym===current?Number(Utilities.formatDate(now,CFG.TZ,'d')):days;
  }
  if(f.period==='semester')return String(f.semester||'1')==='1'?184:181;
  if(f.period==='year')return new Date(y,1,29).getMonth()===1?366:365;
  return 0;
}

function inPeriod_(date,f) {
  const d=new Date(date+'T00:00:00+07:00');
  if(isNaN(d)) return false;
  if(f.period==='day') return date===(f.date||Utilities.formatDate(new Date(),CFG.TZ,'yyyy-MM-dd'));
  if(f.period==='month') return date.slice(0,7)===(f.month||Utilities.formatDate(new Date(),CFG.TZ,'yyyy-MM'));
  if(f.period==='semester') {
    const y=Number(f.year)||new Date().getFullYear(), sem=String(f.semester||'1');
    const start=sem==='1'?`${y}-07-01`:`${y}-01-01`;
    const end=sem==='1'?`${y}-12-31`:`${y}-06-30`;
    return date>=start&&date<=end;
  }
  if(f.period==='year') return date.slice(0,4)===String(f.year||new Date().getFullYear());
  return true;
}

function updateRecord_(p) {
  const row=Number(p.row);
  const sh=SpreadsheetApp.getActive().getSheetByName(CFG.RESPONSES_WEB);
  if(!row || row<2 || row>sh.getLastRow()) throw new Error('Baris data tidak ditemukan.');
  const answers=p.answers||[];
  if(answers.length!==7 || answers.some(a=>!['Ya','Tidak'].includes(a))) throw new Error('Jawaban tidak valid.');
  sh.getRange(row,5,1,7).setValues([answers]);
  sh.getRange(row,12).setValue(new Date());
  sh.getRange(row,13).setValue('Admin');
  return {ok:true,message:'Data berhasil diperbarui.'};
}

function makeCode_(){return Utilities.getUuid().replace(/-/g,'').slice(0,8).toUpperCase();}
function key_(kelas,nama){return String(kelas||'').trim().toUpperCase()+'|'+String(nama||'').trim().toUpperCase();}
function normalizeDate_(v){if(v instanceof Date)return Utilities.formatDate(v,CFG.TZ,'yyyy-MM-dd');return String(v||'').slice(0,10);}
function classSort_(a,b){const na=parseInt(a,10),nb=parseInt(b,10);return na-nb||a.localeCompare(b);}
