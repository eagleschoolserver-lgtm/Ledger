// ✅ LEDGER NILAI + LOGIN + CHALLENGE + SUPER ADMIN + RAPORT EXPORT (Ganjil/Genap) + WALI/GEMBALA KELAS
// -----------------------------------------------------------------------------
// Sheets:
//  - Users (email, name, role, salt, passwordHash, isActive, createdAt, updatedAt)
//  - TeacherAccess (teacherEmail, allowedSubjectsJson, allowedChallengesJson, isHomeroom, updatedAt, homeroomClassesJson)
//
// Per class:
//  - "Kelas X - Students"
//  - "Kelas X - Grades"      (+ kompetensiGanjil, kompetensiGenap)
//  - "Kelas X - Activities"
//  - "Kelas X - Challenges"
//  - "Kelas X - RaportMeta"  (ketidakhadiran, organisasi, prestasi, catatan wali, capaian genap)
//
// Endpoints:
//  GET  ?action=load&className=Kelas%207&token=...
//
//  POST action=login
//  POST action=saveStudent
//  POST action=saveGradesBatch
//  POST action=saveTalent
//  POST action=saveEkskulList
//  POST action=saveChallengesBatch
//
//  POST action=saveRaportMetaBatch      ✅ NEW (wali/admin)
//  POST action=getTeachers             ✅
/*  POST action=getTeacherAccess        ✅ (now includes homeroomClasses)
 *  POST action=saveTeacherAccess       ✅ (now includes isHomeroom + homeroomClasses)
 *  POST action=getGembalaSettings      ✅ (admin/super: daftar akun + gembala per kelas)
 *  POST action=saveGembalaSettings     ✅ (admin/super: assignments { "Kelas 7": email, ... })
 *  load  -> homeroomMap                ✅ { "Kelas 7": { email, name } } dipakai tanda tangan rapor
 */

// =========================
// ✅ ENGLISH LEVEL MANAGEMENT
// Sheet: EnglishLevels (scope, studentId, studentName, level, updatedAt)
// =========================
function ensureEnglishLevelsSheet_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = "EnglishLevels";
  let sh = ss.getSheetByName(name);
  if(!sh) sh = ss.insertSheet(name);
  ensureHeader_(sh, ["scope","studentId","studentName","level","updatedAt"]);
  return sh;
}

/**
 * Load English level map untuk kelas & yearKey tertentu.
 * Returns: { [studentId]: levelString }
 */
function loadEnglishLevelMap_(className, yearKey){
  const scope = buildScopeName_(yearKey, className);
  const sh = ensureEnglishLevelsSheet_();
  const vals = sh.getDataRange().getValues();
  const header = vals[0] || [];
  const rows = vals.slice(1);
  const iScope = header.indexOf("scope");
  const iSid = header.indexOf("studentId");
  const iLevel = header.indexOf("level");
  const map = {};
  rows.forEach(r=>{
    if(String(r[iScope]||"") === scope && r[iSid]){
      const sid = String(r[iSid]);
      const lvl = String(r[iLevel]||"").trim();
      if(lvl) map[sid] = lvl;
    }
  });
  return map;
}

function saveEnglishLevel_(me, className, yearKey, studentId, level){
  if(!isSuperAdmin_(me)) throw new Error("Only super_admin can assign English levels.");
  const scope = buildScopeName_(yearKey, className);
  const sh = ensureEnglishLevelsSheet_();
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  studentId = String(studentId||"").trim();
  level = String(level||"").trim();
  if(!studentId) throw new Error("studentId required");

  // Validasi level (bisa kosong = hapus)
  const allowed = ["English Advance","English Intermediate","English Beginer",""];
  if(allowed.indexOf(level) === -1) throw new Error("Level tidak valid.");

  const vals = sh.getDataRange().getValues();
  const header = vals[0] || [];
  const rows = vals.slice(1);
  const iScope = header.indexOf("scope");
  const iSid = header.indexOf("studentId");

  // Cari nama siswa dari sheet Students
  const { shS } = ensureSheets_(className, yearKey);
  const nameMap = getStudentNameMap_(shS);
  const studentName = getStudentName_(nameMap, studentId);

  for(let i=0; i<rows.length; i++){
    if(String(rows[i][iScope]||"")===scope && String(rows[i][iSid]||"")===studentId){
      if(level === ""){
        // Hapus baris
        sh.deleteRow(i+2);
      } else {
        sh.getRange(i+2, 1, 1, 5).setValues([[scope, studentId, studentName, level, now]]);
      }
      return true;
    }
  }
  if(level !== ""){
    sh.appendRow([scope, studentId, studentName, level, now]);
  }
  return true;
}

function saveEnglishLevelAction_(body, me){
  try{
    const yearKey = requireYearKey_(me, body.yearKey);
    const className = String(body.className||"").trim();
    const studentId = String(body.studentId||"").trim();
    const level = String(body.level||"").trim();
    if(!className) throw new Error("className required");
    saveEnglishLevel_(me, className, yearKey, studentId, level);
    return json_({ok:true, message:"saved"});
  }catch(err){
    return json_({ok:false, message:String(err)});
  }
}

// =========================
// ✅ JSON RESPONSE
// =========================
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/** =========================
 *  SCHOOL YEARS (Tahun Ajaran)
 *  Sheet: SchoolYears (yearKey, yearLabel, isActive, createdAt, updatedAt)
 * ========================= */
function ensureSchoolYearsSheet_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = "SchoolYears";
  let sh = ss.getSheetByName(name);
  if(!sh) sh = ss.insertSheet(name);
  ensureHeader_(sh, ["yearKey","yearLabel","isActive","createdAt","updatedAt"]);
  return sh;
}

function listSchoolYears_(me){
  const sh = ensureSchoolYearsSheet_();
  const vals = sh.getDataRange().getValues();
  const header = vals[0] || [];
  const rows = vals.slice(1);

  const idx = (k)=>header.indexOf(k);
  const iKey = idx("yearKey"), iLabel=idx("yearLabel"), iActive=idx("isActive"), iUpd=idx("updatedAt");

  const list = rows
    .filter(r => r[iKey])
    .map(r => ({
      yearKey: String(r[iKey]),
      yearLabel: String(r[iLabel] || r[iKey]),
      isActive: String(r[iActive]).toLowerCase()==="true" || r[iActive]===true || r[iActive]===1,
      updatedAt: r[iUpd] ? String(r[iUpd]) : ""
    }));

  const activeYears = list.filter(x=>x.isActive);
  if (isSuperAdmin_(me)) return { activeYears, allYears: list };
  return { activeYears };
}

function createSchoolYear_(me, yearKey, yearLabel){
  if(!isSuperAdmin_(me)) throw new Error("Only super_admin can create school years.");
  const sh = ensureSchoolYearsSheet_();

  yearKey = String(yearKey||"").trim();
  yearLabel = String(yearLabel||yearKey).trim();
  if(!yearKey) throw new Error("yearKey required");
  if(/[\\\/\?\*\[\]:]/.test(yearKey)) throw new Error("yearKey tidak boleh mengandung karakter: \\ / ? * [ ] :");

  const vals = sh.getDataRange().getValues();
  const header = vals[0] || [];
  const rows = vals.slice(1);
  const iKey = header.indexOf("yearKey");

  const exists = rows.some(r => String(r[iKey]) === yearKey);
  const now = new Date();
  const nowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  if(exists) throw new Error("Tahun ajaran sudah ada.");

  sh.appendRow([yearKey, yearLabel, false, nowStr, nowStr]);
  return true;
}

function setSchoolYearActive_(me, yearKey, isActive){
  if(!isSuperAdmin_(me)) throw new Error("Only super_admin can set active school years.");
  const sh = ensureSchoolYearsSheet_();

  const vals = sh.getDataRange().getValues();
  const header = vals[0] || [];
  const rows = vals.slice(1);
  const iKey = header.indexOf("yearKey");
  const iActive = header.indexOf("isActive");
  const iUpd = header.indexOf("updatedAt");

  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  for(let r=0;r<rows.length;r++){
    if(String(rows[r][iKey])===String(yearKey)){
      sh.getRange(r+2, iActive+1).setValue(!!isActive);
      sh.getRange(r+2, iUpd+1).setValue(nowStr);
      return true;
    }
  }
  throw new Error("Tahun ajaran tidak ditemukan.");
}

function getActiveYearKeys_(){
  const sh = ensureSchoolYearsSheet_();
  const vals = sh.getDataRange().getValues();
  const header = vals[0] || [];
  const rows = vals.slice(1);
  const iKey = header.indexOf("yearKey");
  const iActive = header.indexOf("isActive");
  return rows
    .filter(r=>r[iKey])
    .filter(r => String(r[iActive]).toLowerCase()==="true" || r[iActive]===true || r[iActive]===1)
    .map(r => String(r[iKey]));
}

function requireYearKey_(me, yearKey){
  const active = getActiveYearKeys_();
  const y = String(yearKey||"").trim();

  if(isSuperAdmin_(me)){
    // super admin boleh akses yg tidak aktif juga, tapi kalau kosong pilih yg pertama aktif jika ada
    if(!y && active.length) return active[0];
    return y || (active[0] || "DEFAULT");
  }

  // admin/guru wajib pilih tahun ajaran aktif
  if(!y) throw new Error("yearKey required (tahun ajaran belum dipilih).");
  if(active.indexOf(y)===-1) throw new Error("Tahun ajaran tidak aktif atau tidak tersedia.");
  return y;
}

/** =========================
 * TERM HELPERS (1..4)
 * ========================= */
function normalizeTermKey_(termKey){
  const t = String(termKey || "").trim();
  if (t === "1" || t === "2" || t === "3" || t === "4") return t;
  // backward-compat: semester 1/2 -> term 1/2
  if (t === "ganjil") return "1";
  if (t === "genap") return "2";
  return "1";
}
function isEvenTerm_(termKey){
  const t = normalizeTermKey_(termKey);
  return t === "2" || t === "4";
}

/** =========================
 * TERM SETTINGS (aktif/nonaktif per term)
 * Sheet: TermSettings (termKey, isActive, updatedAt)
 * ========================= */
function ensureTermSettingsSheet_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = "TermSettings";
  let sh = ss.getSheetByName(name);
  if(!sh){
    sh = ss.insertSheet(name);
    ensureHeader_(sh, ["termKey","isActive","updatedAt"]);
    // seed default: semua term aktif
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    sh.appendRow(["1", true, now]);
    sh.appendRow(["2", true, now]);
    sh.appendRow(["3", true, now]);
    sh.appendRow(["4", true, now]);
  }
  ensureHeader_(sh, ["termKey","isActive","updatedAt"]);
  return sh;
}

function getTermSettings_(){
  const sh = ensureTermSettingsSheet_();
  const vals = sh.getDataRange().getValues();
  const header = vals[0] || [];
  const rows = vals.slice(1);
  const iKey = header.indexOf("termKey");
  const iActive = header.indexOf("isActive");
  const iUpd = header.indexOf("updatedAt");

  const result = { "1": true, "2": true, "3": true, "4": true };
  rows.forEach(r => {
    const k = String(r[iKey] || "").trim();
    if(["1","2","3","4"].indexOf(k) !== -1){
      result[k] = String(r[iActive]).toLowerCase()==="true" || r[iActive]===true || r[iActive]===1;
    }
  });
  return result;
}

function setTermActive_(me, termKey, isActive){
  if(!isSuperAdmin_(me)) throw new Error("Only super_admin can change term settings.");
  const t = normalizeTermKey_(termKey);
  const sh = ensureTermSettingsSheet_();
  const vals = sh.getDataRange().getValues();
  const header = vals[0] || [];
  const rows = vals.slice(1);
  const iKey = header.indexOf("termKey");
  const iActive = header.indexOf("isActive");
  const iUpd = header.indexOf("updatedAt");
  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  for(let r = 0; r < rows.length; r++){
    if(String(rows[r][iKey]).trim() === t){
      sh.getRange(r+2, iActive+1).setValue(!!isActive);
      sh.getRange(r+2, iUpd+1).setValue(nowStr);
      return true;
    }
  }
  // not found, append
  sh.appendRow([t, !!isActive, nowStr]);
  return true;
}

function getTermSettingsAction_(me){
  try{
    const settings = getTermSettings_();
    return json_({ ok:true, termSettings: settings });
  }catch(err){
    return json_({ ok:false, message:String(err) });
  }
}

function setTermActiveAction_(body, me){
  try{
    const termKey = body.termKey;
    const isActive = body.isActive;
    setTermActive_(me, termKey, isActive);
    return json_({ ok:true });
  }catch(err){
    return json_({ ok:false, message:String(err) });
  }
}


function buildScopeName_(yearKey, className){
  const y = String(yearKey||"").trim() || "DEFAULT";
  return `${y} - ${className}`;
}


function listSchoolYearsAction_(me){
  try{
    const data = listSchoolYears_(me);
    const termSettings = getTermSettings_();
    return json_(Object.assign({ ok:true, termSettings }, data));
  }catch(err){
    return json_({ ok:false, message:String(err) });
  }
}

function createSchoolYearAction_(body, me){
  try{
    const yearKey = body.yearKey;
    const yearLabel = body.yearLabel;
    createSchoolYear_(me, yearKey, yearLabel);
    return json_({ ok:true });
  }catch(err){
    return json_({ ok:false, message:String(err) });
  }
}

function setSchoolYearActiveAction_(body, me){
  try{
    const yearKey = body.yearKey;
    const isActive = body.isActive;
    setSchoolYearActive_(me, yearKey, isActive);
    return json_({ ok:true });
  }catch(err){
    return json_({ ok:false, message:String(err) });
  }
}
// =========================
// ✅ SHA256 (HEX)
// =========================
function sha256Hex_(text) {
  const raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    text,
    Utilities.Charset.UTF_8
  );
  return raw.map(b => ("0" + (b & 0xff).toString(16)).slice(-2)).join("");
}

function normalizeEmail_(email) {
  return String(email || "").trim().toLowerCase();
}

// =========================
// ✅ TOKEN (CacheService)
// =========================
function issueToken_(email) {
  const now = new Date().getTime();
  const token = sha256Hex_(email + "|" + now + "|" + Math.random());
  CacheService.getScriptCache().put("t:" + token, email, 60 * 60 * 6); // 6 jam
  return token;
}

function getEmailFromToken_(token) {
  token = String(token || "").trim();
  if (!token) return "";
  return CacheService.getScriptCache().get("t:" + token) || "";
}

function requireAuth_(bodyOrParams) {
  const token = (bodyOrParams && bodyOrParams.token) ? bodyOrParams.token : "";
  const email = getEmailFromToken_(token);
  if (!email) throw new Error("Unauthorized: token invalid / expired. Silakan login ulang.");
  return email;
}

// =========================
// ✅ USERS SHEET (LOGIN)
// =========================
function ensureUsersSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("Users");
  if (!sh) sh = ss.insertSheet("Users");

  const headers = ["email", "name", "role", "salt", "passwordHash", "isActive", "createdAt", "updatedAt"];
  const firstRow = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  const empty = firstRow.every(x => x === "" || x === null);

  if (empty) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function findUserByEmail_(email) {
  const sh = ensureUsersSheet_();
  const e = normalizeEmail_(email);

  const last = sh.getLastRow();
  if (last < 2) return null;

  const values = sh.getRange(2, 1, last - 1, 8).getValues();
  for (let i = 0; i < values.length; i++) {
    const rowEmail = normalizeEmail_(values[i][0]);
    if (rowEmail === e) {
      return {
        rowIndex: i + 2,
        email: rowEmail,
        name: String(values[i][1] || ""),
        role: String(values[i][2] || ""),
        salt: String(values[i][3] || ""),
        passwordHash: String(values[i][4] || ""),
        isActive: String(values[i][5] || "").toUpperCase() === "TRUE",
      };
    }
  }
  return null;
}

// =========================
// ✅ TEACHER ACCESS SHEET (+ isHomeroom)
// =========================
function ensureTeacherAccessSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("TeacherAccess");
  if (!sh) sh = ss.insertSheet("TeacherAccess");

  const headers = ["teacherEmail", "allowedSubjectsJson", "allowedChallengesJson", "isHomeroom", "updatedAt", "homeroomClassesJson"];
  const firstRow = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  const empty = firstRow.every(x => x === "" || x === null);

  if (empty) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else {
    // migrate old header (4 cols) -> new (5 cols)
    if (String(firstRow[3] || "") !== "isHomeroom") {
      sh.insertColumnAfter(3); // insert col4
      sh.getRange(1, 4).setValue("isHomeroom");
      sh.getRange(1, 5).setValue("updatedAt");
    }
    // migrate -> homeroomClassesJson (col6)
    const cur = sh.getRange(1, 1, 1, 6).getValues()[0];
    if (String(cur[5] || "") !== "homeroomClassesJson") {
      sh.getRange(1, 6).setValue("homeroomClassesJson");
    }
  }
  return sh;
}

function parseHomeroomClasses_(raw) {
  let arr = [];
  try { arr = JSON.parse(raw || "[]"); } catch (e) { arr = []; }
  if (!Array.isArray(arr)) arr = [];
  return arr.map(x => String(x || "").trim()).filter(Boolean);
}

function getTeacherAccess_(teacherEmail) {
  const sh = ensureTeacherAccessSheet_();
  const t = normalizeEmail_(teacherEmail);

  const last = sh.getLastRow();
  if (last < 2) return { allowedSubjects: [], allowedChallenges: [], isHomeroom: false, homeroomClasses: [] };

  const values = sh.getRange(2, 1, last - 1, 6).getValues();
  for (let i = 0; i < values.length; i++) {
    const em = normalizeEmail_(values[i][0]);
    if (em === t) {
      let subj = [];
      let ch = [];
      try { subj = JSON.parse(values[i][1] || "[]"); } catch (e) {}
      try { ch = JSON.parse(values[i][2] || "[]"); } catch (e) {}
      if (!Array.isArray(subj)) subj = [];
      if (!Array.isArray(ch)) ch = [];
      const homeroomClasses = parseHomeroomClasses_(values[i][5]);
      const isHomeroom = String(values[i][3] || "").toUpperCase() === "TRUE" || homeroomClasses.length > 0;
      return { allowedSubjects: subj, allowedChallenges: ch, isHomeroom, homeroomClasses };
    }
  }
  return { allowedSubjects: [], allowedChallenges: [], isHomeroom: false, homeroomClasses: [] };
}

function upsertTeacherAccess_(teacherEmail, allowedSubjects, allowedChallenges, isHomeroom, homeroomClasses) {
  const sh = ensureTeacherAccessSheet_();
  const t = normalizeEmail_(teacherEmail);

  allowedSubjects = Array.isArray(allowedSubjects) ? allowedSubjects : [];
  allowedChallenges = Array.isArray(allowedChallenges) ? allowedChallenges : [];
  const hmClasses = Array.isArray(homeroomClasses) ? homeroomClasses.map(x => String(x || "").trim()).filter(Boolean) : [];
  const homeroomFlag = String(isHomeroom || "").toUpperCase() === "TRUE" || isHomeroom === true || hmClasses.length > 0;

  const last = sh.getLastRow();
  const now = new Date();

  if (last >= 2) {
    const values = sh.getRange(2, 1, last - 1, 6).getValues();
    for (let i = 0; i < values.length; i++) {
      if (normalizeEmail_(values[i][0]) === t) {
        sh.getRange(i + 2, 1, 1, 6).setValues([[
          t,
          JSON.stringify(allowedSubjects),
          JSON.stringify(allowedChallenges),
          homeroomFlag,
          now,
          JSON.stringify(hmClasses)
        ]]);
        removeClassesFromOtherTeachers_(t, hmClasses);
        return;
      }
    }
  }

  sh.appendRow([t, JSON.stringify(allowedSubjects), JSON.stringify(allowedChallenges), homeroomFlag, now, JSON.stringify(hmClasses)]);
  removeClassesFromOtherTeachers_(t, hmClasses);
}

// Kelas yang sama hanya boleh dipegang satu guru
function removeClassesFromOtherTeachers_(teacherEmail, classes) {
  if (!classes || !classes.length) return;
  const keep = {};
  classes.forEach(c => { keep[c] = true; });
  const rows = getAllTeacherAccessRows_();
  rows.forEach(r => {
    if (r.email === teacherEmail) return;
    const next = r.homeroomClasses.filter(c => !keep[c]);
    if (next.length === r.homeroomClasses.length) return;
    const sh = ensureTeacherAccessSheet_();
    sh.getRange(r.row, 6).setValue(JSON.stringify(next));
    sh.getRange(r.row, 4).setValue(next.length > 0);
  });
}

// =========================
// ✅ HOMEROOM (GEMBALA KELAS) MAP
// =========================
function getAllTeacherAccessRows_() {
  const sh = ensureTeacherAccessSheet_();
  const last = sh.getLastRow();
  const out = [];
  if (last < 2) return out;
  const values = sh.getRange(2, 1, last - 1, 6).getValues();
  for (let i = 0; i < values.length; i++) {
    const email = normalizeEmail_(values[i][0]);
    if (!email) continue;
    let subj = [], ch = [];
    try { subj = JSON.parse(values[i][1] || "[]"); } catch (e) {}
    try { ch = JSON.parse(values[i][2] || "[]"); } catch (e) {}
    if (!Array.isArray(subj)) subj = [];
    if (!Array.isArray(ch)) ch = [];
    const homeroomClasses = parseHomeroomClasses_(values[i][5]);
    const isHomeroom = String(values[i][3] || "").toUpperCase() === "TRUE" || homeroomClasses.length > 0;
    out.push({ row: i + 2, email, allowedSubjects: subj, allowedChallenges: ch, isHomeroom, homeroomClasses });
  }
  return out;
}

function userNameByEmailMap_() {
  const map = {};
  const sh = ensureUsersSheet_();
  const last = sh.getLastRow();
  if (last < 2) return map;
  const values = sh.getRange(2, 1, last - 1, 8).getValues();
  values.forEach(r => {
    const email = normalizeEmail_(r[0]);
    if (email) map[email] = String(r[1] || "");
  });
  return map;
}

function buildHomeroomMap_() {
  const names = userNameByEmailMap_();
  const out = {};
  getAllTeacherAccessRows_().forEach(row => {
    row.homeroomClasses.forEach(cls => {
      if (!out[cls]) out[cls] = { email: row.email, name: names[row.email] || "" };
    });
  });
  return out;
}

function saveHomeroomAssignments_(assignments) {
  const sh = ensureTeacherAccessSheet_();
  const rows = getAllTeacherAccessRows_();
  const byEmail = {};
  rows.forEach(r => { byEmail[r.email] = r; });

  // normalize: class -> email
  const classByEmail = {};
  Object.keys(assignments || {}).forEach(cls => {
    const em = normalizeEmail_(assignments[cls]);
    if (!em) return;
    if (!classByEmail[em]) classByEmail[em] = [];
    classByEmail[em].push(String(cls));
  });

  const now = new Date();
  const touched = {};

  rows.forEach(r => {
    const hm = classByEmail[r.email] || [];
    touched[r.email] = true;
    const prevKey = r.homeroomClasses.slice().sort().join("|");
    const nextKey = hm.slice().sort().join("|");
    // kalau penugasan kelas tidak berubah, flag lama dipertahankan (kompatibilitas)
    const flag = (prevKey === nextKey) ? !!r.isHomeroom : hm.length > 0;
    sh.getRange(r.row, 1, 1, 6).setValues([[
      r.email,
      JSON.stringify(r.allowedSubjects),
      JSON.stringify(r.allowedChallenges),
      flag,
      now,
      JSON.stringify(hm)
    ]]);
  });

  // guru/user baru yang belum punya baris di TeacherAccess
  Object.keys(classByEmail).forEach(em => {
    if (touched[em]) return;
    const hm = classByEmail[em];
    sh.appendRow([em, "[]", "[]", true, now, JSON.stringify(hm)]);
  });
}

// =========================
// ✅ LOGIN ACTION
// =========================
function login_(body) {
  const email = normalizeEmail_(body.email);
  const password = String(body.password || "").trim();

  if (!email || !password) return json_({ ok: false, message: "Email dan password wajib." });

  const user = findUserByEmail_(email);
  if (!user) return json_({ ok: false, message: "User tidak ditemukan." });
  if (!user.isActive) return json_({ ok: false, message: "User nonaktif." });

  const checkHash = sha256Hex_(user.salt + password);
  if (checkHash !== user.passwordHash) return json_({ ok: false, message: "Password salah." });

  const token = issueToken_(user.email);
  const ctxTerm = normalizeTermKey_(body.termKey || body.term || "T1");

  // attach extra flag for guru
  let teacherFlags = { isHomeroom: false };
  if (String(user.role || "").toLowerCase() === "guru") {
    const acc = getTeacherAccess_(user.email);
    teacherFlags.isHomeroom = !!acc.isHomeroom;
  }

  return json_({
    ok: true,
    termKey: ctxTerm,
    message: "Login berhasil",
    token,
    user: { email: user.email, name: user.name, role: user.role, ...teacherFlags }
  });
}

// =========================
// ✅ ROUTER
// =========================
function doGet(e) {
  try {
    const action = (e.parameter.action || "").toLowerCase();
    if (action === "load") {
      const className = (e.parameter.className || "").trim();
      const token = (e.parameter.token || "").trim();
            const yearKey = (e.parameter.yearKey || "").trim();
      const termKey = (e.parameter.termKey || "").trim();
      return loadClass_(className, token, yearKey, termKey);
    }
    return json_({ ok: false, message: "Invalid GET. Gunakan ?action=load&className=Kelas%207&token=..." });
  } catch (err) {
    return json_({ ok: false, message: "Server error (GET): " + err });
  }
}

function doPost(e) {
  try {
    let body = {};
    try {
      body = JSON.parse(e.postData.contents || "{}");
    } catch (err) {
      return json_({ ok: false, message: "Invalid JSON body" });
    }

    const action = (body.action || "").toLowerCase();

    // login does NOT require token
    if (action === "login") return login_(body);

    // Aliases
    if (action === "savechallengebatch") body.action = "saveChallengesBatch";
    if (action === "savechallengesbatch") body.action = "saveChallengesBatch";

    // All other actions require auth
    const email = requireAuth_(body);
    const me = findUserByEmail_(email);
    if (!me) return json_({ ok: false, message: "Unauthorized: user not found." });
    if (!me.isActive) return json_({ ok: false, message: "Unauthorized: user inactive." });

    // enrich flags for guru
    if (isGuru_(me)) {
      const acc = getTeacherAccess_(me.email);
      me.isHomeroom = !!acc.isHomeroom;
    } else {
      me.isHomeroom = false;
    }

    const act = String(body.action || "").toLowerCase();

    if (act === "savestudent") return saveStudent_(body, me);
    if (act === "savegradesbatch") return saveGradesBatch_(body, me);
    if (act === "savetalent") return saveTalent_(body, me);
    if (act === "saveekskullist") return saveEkskulList_(body, me);
    if (act === "savechallengesbatch") return saveChallengesBatch_(body, me);

    if (act === "saveraportmetabatch") return saveRaportMetaBatch_(body, me);

    if (act === "saveenglishlevel") return saveEnglishLevelAction_(body, me);

    if (act === "listschoolyears") return listSchoolYearsAction_(me);
    if (act === "createschoolyear") return createSchoolYearAction_(body, me);
    if (act === "setschoolyearactive") return setSchoolYearActiveAction_(body, me);

    if (act === "gettermsettings") return getTermSettingsAction_(me);
    if (act === "settermactive") return setTermActiveAction_(body, me);

    if (act === "getteachers") return getTeachers_(me);
    if (act === "getteacheraccess") return getTeacherAccessAction_(body, me);
    if (act === "saveteacheraccess") return saveTeacherAccessAction_(body, me);

    if (act === "getgembalasettings") return getGembalaUsers_(me);
    if (act === "savegembalasettings") return saveGembalaSettingsAction_(body, me);

    
    if (act === "createuser") return createUserAction_(body, me);
    if (act === "listusers") return listUsersAction_(body, me);
    if (act === "updateuser") return updateUserAction_(body, me);
    if (act === "setuserpassword") return setUserPasswordAction_(body, me);
    if (act === "deleteuser") return deleteUserAction_(body, me);

    return json_({ ok: false, message: "Invalid POST action: " + act });
  } catch (err) {
    return json_({ ok: false, message: "Server error (POST): " + err });
  }
}

// =========================
// ✅ ROLE HELPERS
// =========================
function isAdmin_(me) {
  const r = String(me.role || "").toLowerCase();
  return r === "admin" || r === "super_admin";
}
function isSuper_(me) {
  return String(me.role || "").toLowerCase() === "super_admin";
}
function isSuperAdmin_(me) {
  return isSuper_(me);
}
function isGuru_(me) {
  return String(me.role || "").toLowerCase() === "guru";
}
function isHomeroom_(me) {
  return !!me.isHomeroom;
}

// =========================
// ✅ SHEET PER CLASS
// =========================
function sheetName_(className, kind) {
  const safe = String(className || "").replace(/[\\/?*[\]]/g, "-").trim();
  return `${safe} - ${kind}`;
}

function ensureHeader_(sheet, header) {
  const firstRow = sheet.getRange(1, 1, 1, header.length).getValues()[0];
  const mismatch = header.some((h, i) => (firstRow[i] || "").toString().trim() !== h);
  if (mismatch) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.setFrozenRows(1);
  }
}

function ensureSheets_(className, yearKey) {
  if (!className) throw new Error("className required");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scope = buildScopeName_(yearKey, className);


  // Students
  const sName = sheetName_(scope, "Students");
  let shS = ss.getSheetByName(sName);
  if (!shS) shS = ss.insertSheet(sName);
  ensureHeader_(shS, ["studentId", "studentName", "nis", "nisn", "createdAt", "updatedAt"]);

  // Grades (+ kompetensi)
  const gName = sheetName_(scope, "Grades");
  let shG = ss.getSheetByName(gName);
  if (!shG) shG = ss.insertSheet(gName);
  ensureHeader_(shG, [
    "studentId", "studentName", "subject",
    "fN1","fN2","fN3","fN4","fNR",
    "pN1","pN2","pNR",
    "asesmen","sikap","final","pred",
    "kompetensiGanjil","kompetensiGenap",
    "updatedAt","term"
  ]);

  // Activities
  const aName = sheetName_(scope, "Activities");
  let shA = ss.getSheetByName(aName);
  if (!shA) shA = ss.insertSheet(aName);
  ensureHeader_(shA, ["studentId","studentName","type","name","predicate","description","updatedAt","term"]);

  // Challenges
  const cName = sheetName_(scope, "Challenges");
  let shC = ss.getSheetByName(cName);
  if (!shC) shC = ss.insertSheet(cName);
  ensureHeader_(shC, ["studentId","studentName","group","challengeName","NC","NCS","updatedAt","term"]);

  // Raport Meta
  const rName = sheetName_(scope, "RaportMeta");
  let shR = ss.getSheetByName(rName);
  if (!shR) shR = ss.insertSheet(rName);
  ensureHeader_(shR, [
    "studentId","studentName","nisn",
    "term",                  // "1" / "2" / "3" / "4"
    "sakit","izin","alpha",
    "organisasi",
    "prestasi",
    "catatanWali",
    "capaianKompetensiGenap",
    "updatedAt"
  ]);

  return { shS, shG, shA, shC, shR, sName, gName, aName, cName, rName };
}

// =========================
// ✅ Student Name Helpers
// =========================
function getStudentNameMap_(shS) {
  const values = shS.getDataRange().getValues();
  const map = new Map();
  for (let i = 1; i < values.length; i++) {
    const sid = String(values[i][0] || "").trim();
    const name = String(values[i][1] || "").trim();
    if (sid) map.set(sid, name);
  }
  return map;
}
function getStudentName_(nameMap, studentId) {
  const sid = String(studentId || "").trim();
  return nameMap.has(sid) ? nameMap.get(sid) : "";
}

// =========================
// ✅ LOAD CLASS
// =========================
function loadClass_(className, token, yearKey, termKey) {
  if (!className) return json_({ ok: false, message: "className required" });

  const email = getEmailFromToken_(token);
  if (!email) return json_({ ok:false, message:"Unauthorized: token invalid / expired. Login ulang." });

  const me = findUserByEmail_(email);
  if (!me || !me.isActive) return json_({ ok:false, message:"Unauthorized user." });

  // enrich teacher flags
  let permissions = { allowedSubjects: null, allowedChallenges: null, isHomeroom: false };
  if (isGuru_(me)) {
    const acc = getTeacherAccess_(me.email);
    permissions = { allowedSubjects: acc.allowedSubjects, allowedChallenges: acc.allowedChallenges, isHomeroom: !!acc.isHomeroom };
    me.isHomeroom = !!acc.isHomeroom;
  } else {
    me.isHomeroom = false;
  }

      const ctxYear = requireYearKey_(me, yearKey);
  const ctxTerm = normalizeTermKey_(termKey);
  const { shS, shG, shA, shC, shR, sName, gName, aName, cName, rName } = ensureSheets_(className, ctxYear);

  return json_({
    ok: true,
    className,
    me: { email: me.email, name: me.name, role: me.role, isHomeroom: !!me.isHomeroom },
    permissions,
    homeroomMap: buildHomeroomMap_(),
    sheets: {
      studentsSheet: sName,
      gradesSheet: gName,
      activitiesSheet: aName,
      challengesSheet: cName,
      raportMetaSheet: rName
    },
    students: loadStudents_(shS),
    grades: loadGrades_(shG, ctxTerm),
    activities: loadActivities_(shA, ctxTerm),
    challenges: loadChallenges_(shC, ctxTerm),
    raportMeta: loadRaportMeta_(shR, ctxTerm),
    englishLevelMap: loadEnglishLevelMap_(className, ctxYear)
  });
}

function loadStudents_(shS) {
  const values = shS.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const [studentId, studentName, nis, nisn] = values[i];
    if (!studentId) continue;
    out.push({ id: String(studentId), name: String(studentName || ""), nis: String(nis||""), nisn: String(nisn||"") });
  }
  return out;
}

function loadGrades_(shG, termKey) {
  const values = shG.getDataRange().getValues();
  const out = {};
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const sid = String(row[0] || "");
        const subject = String(row[2] || "");
    if (!sid || !subject) continue;
    const t = normalizeTermKey_(row[18] || "1");
    if (normalizeTermKey_(termKey) !== t) continue;

    if (!out[sid]) out[sid] = {};
    out[sid][subject] = {
      formatif: { N1: row[3]||0, N2: row[4]||0, N3: row[5]||0, N4: row[6]||0, NR: row[7]||0 },
      project:  { N1: row[8]||0, N2: row[9]||0, NR: row[10]||0 },
      asesmen: row[11]||0,
      sikap: row[12]||0,
      final: row[13]||0,
      pred: row[14]||"C",
      kompetensiGanjil: String(row[15] || ""),
      kompetensiGenap:  String(row[16] || "")
    };
  }
  return out;
}

function loadActivities_(shA, termKey) {
  const values = shA.getDataRange().getValues();
  const out = {};
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const sid = String(row[0] || "");
        const type = String(row[2] || "");
    if (!sid || !type) continue;
    const t = normalizeTermKey_(row[7] || "1");
    if (normalizeTermKey_(termKey) !== t) continue;

    if (!out[sid]) out[sid] = { talent:{name:"",predicate:"",description:""}, ekskul:[] };

    const name = String(row[3]||"");
    const predicate = String(row[4]||"");
    const description = String(row[5]||"");

    if (type === "talent") out[sid].talent = { name, predicate, description };
    if (type === "ekskul") out[sid].ekskul.push({ name, predicate, description });
  }
  return out;
}

function loadChallenges_(shC, termKey) {
  const values = shC.getDataRange().getValues();
  const out = {};
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const sid = String(row[0] || "").trim();
    const group = String(row[2] || "").trim();
    const name = String(row[3] || "").trim();
    if (!sid || !group || !name) continue;
    const t = normalizeTermKey_(row[7] || "1");
    if (normalizeTermKey_(termKey) !== t) continue;

    if (!out[sid]) out[sid] = {};
    if (!out[sid][group]) out[sid][group] = {};
    out[sid][group][name] = { NC: Number(row[4] || 0), NCS: Number(row[5] || 0) };
  }
  return out;
}

function loadRaportMeta_(shR, termKey) {
  const values = shR.getDataRange().getValues();
    const out = {}; // out[sid][term] = {...}
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const sid = String(row[0] || "").trim();
        const term = normalizeTermKey_(row[3] || "1");
        if (!sid) continue;
    if (normalizeTermKey_(termKey) !== term) continue;

    if (!out[sid]) out[sid] = {};
    out[sid][term] = {
      nisn: String(row[2] || ""),
      sakit: Number(row[4] || 0),
      izin:  Number(row[5] || 0),
      alpha: Number(row[6] || 0),
      organisasi: String(row[7] || ""),
      prestasi:   String(row[8] || ""),
      catatanWali: String(row[9] || ""),
      capaianKompetensiGenap: String(row[10] || "")
    };
  }
  return out;
}

// =========================
// ✅ SAVE STUDENT (now includes nis/nisn)
// =========================
function saveStudent_(body, me) {
  const yearKey = requireYearKey_(me, body.yearKey);

  if (!isAdmin_(me)) return json_({ ok:false, message:"Forbidden: hanya admin/super admin." });

  const className = String(body.className || "").trim();
  const studentId = String(body.studentId || "").trim();
  const studentName = String(body.studentName || "").trim();
  const nis = String(body.nis || "").trim();
  const nisn = String(body.nisn || "").trim();

  if (!className || !studentId || !studentName) return json_({ ok:false, message:"className, studentId, studentName required" });

  const { shS } = ensureSheets_(className, yearKey);
  const values = shS.getDataRange().getValues();
  const now = new Date();

  let found = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]||"") === studentId) { found = i + 1; break; }
  }

  if (found > 0) {
    const createdAt = shS.getRange(found, 5).getValue() || now;
    shS.getRange(found, 1, 1, 6).setValues([[studentId, studentName, nis, nisn, createdAt, now]]);
    return json_({ ok:true, message:"updated" });
  }

  shS.appendRow([studentId, studentName, nis, nisn, now, now]);
  return json_({ ok:true, message:"inserted" });
}

// =========================
// ✅ SAVE GRADES BATCH (+ kompetensi fields)
// =========================
function saveGradesBatch_(body, me) {
  const yearKey = requireYearKey_(me, body.yearKey);
  const termKey = normalizeTermKey_(body.termKey);

  const className = String(body.className || "").trim();
  const items = Array.isArray(body.items) ? body.items : [];
  if (!className || items.length === 0) return json_({ ok:false, message:"className and items[] required" });

  const { shS, shG } = ensureSheets_(className, yearKey);
  const nameMap = getStudentNameMap_(shS);

  const values = shG.getDataRange().getValues();
  const now = new Date();

  const index = new Map(); // sid|subject => rowIndex
  for (let i = 1; i < values.length; i++) {
    const sid = String(values[i][0]||"");
    const subject = String(values[i][2]||"");
    const t = normalizeTermKey_(values[i][18] || "1");
    if (sid && subject) index.set(`${sid}|${subject}|${t}`, i + 1);
  }

  let updated = 0, inserted = 0;

  items.forEach(it => {
    const sid = String(it.studentId || "").trim();
    const subject = String(it.subject || "").trim();
    if (!sid || !subject) return;

    // Optional: enforce guru subject permission (light)
    if (isGuru_(me)) {
      const acc = getTeacherAccess_(me.email);
      if (Array.isArray(acc.allowedSubjects) && acc.allowedSubjects.length > 0) {
        if (acc.allowedSubjects.indexOf(subject) === -1) return; // silently skip
      }
    }

    const studentName = getStudentName_(nameMap, sid);
    const f = it.formatif || {};
    const p = it.project || {};

    const key = `${sid}|${subject}|${termKey}`;
    const rowIdx = index.has(key) ? index.get(key) : 0;
    const prev = rowIdx ? values[rowIdx - 1] : null;

    // Capaian kompetensi per mapel tidak boleh tertimpa kosong saat payload
    // tidak membawa field-nya (mis. simpan nilai biasa / bulk paste lama).
    const kompG = (Object.prototype.hasOwnProperty.call(it, "kompetensiGanjil")
      ? String(it.kompetensiGanjil || "")
      : String((prev && prev[15]) || "")).slice(0, 250);
    const kompE = (Object.prototype.hasOwnProperty.call(it, "kompetensiGenap")
      ? String(it.kompetensiGenap || "")
      : String((prev && prev[16]) || "")).slice(0, 250);

    const row = [
      sid, studentName, subject,
      Number(f.N1)||0, Number(f.N2)||0, Number(f.N3)||0, Number(f.N4)||0, Number(f.NR)||0,
      Number(p.N1)||0, Number(p.N2)||0, Number(p.NR)||0,
      Number(it.asesmen)||0,
      Number(it.sikap)||0,
      Number(it.final)||0,
      String(it.pred || "C"),
      kompG, kompE,
      now,
      termKey
    ];

    if (index.has(key)) {
      shG.getRange(rowIdx, 1, 1, row.length).setValues([row]);
      updated++;
    } else {
      shG.appendRow(row);
      inserted++;
    }
  });

  return json_({ ok:true, message:"saved", updated, inserted });
}

// =========================
// ✅ SAVE TALENT
// =========================
function saveTalent_(body, me) {
  const yearKey = requireYearKey_(me, body.yearKey);
  const termKey = normalizeTermKey_(body.termKey);

  const className = String(body.className || "").trim();
  const studentId = String(body.studentId || "").trim();
  const talent = body.talent || null;
  if (!className || !studentId || !talent) return json_({ ok:false, message:"className, studentId, talent required" });

  const { shS, shA } = ensureSheets_(className, yearKey);
  const nameMap = getStudentNameMap_(shS);
  const studentName = getStudentName_(nameMap, studentId);

  const values = shA.getDataRange().getValues();
  const now = new Date();

  let found = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]||"") === studentId && String(values[i][2]||"") === "talent" && normalizeTermKey_(values[i][7] || "1") === termKey) {
      found = i + 1; break;
    }
  }

  const row = [
    studentId, studentName, "talent",
    String(talent.name || ""),
    String(talent.predicate || ""),
    String(talent.description || ""),
    now,
    termKey
  ];

  if (found > 0) {
    shA.getRange(found, 1, 1, row.length).setValues([row]);
    return json_({ ok:true, message:"updated" });
  }
  shA.appendRow(row);
  return json_({ ok:true, message:"inserted" });
}

// =========================
// ✅ SAVE EKSKUL LIST
// =========================
function saveEkskulList_(body, me) {
  const yearKey = requireYearKey_(me, body.yearKey);
  const termKey = normalizeTermKey_(body.termKey);

  const className = String(body.className || "").trim();
  const studentId = String(body.studentId || "").trim();
  const list = Array.isArray(body.ekskul) ? body.ekskul : [];
  if (!className || !studentId) return json_({ ok:false, message:"className, studentId required" });

  const { shS, shA } = ensureSheets_(className, yearKey);
  const nameMap = getStudentNameMap_(shS);
  const studentName = getStudentName_(nameMap, studentId);

  const values = shA.getDataRange().getValues();
  const now = new Date();

  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][0]||"") === studentId && String(values[i][2]||"") === "ekskul" && normalizeTermKey_(values[i][7] || "1") === termKey) {
      shA.deleteRow(i + 1);
    }
  }

  list.forEach(item => {
    shA.appendRow([
      studentId, studentName, "ekskul",
      String(item.name || ""),
      String(item.predicate || ""),
      String(item.description || ""),
      now,
      termKey
    ]);
  });

  return json_({ ok:true, message:"replaced", count:list.length });
}

// =========================
// ✅ SAVE CHALLENGES BATCH
// =========================
function saveChallengesBatch_(body, me) {
  const yearKey = requireYearKey_(me, body.yearKey);
  const termKey = normalizeTermKey_(body.termKey);

  const className = String(body.className || "").trim();
  const items = Array.isArray(body.items) ? body.items : [];
  if (!className || items.length === 0) return json_({ ok:false, message:"className and items[] required" });

  const { shS, shC } = ensureSheets_(className, yearKey);
  const nameMap = getStudentNameMap_(shS);

  const values = shC.getDataRange().getValues();
  const now = new Date();

  const index = new Map(); // sid|group|name => rowIndex
  for (let i = 1; i < values.length; i++) {
    const sid = String(values[i][0]||"").trim();
    const group = String(values[i][2]||"").trim();
    const name = String(values[i][3]||"").trim();
    const t = normalizeTermKey_(values[i][7] || "1");
    if (sid && group && name) index.set(`${sid}|${group}|${name}|${t}`, i + 1);
  }

  let updated = 0, inserted = 0;

  items.forEach(it => {
    const sid = String(it.studentId || "").trim();
    const group = String(it.group || "").trim();
    const name = String(it.challengeName || "").trim();
    if (!sid || !group || !name) return;

    const studentName = getStudentName_(nameMap, sid);

    const row = [ sid, studentName, group, name, Number(it.NC)||0, Number(it.NCS)||0, now, termKey ];

    const key = `${sid}|${group}|${name}|${termKey}`;
    if (index.has(key)) {
      shC.getRange(index.get(key), 1, 1, row.length).setValues([row]);
      updated++;
    } else {
      shC.appendRow(row);
      inserted++;
    }
  });

  return json_({ ok:true, message:"saved", updated, inserted });
}

// =========================
// ✅ RAPORT META (wali/admin)
// =========================
function saveRaportMetaBatch_(body, me) {
  const yearKey = requireYearKey_(me, body.yearKey);
  const termKey = normalizeTermKey_(body.termKey);

  const className = String(body.className || "").trim();
  const items = Array.isArray(body.items) ? body.items : [];
  if (!className || items.length === 0) return json_({ ok:false, message:"className and items[] required" });

  // only admin/super or homeroom teacher
  if (!isAdmin_(me) && !isHomeroom_(me)) {
    return json_({ ok:false, message:"Forbidden: hanya admin/super admin atau wali/gembala kelas." });
  }

  const { shS, shR } = ensureSheets_(className, yearKey);
  const nameMap = getStudentNameMap_(shS);

  const values = shR.getDataRange().getValues();
  const now = new Date();

  const index = new Map(); // sid|term => rowIndex
  for (let i = 1; i < values.length; i++) {
    const sid = String(values[i][0]||"").trim();
    const t = normalizeTermKey_(values[i][3] || "1");
    if (sid) index.set(`${sid}|${t}`, i + 1);
  }

  let updated = 0, inserted = 0;

  items.forEach(it => {
    const sid = String(it.studentId || "").trim();
    const t = termKey;
    if (!sid) return;

    const studentName = getStudentName_(nameMap, sid);

    const row = [
      sid, studentName,
      String(it.nisn || ""),
      t,
      Number(it.sakit)||0,
      Number(it.izin)||0,
      Number(it.alpha)||0,
      String(it.organisasi || ""),
      String(it.prestasi || ""),
      String(it.catatanWali || ""),
      String(it.capaianKompetensiGenap || ""),
      now
    ];

    const key = `${sid}|${t}`;
    if (index.has(key)) {
      shR.getRange(index.get(key), 1, 1, row.length).setValues([row]);
      updated++;
    } else {
      shR.appendRow(row);
      inserted++;
    }
  });

  return json_({ ok:true, message:"saved", updated, inserted });
}

// =========================
// ✅ SUPER ADMIN ACTIONS
// =========================
function getTeachers_(me) {
  if (!isSuper_(me)) return json_({ ok:false, message:"Forbidden: hanya super admin." });

  const sh = ensureUsersSheet_();
  const last = sh.getLastRow();
  const out = [];
  if (last >= 2) {
    const values = sh.getRange(2, 1, last - 1, 8).getValues();
    values.forEach(r => {
      const email = normalizeEmail_(r[0]);
      const name = String(r[1] || "");
      const role = String(r[2] || "").toLowerCase();
      const isActive = String(r[5] || "").toUpperCase() === "TRUE";
      if (email && role === "guru" && isActive) out.push({ email, name, role: "guru" });
    });
  }
  return json_({ ok:true, teachers: out });
}

function getTeacherAccessAction_(body, me) {
  if (!isAdmin_(me)) return json_({ ok:false, message:"Forbidden: hanya admin/super admin." });

  const teacherEmail = normalizeEmail_(body.teacherEmail);
  if (!teacherEmail) return json_({ ok:false, message:"teacherEmail required" });

  const acc = getTeacherAccess_(teacherEmail);
  return json_({ ok:true, allowedSubjects: acc.allowedSubjects, allowedChallenges: acc.allowedChallenges, isHomeroom: !!acc.isHomeroom, homeroomClasses: acc.homeroomClasses });
}

function saveTeacherAccessAction_(body, me) {
  if (!isAdmin_(me)) return json_({ ok:false, message:"Forbidden: hanya admin/super admin." });

  const teacherEmail = normalizeEmail_(body.teacherEmail);
  if (!teacherEmail) return json_({ ok:false, message:"teacherEmail required" });

  upsertTeacherAccess_(teacherEmail, body.allowedSubjects, body.allowedChallenges, body.isHomeroom, body.homeroomClasses);
  return json_({ ok:true, message:"saved", homeroomMap: buildHomeroomMap_() });
}

// =========================
// ✅ PENGATURAN GEMBALA KELAS (admin / super admin)
// =========================
function getGembalaUsers_(me) {
  if (!isAdmin_(me)) return json_({ ok:false, message:"Forbidden: hanya admin/super admin." });
  const sh = ensureUsersSheet_();
  const last = sh.getLastRow();
  const out = [];
  if (last >= 2) {
    const values = sh.getRange(2, 1, last - 1, 8).getValues();
    values.forEach(r => {
      const email = normalizeEmail_(r[0]);
      const name = String(r[1] || "");
      const role = String(r[2] || "").toLowerCase();
      const isActive = String(r[5] || "").toUpperCase() === "TRUE";
      if (email && name && isActive && ["guru", "admin", "super_admin"].indexOf(role) >= 0) {
        out.push({ email, name, role });
      }
    });
  }
  out.sort((a, b) => a.name.localeCompare(b.name));
  return json_({ ok:true, users: out, homeroomMap: buildHomeroomMap_() });
}

function saveGembalaSettingsAction_(body, me) {
  if (!isAdmin_(me)) return json_({ ok:false, message:"Forbidden: hanya admin/super admin." });
  const assignments = body.assignments && typeof body.assignments === "object" ? body.assignments : {};
  saveHomeroomAssignments_(assignments);
  return json_({ ok:true, message:"Pengaturan gembala kelas tersimpan.", homeroomMap: buildHomeroomMap_() });
}

// =========================
// ✅ CREATE USER MANUAL (RUN FROM EDITOR)
// =========================

function createUserAction_(body, me){
  if (!isSuper_(me)) return json_({ ok:false, message:"Forbidden: only super_admin." });

  const email = normalizeEmail_(String(body.email || ""));
  const name  = String(body.name || "").trim();
  const role  = String(body.role || "").trim().toLowerCase(); // guru/admin/super_admin
  const password = String(body.password || "");

  if (!email) return json_({ ok:false, message:"Email wajib diisi." });
  if (!name)  return json_({ ok:false, message:"Nama wajib diisi." });
  if (!password || password.length < 4) return json_({ ok:false, message:"Password minimal 4 karakter." });

  const allowed = ["guru","admin","super_admin"];
  if (allowed.indexOf(role) === -1) return json_({ ok:false, message:"Role tidak valid. Gunakan: guru/admin/super_admin." });

  const sh = ensureUsersSheet_();
  const found = findUserByEmail_(email);
  const now = new Date();

  if (found && !body.forceUpdate) {
    return json_({ ok:false, message:"User sudah ada. Centang 'Update jika sudah ada' untuk menimpa." });
  }

  const salt = Utilities.getUuid();
  const passwordHash = sha256Hex_(salt + password);

  if (!found){
    sh.appendRow([email, name, role, salt, passwordHash, true, now, now]);
    return json_({ ok:true, message:"Akun berhasil dibuat.", email: email, role: role });
  }

  // update existing row
  const data = sh.getDataRange().getValues();
  // header assumed in row 1
  let rowIndex = -1;
  for (let r=1; r<data.length; r++){
    if (normalizeEmail_(data[r][0]) === email){ rowIndex = r+1; break; }
  }
  if (rowIndex === -1) {
    // fallback: append
    sh.appendRow([email, name, role, salt, passwordHash, true, now, now]);
    return json_({ ok:true, message:"Akun dibuat (fallback).", email: email, role: role });
  }

  sh.getRange(rowIndex, 2).setValue(name);
  sh.getRange(rowIndex, 3).setValue(role);
  sh.getRange(rowIndex, 4).setValue(salt);
  sh.getRange(rowIndex, 5).setValue(passwordHash);
  sh.getRange(rowIndex, 6).setValue(true);
  sh.getRange(rowIndex, 8).setValue(now);

  return json_({ ok:true, message:"Akun berhasil diupdate.", email: email, role: role });
}




// =========================
// ✅ SUPER ADMIN: USER MANAGEMENT
// =========================
function listUsersAction_(body, me){
  if (!isSuper_(me)) return json_({ ok:false, message:"Forbidden: super_admin only." });
  const sh = ensureUsersSheet_();
  const last = sh.getLastRow();
  if (last < 2) return json_({ ok:true, users: [] });

  const values = sh.getRange(2, 1, last - 1, 8).getValues();
  const users = values.map(r => ({
    email: normalizeEmail_(r[0]),
    name: String(r[1] || ""),
    role: String(r[2] || ""),
    isActive: String(r[5] || "").toUpperCase() === "TRUE",
    createdAt: r[6] || "",
    updatedAt: r[7] || ""
  }));
  return json_({ ok:true, users: users });
}

function updateUserAction_(body, me){
  if (!isSuper_(me)) return json_({ ok:false, message:"Forbidden: super_admin only." });
  const email = normalizeEmail_(body.email);
  if (!email) return json_({ ok:false, message:"email required" });

  // prevent locking yourself out accidentally
  const meEmail = normalizeEmail_(me.email || "");
  const sh = ensureUsersSheet_();
  const u = findUserByEmail_(email);
  if (!u) return json_({ ok:false, message:"User not found" });

  const name = (body.name !== undefined) ? String(body.name || "") : null;
  const role = (body.role !== undefined) ? String(body.role || "").toLowerCase().trim() : null;
  const isActive = (body.isActive !== undefined) ? !!body.isActive : null;

  if (role !== null && ["guru","admin","super_admin"].indexOf(role) === -1){
    return json_({ ok:false, message:"Invalid role" });
  }
  if (email === meEmail){
    if (isActive === false) return json_({ ok:false, message:"Tidak boleh menonaktifkan akun sendiri." });
    if (role !== null && role !== "super_admin") return json_({ ok:false, message:"Tidak boleh menurunkan role akun sendiri." });
  }

  if (name !== null) sh.getRange(u.rowIndex, 2).setValue(name);
  if (role !== null) sh.getRange(u.rowIndex, 3).setValue(role);
  if (isActive !== null) sh.getRange(u.rowIndex, 6).setValue(isActive ? "TRUE" : "FALSE");
  sh.getRange(u.rowIndex, 8).setValue(new Date());

  return json_({ ok:true, message:"User updated", email: email });
}

function setUserPasswordAction_(body, me){
  if (!isSuper_(me)) return json_({ ok:false, message:"Forbidden: super_admin only." });
  const email = normalizeEmail_(body.email);
  const password = String(body.password || "").trim();
  if (!email) return json_({ ok:false, message:"email required" });
  if (!password || password.length < 4) return json_({ ok:false, message:"Password minimal 4 karakter." });

  const u = findUserByEmail_(email);
  if (!u) return json_({ ok:false, message:"User not found" });

  const salt = sha256Hex_(email + "|" + new Date().getTime() + "|" + Math.random()).slice(0, 16);
  const hash = sha256Hex_(salt + "|" + password);

  const sh = ensureUsersSheet_();
  sh.getRange(u.rowIndex, 4).setValue(salt);
  sh.getRange(u.rowIndex, 5).setValue(hash);
  sh.getRange(u.rowIndex, 8).setValue(new Date());

  return json_({ ok:true, message:"Password updated", email: email });
}

function deleteUserAction_(body, me){
  if (!isSuper_(me)) return json_({ ok:false, message:"Forbidden: super_admin only." });
  const email = normalizeEmail_(body.email);
  if (!email) return json_({ ok:false, message:"email required" });

  const meEmail = normalizeEmail_(me.email || "");
  if (email === meEmail) return json_({ ok:false, message:"Tidak boleh menghapus akun sendiri." });

  const u = findUserByEmail_(email);
  if (!u) return json_({ ok:false, message:"User not found" });

  const hard = String(body.hard || "").toLowerCase() === "true";
  const sh = ensureUsersSheet_();

  if (hard){
    sh.deleteRow(u.rowIndex);
    return json_({ ok:true, message:"User deleted (hard)", email: email });
  }

  // soft delete: deactivate
  sh.getRange(u.rowIndex, 6).setValue("FALSE");
  sh.getRange(u.rowIndex, 8).setValue(new Date());
  return json_({ ok:true, message:"User deactivated", email: email });
}

function createUserManual() {
  const email = "samuel@gmail.com";     // ganti
  const name  = "jovelyn";             // ganti
  const role  = "admin";                // guru / admin / super_admin
  const password = "123456";           // ganti password

  const sh = ensureUsersSheet_();
  const e = normalizeEmail_(email);

  const salt = Utilities.getUuid();
  const passwordHash = sha256Hex_(salt + password);
  const now = new Date();

  sh.appendRow([e, name, role, salt, passwordHash, true, now, now]);

  Logger.log("User dibuat: " + e);
  Logger.log("Password: " + password);
}
