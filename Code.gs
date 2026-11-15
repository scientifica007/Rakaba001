/**
 * Rakaba001 — سجل الحصص وغيابات المتكونين.
 * هذا المشروع مخصص لبرنامج Apps Script المرتبط بقاعدة البيانات في Google Sheets.
 * لا تضع رموز دخول المستخدمين أو بيانات المتكونين في مستودع GitHub.
 */
const TZ = 'Africa/Algiers';
const SHEETS = ['Config','Groups','Students','Schedule','Sessions','Attendance','Users','Audit','Messages','WeeklyOverrides'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('رقابة 001 — متابعة الحصص والغيابات')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function now_() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss"); }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function uuid_(prefix) { return prefix + '-' + Utilities.getUuid(); }
function clean_(v) { return String(v == null ? '' : v).trim(); }
function fail_(s) { throw new Error(s); }
function assert_(ok,s) { if (!ok) fail_(s); }
function db_() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  assert_(ss, 'يجب إنشاء مشروع Apps Script من داخل قاعدة بيانات Google Sheets.');
  SHEETS.forEach(n=>assert_(ss.getSheetByName(n), 'ورقة قاعدة البيانات غير موجودة: '+n));
  return ss;
}
function rows_(name) {
  const a=db_().getSheetByName(name).getDataRange().getValues();
  const keys=a.shift().map(clean_);
  return a.filter(r=>r.some(x=>x!==''&&x!==null)).map((r,i)=>{
    const obj={_row:i+2};
    keys.forEach((key,j)=>obj[key]=r[j] instanceof Date?Utilities.formatDate(r[j],TZ,key==='start'?'HH:mm':'yyyy-MM-dd'):r[j]);
    return obj;
  });
}
function append_(name,records) {
  if (!records.length) return;
  const s=db_().getSheetByName(name);
  const head=s.getRange(1,1,1,s.getLastColumn()).getValues()[0].map(clean_);
  const values=records.map(rec=>head.map(k=>Object.prototype.hasOwnProperty.call(rec,k)?rec[k]:''));
  s.getRange(s.getLastRow()+1,1,values.length,head.length).setValues(values);
}
function audit_(user,action,entity,detail) {
  append_('Audit',[{at:now_(),actor:user.username,action:action,entity:entity,detail:JSON.stringify(detail).slice(0,5000)}]);
}
function secret_() {
  let p=PropertiesService.getScriptProperties().getProperty('PRIVATE_PEPPER');
  if (!p) { p=Utilities.getUuid()+Utilities.getUuid(); PropertiesService.getScriptProperties().setProperty('PRIVATE_PEPPER',p); }
  return p;
}
function hash_(pin) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(String(pin),secret_())); }
function pin_() { return (Utilities.getUuid()+Utilities.getUuid()).replace(/-/g,'').slice(0,12).toUpperCase(); }

/** ينفذ يدويًا من محرر Apps Script مرة واحدة فقط؛ اقرأ بيانات الدخول في Execution log. */
function setupSystem() {
  db_();
  assert_(rows_('Users').length===0,'الحساب الرئيسي موجود. لن يُنشأ حساب رئيسي ثانٍ.');
  const password=pin_();
  append_('Users',[{id:uuid_('USR'),username:'owner',hash:hash_(password),role:'admin',groups:'*',enabled:true,created_at:now_()}]);
  Logger.log('RAKABA001 ADMIN USER: owner');
  Logger.log('RAKABA001 ONE-TIME ACCESS CODE: '+password);
  Logger.log('انسخ الرمز الآن واحفظه في مكان آمن. لا تضعه في GitHub أو تشاركه.');
  return 'تم إنشاء الحساب الرئيسي. راجع سجل التنفيذ لمعرفة رمز الدخول.';
}
function login_(p) {
  const username=clean_(p.username).toLowerCase(), password=clean_(p.password);
  assert_(/^[a-z0-9._-]{3,30}$/.test(username),'اسم المستخدم غير صالح.');
  const cache=CacheService.getScriptCache(),key='fail_'+username;
  assert_(Number(cache.get(key)||0)<8,'محاولات كثيرة. أعد المحاولة بعد عشر دقائق.');
  const user=rows_('Users').find(x=>clean_(x.username).toLowerCase()===username && x.enabled===true);
  if (!user||hash_(password)!==clean_(user.hash)) {
    cache.put(key,String(Number(cache.get(key)||0)+1),600);
    fail_('اسم المستخدم أو رمز الدخول غير صحيح.');
  }
  cache.remove(key);
  const token=(Utilities.getUuid()+Utilities.getUuid()).replace(/-/g,'');
  cache.put('auth_'+token,clean_(user.id),21600);
  return {token:token,user:{username:user.username,role:user.role,groups:user.groups}};
}
function auth_(token) {
  const t=clean_(token);
  assert_(/^[a-f0-9]{64}$/i.test(t),'الجلسة غير صالحة. أعد تسجيل الدخول.');
  const id=CacheService.getScriptCache().get('auth_'+t);
  const user=rows_('Users').find(u=>u.id===id && u.enabled===true);
  assert_(user,'انتهت الجلسة أو أُلغيت. أعد تسجيل الدخول.');
  return user;
}
function can_(user,groupId) {
  return user.role==='admin'||clean_(user.groups)==='*'||clean_(user.groups).split(',').includes(groupId);
}
function admin_(u) { assert_(u.role==='admin','هذه العملية مخصصة للإدارة.'); }
function worker_(u) { assert_(u.role==='admin'||u.role==='monitor','هذه العملية مخصصة للمراقب أو الإدارة.'); }
function group_(id) { const g=rows_('Groups').find(x=>x.id===id&&x.active===true);assert_(g,'الفوج غير موجود أو غير مفعل.');return g; }
function day_(v) {
  const x=clean_(v); assert_(/^\d{4}-\d{2}-\d{2}$/.test(x),'التاريخ غير صالح.');
  const d=new Date(x+'T12:00:00Z');
  assert_(!isNaN(d.getTime())&&d.toISOString().slice(0,10)===x,'التاريخ غير صالح.');
  return x;
}
function weekday_(date) { return new Date(date+'T12:00:00Z').getUTCDay(); }
function inDate_(s,date) {
  return Number(s.weekday)===weekday_(date) && (!s.valid_from||clean_(s.valid_from)<=date) && (!s.valid_to||clean_(s.valid_to)>=date);
}
function eligible_(student,date) {
  return student.status!=='REVIEW' && (
    student.status==='ACTIVE' || (student.status==='INACTIVE'&&student.end_date)
  ) && (!student.start_date||clean_(student.start_date)<=date)
    && (!student.end_date||clean_(student.end_date)>=date);
}
function latest_(all) {
  const obj={};
  all.forEach(s=>{ if(!obj[s.session_key]||Number(obj[s.session_key].version)<Number(s.version))obj[s.session_key]=s; });
  return obj;
}
function context_(u,p) {
  const from=day_(p.from||today_()),to=day_(p.to||today_());
  const delta=(new Date(to+'T12:00:00Z')-new Date(from+'T12:00:00Z'))/86400000;
  assert_(delta>=0&&delta<=45,'حدد نطاقًا لا يتجاوز 46 يومًا.');
  const groups=rows_('Groups').filter(g=>can_(u,g.id)&&g.active===true);
  const gids=groups.map(g=>g.id);
  const students=rows_('Students').filter(s=>gids.includes(s.group_id));
  const schedule=rows_('Schedule').filter(s=>gids.includes(s.group_id));
  const all=rows_('Sessions').filter(s=>gids.includes(s.group_id)&&s.date>=from&&s.date<=to);
  const sessions=Object.values(latest_(all));
  const keys=new Set(sessions.map(s=>s.session_key+'|'+s.version));
  const attendance=rows_('Attendance').filter(a=>keys.has(a.session_key+'|'+a.version));
  const messages=rows_('Messages').filter(m=>gids.includes(m.group_id)).slice(-100);
  const overrides=rows_('WeeklyOverrides').filter(x=>gids.includes(x.group_id)&&x.week_start>=from&&x.week_start<=to);
  return {user:{username:u.username,role:u.role,groups:u.groups},from:from,to:to,groups:groups,
    students:students,schedule:schedule,sessions:sessions,attendance:attendance,
    messages:messages,overrides:overrides};
}
function withLock_(f) {
  const lock=LockService.getScriptLock();
  lock.waitLock(25000);
  try { return f(); } finally { lock.releaseLock(); }
}
function record_(u,p,skipConflict) {
  const groupId=clean_(p.groupId),date=day_(p.date),slotId=clean_(p.slotId);
  assert_(can_(u,groupId),'لا تملك صلاحية هذا الفوج.');
  group_(groupId);
  const slot=rows_('Schedule').find(s=>s.id===slotId&&s.group_id===groupId&&inDate_(s,date));
  assert_(slot,'الحصة غير موجودة في جدول هذا اليوم.');
  const status=clean_(p.status);
  assert_(status==='HELD'||status==='CANCELLED','حالة الحصة غير صالحة.');
  const reason=clean_(p.reason),note=clean_(p.note).slice(0,500);
  if(status==='CANCELLED')assert_(reason.length>=3,'أدخل سبب عدم إنجاز الحصة.');
  const key=groupId+'|'+date+'|'+slotId;
  const old=latest_(rows_('Sessions'))[key];
  const ver=old?Number(old.version):0;
  if(!skipConflict)assert_(Number(p.expectedVersion||0)===ver,'سجّل مستخدم آخر تعديلًا على الحصة. حدّث البيانات قبل إعادة الحفظ.');
  const requested=Array.isArray(p.absences)?p.absences:[];
  const users=rows_('Students').filter(s=>s.group_id===groupId&&eligible_(s,date));
  const allowed=new Set(users.map(s=>s.id));
  const seen=new Set(),abs=[];
  if(status==='HELD')requested.forEach(a=>{
    assert_(allowed.has(a.studentId),'متكون غير نشط أو غير منتسب للفوج.');
    assert_(!seen.has(a.studentId),'تكرار المتكون في التسجيل.');
    seen.add(a.studentId);
    const h=Number(a.hours);
    assert_((h===1||h===2)&&h<=Number(slot.duration),'ساعات الغياب غير صالحة.');
    abs.push({session_key:key,version:ver+1,student_id:a.studentId,absence_hours:h,justified:a.justified===true,notes:clean_(a.note).slice(0,200)});
  });
  append_('Sessions',[{session_key:key,version:ver+1,group_id:groupId,date:date,slot_id:slotId,start:slot.start,duration:Number(slot.duration),status:status,reason:status==='CANCELLED'?reason:'',notes:note,actor:u.username,at:now_()}]);
  append_('Attendance',abs);
  audit_(u,'SAVE_SESSION',key,{version:ver+1,oldStatus:old?old.status:null,status:status,absences:abs.length,reason:reason});
  return {sessionKey:key,version:ver+1};
}
function saveSession_(u,p) { worker_(u);return withLock_(()=>record_(u,p,false)); }
function cancelDay_(u,p) {
  worker_(u);
  return withLock_(()=>{
    const groupId=clean_(p.groupId),date=day_(p.date),reason=clean_(p.reason);
    assert_(can_(u,groupId),'ليست لديك صلاحية هذا الفوج.');
    assert_(reason.length>=3,'اذكر سبب توقف الحصص.');
    const slots=rows_('Schedule').filter(s=>s.group_id===groupId&&inDate_(s,date));
    assert_(slots.length>0,'لا توجد حصص مبرمجة في هذا اليوم.');
    const existing=latest_(rows_('Sessions'));
    const held=slots.some(s=>(existing[groupId+'|'+date+'|'+s.id]||{}).status==='HELD');
    assert_(!held,'اليوم يتضمن حصصًا مسجلة على أنها منجزة. عدّلها بصورة منفردة أولًا.');
    slots.forEach(s=>{
      const key=groupId+'|'+date+'|'+s.id;
      const old=existing[key];
      if(!old||old.status!=='CANCELLED'||old.reason!==reason)record_(u,{groupId:groupId,date:date,slotId:s.id,status:'CANCELLED',reason:reason,expectedVersion:old?old.version:0},true);
    });
    return {count:slots.length};
  });
}
function updateGroup_(u,p) {
  admin_(u);
  const id=clean_(p.groupId),h=Number(p.hours);
  assert_(Number.isInteger(h)&&h>=1&&h<=36,'الحجم الأسبوعي يجب أن يكون عدد ساعات من 1 إلى 36.');
  return withLock_(()=>{
    const g=group_(id);
    db_().getSheetByName('Groups').getRange(g._row,3).setValue(h);
    audit_(u,'WEEKLY_HOURS',id,{old:g.weekly_hours,new:h});
    return {hours:h};
  });
}
function addSlot_(u,p) {
  admin_(u);
  const groupId=clean_(p.groupId),start=clean_(p.start),weekday=Number(p.weekday),duration=Number(p.duration),from=day_(p.from||today_());
  group_(groupId);
  assert_(Number.isInteger(weekday)&&weekday>=0&&weekday<=6,'اليوم غير صالح.');
  assert_(/^([01]\d|2[0-3]):[0-5]\d$/.test(start),'أدخل الساعة بنظام 24 ساعة مثل 08:00.');
  assert_(duration===1||duration===2,'مدة الحصة ساعة أو ساعتان فقط.');
  const mins=Number(start.slice(0,2))*60+Number(start.slice(3));
  assert_(mins+duration*60<=1440,'الحصة تتجاوز نهاية اليوم.');
  return withLock_(()=>{
    const existing=rows_('Schedule').filter(s=>s.group_id===groupId&&Number(s.weekday)===weekday&&(!s.valid_to||s.valid_to>=from));
    assert_(!existing.some(s=>{const x=Number(String(s.start).slice(0,2))*60+Number(String(s.start).slice(3));return mins<x+Number(s.duration)*60 && x<mins+duration*60;}),'توجد حصة متداخلة في التوقيت.');
    const id=uuid_('SL');
    append_('Schedule',[{id:id,group_id:groupId,weekday:weekday,start:start,duration:duration,active:true,valid_from:from,valid_to:''}]);
    audit_(u,'ADD_SLOT',id,{groupId:groupId,weekday:weekday,start:start,duration:duration,from:from});
    return {id:id};
  });
}
function deactivateSlot_(u,p) {
  admin_(u);
  return withLock_(()=>{
    const s=rows_('Schedule').find(x=>x.id===clean_(p.slotId));
    assert_(s,'لم تُعثر على الحصة.');
    const end=day_(p.endDate||today_());
    assert_(!s.valid_to||s.valid_to>=end,'هذه الحصة أُوقفت سابقًا.');
    db_().getSheetByName('Schedule').getRange(s._row,8).setValue(end);
    audit_(u,'CLOSE_SLOT',s.id,{end:end});
    return {id:s.id};
  });
}
function updateStudent_(u,p) {
  admin_(u);
  const id=clean_(p.studentId),status=clean_(p.status);
  assert_(['REVIEW','ACTIVE','INACTIVE'].includes(status),'الحالة غير صالحة.');
  const effective=p.effectiveDate?day_(p.effectiveDate):today_();
  return withLock_(()=>{
    const s=rows_('Students').find(x=>x.id===id);
    assert_(s,'المتكون غير موجود.');
    const sh=db_().getSheetByName('Students');
    sh.getRange(s._row,4).setValue(status);
    if(status==='ACTIVE'){if(!s.start_date)sh.getRange(s._row,5).setValue(effective);sh.getRange(s._row,6).clearContent();}
    if(status==='INACTIVE')sh.getRange(s._row,6).setValue(effective);
    audit_(u,'STUDENT_STATUS',id,{old:s.status,new:status,effective:effective});
    return {id:id,status:status};
  });
}
function addStudent_(u,p) {
  admin_(u);
  const groupId=clean_(p.groupId),name=clean_(p.name),start=p.startDate?day_(p.startDate):today_();
  group_(groupId);
  assert_(name.length>=4&&name.length<=100,'أدخل اسم المتكون كاملًا.');
  return withLock_(()=>{
    assert_(!rows_('Students').some(s=>s.group_id===groupId&&s.name===name&&s.status!=='INACTIVE'),'يوجد متكون بهذا الاسم ضمن الفوج.');
    const id=uuid_('ST');
    append_('Students',[{id:id,group_id:groupId,name:name,status:'ACTIVE',start_date:start,end_date:'',notes:'أضيف من واجهة الإدارة'}]);
    audit_(u,'ADD_STUDENT',id,{name:name,groupId:groupId});
    return {id:id};
  });
}
function setOverride_(u,p) {
  admin_(u);
  const groupId=clean_(p.groupId),weekStart=day_(p.weekStart),hours=Number(p.hours);
  group_(groupId);
  assert_(weekday_(weekStart)===0,'تاريخ بداية الأسبوع يجب أن يكون يوم الأحد.');
  assert_(Number.isInteger(hours)&&hours>=0&&hours<=36,'حجم الساعات الأسبوعي يجب أن يكون بين 0 و36.');
  return withLock_(()=>{
    const sh=db_().getSheetByName('WeeklyOverrides');
    const old=rows_('WeeklyOverrides').find(x=>x.group_id===groupId&&x.week_start===weekStart);
    if(old)sh.getRange(old._row,3,1,2).setValues([[hours,clean_(p.comment).slice(0,200)]]);
    else append_('WeeklyOverrides',[{group_id:groupId,week_start:weekStart,planned_hours:hours,comment:clean_(p.comment).slice(0,200)}]);
    audit_(u,'WEEK_OVERRIDE',groupId,{weekStart:weekStart,hours:hours});
    return {hours:hours};
  });
}
function createUser_(u,p) {
  admin_(u);
  const username=clean_(p.username).toLowerCase(),role=clean_(p.role),groups=clean_(p.groups||'*');
  assert_(/^[a-z0-9._-]{3,30}$/.test(username),'اختر اسم مستخدم لاتينيًا بلا مسافات (3–30 رمزًا).');
  assert_(['admin','monitor','inspector'].includes(role),'الدور غير صالح.');
  if(role==='monitor')assert_(groups!=='*'&&groups.length>0,'حدد للمراقب فوجًا أو أكثر.');
  return withLock_(()=>{
    assert_(!rows_('Users').some(x=>clean_(x.username).toLowerCase()===username),'اسم المستخدم موجود مسبقًا.');
    const password=pin_();
    append_('Users',[{id:uuid_('USR'),username:username,hash:hash_(password),role:role,groups:groups,enabled:true,created_at:now_()}]);
    audit_(u,'CREATE_USER',username,{role:role,groups:groups});
    return {username:username,password:password,role:role};
  });
}
function sendMessage_(u,p) {
  const groupId=clean_(p.groupId),body=clean_(p.body),sessionKey=clean_(p.sessionKey).slice(0,250);
  assert_(can_(u,groupId),'ليست لديك صلاحية هذا الفوج.');
  assert_(body.length>=2&&body.length<=500,'الملاحظة يجب أن تتكون من 2 إلى 500 حرف.');
  return withLock_(()=>{
    const id=uuid_('MSG');
    append_('Messages',[{id:id,at:now_(),actor:u.username,role:u.role,group_id:groupId,session_key:sessionKey,body:body}]);
    return {id:id};
  });
}
/** نقطة الدخول الوحيدة للواجهة. كل عملية كتابة تتطلب جلسة صالحة وصلاحية مناسبة. */
function api(action,payload,token) {
  payload=payload||{};
  if(action==='login')return login_(payload);
  const u=auth_(token);
  switch(action) {
    case 'context': return context_(u,payload);
    case 'saveSession': return saveSession_(u,payload);
    case 'cancelDay': return cancelDay_(u,payload);
    case 'setWeeklyHours': return updateGroup_(u,payload);
    case 'addSlot': return addSlot_(u,payload);
    case 'deactivateSlot': return deactivateSlot_(u,payload);
    case 'setStudentStatus': return updateStudent_(u,payload);
    case 'addStudent': return addStudent_(u,payload);
    case 'setWeeklyOverride': return setOverride_(u,payload);
    case 'createUser': return createUser_(u,payload);
    case 'sendMessage': return sendMessage_(u,payload);
    default: fail_('عملية غير معروفة.');
  }
}