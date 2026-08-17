(function(){
  const key='bke-demo-cases-v1';
  const now=()=>new Date().toISOString();
  const day=()=>new Date().toISOString().slice(0,10).replaceAll('-','');
  function all(){try{const value=JSON.parse(localStorage.getItem(key)||'[]');return Array.isArray(value)?value:[]}catch{return[]}}
  function save(cases){localStorage.setItem(key,JSON.stringify(cases));return cases}
  function snapshot(form){const data={};form.querySelectorAll('[name]').forEach(field=>{if(field.type==='file')return;if(field.type==='checkbox'){data[field.name]??=[];if(field.checked)data[field.name].push(field.value)}else data[field.name]=field.value});return data}
  function task(title,role,name='未指派',status='未開始'){return{id:'task-'+Date.now()+'-'+Math.random().toString(36).slice(2,7),title,ownerRole:role,ownerName:name,status,due:'',note:'',updatedAt:now()}}
  function create({type,title,customer,applicant,summary,tasks}){const item={id:`${type}-${day()}-${String(Math.floor(Math.random()*900)+100)}`,type,title:title||`${type} 表需求案件`,customer:customer||'尚未填寫客戶',applicant:applicant||'尚未填寫申請人',status:'進行中',createdAt:now(),updatedAt:now(),summary:summary||{},tasks:tasks||[],activities:[]};const cases=all();cases.unshift(item);save(cases);return item}
  function update(item){const cases=all(),index=cases.findIndex(entry=>entry.id===item.id);item.updatedAt=now();if(index>=0)cases[index]=item;else cases.unshift(item);save(cases);return item}
  function find(id){return all().find(item=>item.id===id)||null}
  function tasksA(data){const applicant=data['申請人員']||'申請人員';return[task('客戶需求確認','業管',applicant,'已完成'),task('配方評估與報價','營養師'),task('樣品打樣','研發'),task('客戶確認與結案','業管',applicant)]}
  function tasksB(data){const result=[];const services=data['其他服務']||[],needs=data['產品需求']||[];if((data['成分明細表']||[]).length)result.push(task('成分明細表','營養師'));if((data['產品單張']||[]).length)result.push(task('產品單張','營養師'));if(data['彙辦單版本']||data['彙辦單_換算單位'])result.push(task('彙辦單','營養師'));if(['第一次校稿_申請日期','第二次校稿_申請日期','第三次校稿_申請日期'].some(key=>data[key]))result.push(task('包裝校稿','營養師'));if(needs.includes('商標授權')||(data['需商標授權']||[]).length)result.push(task('商標授權','營養師'));if(needs.includes('申請HALAL')||needs.includes('申請A.A.')||needs.includes('申請查驗登記')||data['技術資料申請目的'])result.push(task('認證／查驗登記','營養師'));services.forEach(name=>result.push(task(name,'行銷顧問')));return result.length?result:[task('確認產品資料需求','營養師')]}
  function tasksC(data){const result=[];if(data['官網審稿_提出時間']||data['官網審稿_需求時間']||data['官網審稿_網址'])result.push(task('官網產品審稿','行銷顧問'));if(data['產品介紹PPT_提出時間']||data['產品介紹PPT_需求時間'])result.push(task('產品介紹 PPT','行銷顧問'));const types=data['教育訓練類型']||[];if(types.includes('僅提供PPT')||data['教育訓練對象']||data['教育訓練包含產品'])result.push(task('教育訓練 PPT','行銷顧問'));if(types.includes('須外出授課')||data['外出授課需求時間']||data['外出授課場次'])result.push(task('外出教育訓練','行銷顧問'));return result.length?result:[task('確認行銷顧問需求','行銷顧問')]}
  window.BKECases={all,find,create,update,snapshot,task,tasksA,tasksB,tasksC,key};
})();
