const profileSources=new Set(['company','manager']);
const ordered=values=>[...new Set(values)].sort((a,b)=>a-b);

/** No contract, product inference or parameter extraction occurs on this path. */
export function createIntroductionTask(templates,{now=new Date(),month,productionDate}={}){
 const task={updateMode:'introductions',target:{name:'公司和经理介绍更新',shortName:'介绍更新'},templates,month,productionDate,now,facts:[],materials:[],warnings:[],inputIssues:[],_records:[],stats:{}};
 for(const t of templates){
  if(t.kind==='pptx'&&/海报/.test(t.name)){task.inputIssues.push({sourceName:t.name,reason:'poster',message:'海报不更新公司和经理介绍，已跳过。'});continue;}
  task.materials.push({materialId:t.id,template:t,kind:t.kind==='docx'?'word':'ppt',updateMode:'introductions',selectedPages:[],profilePages:{company:[],manager:[]},operations:[],fields:[],warnings:[...(t.warnings||[])]});
 }
 return task;
}

export function finalizeIntroductionTask(task){
 const only=task.updateMode==='introductions';task.inputIssues??=[];
 for(const m of task.materials){
  if(only&&m.operations.some(o=>!profileSources.has(o.source)))throw new Error('仅更新介绍的任务包含非介绍修改，已停止输出。');
  const fields=m.fields.filter(f=>profileSources.has(f.source));
  if(m.kind==='ppt'){
   const valid=new Set(m.template.slides.map(s=>s.page));
   const pages=ordered(['company','manager'].flatMap(key=>m.profilePages?.[key]||[]));
   if(pages.some(page=>!Number.isInteger(page)||!valid.has(page)))throw new Error('介绍页面位置无效，请重新核对模板。');
   if(!only)m.productSelectedPages??=[...(m.selectedPages||[])];
   m.selectedPages=ordered([...(only?[]:m.productSelectedPages),...pages]);
  }
  m.introductionsLocated=fields.length>0;
 }
 if(only){
  task.materials=task.materials.filter(m=>{
   if(m.introductionsLocated&&(m.kind!=='ppt'||m.selectedPages.length))return true;
   task.inputIssues.push({sourceName:m.template.name,reason:'no_introductions',message:'未找到可与当前资料匹配的公司或经理介绍，已跳过。'});return false;
  });
  const ids=new Set(task.materials.map(m=>m.materialId));
  task.facts=task.facts.filter(f=>profileSources.has(f.source)&&f.occurrences?.some(o=>ids.has(o.materialId)));
 }
 task.stats={...task.stats,facts:task.facts.length,verified:task.facts.filter(f=>f.status==='verified').length,pending:task.facts.filter(f=>f.status!=='verified').length,changes:task.materials.reduce((n,m)=>n+m.operations.filter(o=>o.kind==='replace').length,0),materials:task.materials.length};
 return task;
}
