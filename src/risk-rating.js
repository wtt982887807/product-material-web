// Chinese names are a fallback only. Explicit R codes are resolved by the engine first.
const LEVELS={'低风险':'R1','中低风险':'R2','中风险':'R3','中高风险':'R4','高风险':'R5'};
const NAME='(中低风险|中高风险|中风险|低风险|高风险)';
const LABEL='(?:风险等级|风险评级|产品评级)';
const SUBJECT='本(?:集合资产管理计划|集合资金信托计划|资产管理计划|集合计划|信托计划|计划|产品|基金)';
const patterns=[
 new RegExp(LABEL+'[^。；;，,\\n]{0,65}?'+NAME,'g'),
 // A heading and its value may be in adjacent paragraphs/cells. Do not skip text.
 new RegExp(LABEL+'[：: \\t]*\\n[ \\t【\\[]*'+NAME,'g'),
 new RegExp(SUBJECT+'(?:属于|为|评定为|评价为|被评定为|被评价为)[\\s【\\[]*(?:R[1-5][（(\\s]*)?'+NAME,'g'),
];

export function verbalRiskCandidates(records,issuer='manager'){
 const result=[],explicit=[];
 const originals=records.filter(b=>!b.joined),positions=new Map(originals.map((b,i)=>[b.id,i]));
 for(const block of records){
  for(const pattern of patterns){
   for(const match of block.text.matchAll(pattern)){
    const prefix=block.text.slice(0,match.index);
    const clauseStart=Math.max(...['。','；',';'].map(c=>prefix.lastIndexOf(c)))+1;
    const previous=originals[positions.get((block.base||block).id)-1]?.text.trim()||'';
    const subjectHeading=previous.length<40&&/^(?:标的产品|底层|投资标的|投资者|客户|代销机构|销售机构|销售银行)(?:[：:\s]|$)/.test(previous)?previous:'';
    const context=subjectHeading+'\n'+block.text.slice(clauseStart,match.index+match[0].length);
    // Only the assertion's own subject matters, not the suitability text following it.
    if(/投资者|投资人|客户|底层|标的|所投|拟投/.test(context))continue;
    if(/并非|不是|不属于|不为|不评|不高于|不低于|原为|曾为|原评级|可能|若|如果|假如|假设|分为|划分|分成|分别|包括/.test(context))continue;
    const actors=[...context.matchAll(/管理人|代销(?:机构)?|销售机构|销售银行/g)];
    const distributor=actors.length>0&&actors.at(-1)[0]!=='管理人';
    if(!distributor&&/银行|证券公司|信托公司/.test(context)&&!/管理人/.test(context))continue;
    if((issuer==='distributor')!==distributor)continue;
    const following=block.text.slice(match.index+match[0].length);
    if(/^(?:或者|或|及|和|、|\/|／)+(?:中低风险|中高风险|中风险|低风险|高风险)/.test(following))continue;
    const code=match[0].match(/R[1-5]/)?.[0]||following.match(/^[（(【\[\s]*(R[1-5])\b/)?.[1];
    (code?explicit:result).push({value:code||LEVELS[match[1]],evidence:[{blockId:(block.base||block).id,quote:block.text.trim().slice(0,1600)}]});
   }
  }
 }
 return explicit.length?explicit:result;
}
