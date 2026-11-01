import { useMemo, useState } from 'react';
import './DiscountSimulator.css';
const KEY='yu_discount_sim_plans_v1';
const num=v=>Math.max(0,Number(v)||0);
const won=v=>Math.round(v).toLocaleString('ko-KR')+'원';
const read=()=>{try{return JSON.parse(localStorage.getItem(KEY)||'[]')}catch{return []}};
const snapshot=products=>(products||[]).map(p=>({id:String(p.id),name:p.name||'이름 없음',stock:num(p.stock),wholesale:num(p.wholesale),retail:num(p.retail),category:p.category||'',included:true,rate:null,cap:null}));
export default function DiscountSimulator({products=[]}){
 const [plans,setPlans]=useState(read);
 const [planId,setPlanId]=useState(null);
 const [rows,setRows]=useState(()=>snapshot(products));
 const [globalRate,setGlobalRate]=useState(10);
 const [globalCap,setGlobalCap]=useState(3000);
 const [name,setName]=useState('재고정리 할인 계획');
 const [search,setSearch]=useState('');
 const [onlyIncluded,setOnlyIncluded]=useState(false);
 const [sort,setSort]=useState('name');
 const [dirty,setDirty]=useState(false);
 const update=(fn)=>{setRows(old=>fn(old));setDirty(true)};
 const change=(id,patch)=>update(old=>old.map(r=>r.id===id?{...r,...patch}:r));
 const price=r=>{const rate=r.rate===null?globalRate:r.rate;const cap=r.cap===null?globalCap:r.cap;const discount=Math.min(r.retail,num(cap),Math.round(r.retail*num(rate)/100));return {discount,sale:r.retail-discount,profit:r.retail-discount-r.wholesale}};
 const totals=useMemo(()=>rows.filter(r=>r.included).reduce((a,r)=>{const p=price(r),q=num(r.stock);a.qty+=q;a.cost+=r.wholesale*q;a.original+=r.retail*q;a.sales+=p.sale*q;a.discount+=p.discount*q;a.profit+=p.profit*q;return a},{qty:0,cost:0,original:0,sales:0,discount:0,profit:0}),[rows,globalRate,globalCap]);
 const visible=useMemo(()=>rows.filter(r=>(!onlyIncluded||r.included)&&`${r.name} ${r.category}`.toLowerCase().includes(search.toLowerCase())).sort((a,b)=>sort==='stock'?b.stock-a.stock:sort==='profit'?price(a).profit-price(b).profit:sort==='loss'?price(a).profit-price(b).profit:a.name.localeCompare(b.name,'ko')),[rows,search,onlyIncluded,sort,globalRate,globalCap]);
 const persist=next=>{localStorage.setItem(KEY,JSON.stringify(next));setPlans(next)};
 const save=()=>{const id=planId||String(Date.now());const item={id,name:name.trim()||'이름 없는 계획',globalRate,globalCap,rows,updatedAt:new Date().toISOString()};persist([item,...plans.filter(p=>p.id!==id)]);setPlanId(id);setDirty(false);alert('할인 계획을 이 브라우저에 저장했어요. 실제 재고는 변경되지 않았습니다.');};
 const load=id=>{const p=plans.find(p=>p.id===id);if(!p)return;if(dirty&&!confirm('저장하지 않은 변경사항을 버리고 계획을 불러올까요?'))return;setPlanId(id);setName(p.name);setGlobalRate(p.globalRate);setGlobalCap(p.globalCap);setRows(p.rows);setDirty(false)};
 const reset=()=>{if(dirty&&!confirm('저장하지 않은 변경사항을 버릴까요?'))return;setRows(snapshot(products));setPlanId(null);setName('재고정리 할인 계획');setGlobalRate(10);setGlobalCap(3000);setDirty(false)};
 const refresh=()=>{if(!confirm('실제 재고의 최신 수량·원가·정가를 다시 불러올까요? 현재 계획의 상품별 할인율과 제외 설정은 유지합니다. (저장 전까지 실제 재고에는 영향 없음)'))return;const prev=new Map(rows.map(r=>[r.id,r]));update(()=>snapshot(products).map(r=>({...r,included:prev.get(r.id)?.included??true,rate:prev.get(r.id)?.rate??null,cap:prev.get(r.id)?.cap??null})));};
 const exportPlan=()=>{const data={name,globalRate,globalCap,rows,exportedAt:new Date().toISOString()};const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([JSON.stringify(data,null,2)],{type:'application/json'}));a.download='재고할인계획.json';a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)};
 return <section className="discountSim panel"><h2>재고 할인 시뮬레이터</h2><p className="simHint">실제 재고관리 데이터는 절대 수정하지 않습니다. 상품 제외와 수량 변경은 이 계획 안에서만 적용돼요. 저장한 계획은 현재 브라우저에 보관됩니다.</p>
 <div className="simToolbar"><input aria-label="계획 이름" value={name} onChange={e=>{setName(e.target.value);setDirty(true)}} placeholder="계획 이름"/><button onClick={save}>계획 저장{dirty?' *':''}</button><select value={planId||''} onChange={e=>load(e.target.value)}><option value="">저장된 계획 불러오기</option>{plans.map(p=><option key={p.id} value={p.id}>{p.name} ({new Date(p.updatedAt).toLocaleDateString('ko-KR')})</option>)}</select><button onClick={reset}>새 계획</button><button onClick={refresh}>현재 재고 다시 불러오기</button><button onClick={exportPlan}>JSON 백업</button>{planId&&<button onClick={()=>{if(confirm('저장된 계획을 삭제할까요?')){persist(plans.filter(p=>p.id!==planId));setPlanId(null)}}}>저장 계획 삭제</button>}</div>
 <div className="simSettings"><label>전체 할인율 (%)<input type="number" min="0" max="100" value={globalRate} onChange={e=>{setGlobalRate(Math.min(100,num(e.target.value)));setDirty(true)}}/></label><label>최대 할인금액 (원)<input type="number" min="0" value={globalCap} onChange={e=>{setGlobalCap(num(e.target.value));setDirty(true)}}/></label><span>품목별 할인율·상한을 입력하면 전체 설정 대신 적용됩니다.</span></div>
 <div className="simStats">{[['판매 대상 수량',`${totals.qty.toLocaleString()}개`],['재고 원금',won(totals.cost)],['정가 기준 매출',won(totals.original)],['총 할인금액',won(totals.discount)],['할인 후 예상 매출',won(totals.sales)],['예상 매출총이익',won(totals.profit)],['원가 대비 이익률',totals.cost?`${(totals.profit/totals.cost*100).toFixed(1)}%`:'—']].map(([k,v])=><div key={k}><small>{k}</small><strong>{v}</strong></div>)}</div><p className="simHint">매출총이익 = 할인 후 판매금액 − 매입원가. 배송비, 플랫폼 수수료, 세금 등은 포함되지 않습니다.</p>
 <div className="simToolbar"><input placeholder="상품명 / 카테고리 검색" value={search} onChange={e=>setSearch(e.target.value)}/><label><input type="checkbox" checked={onlyIncluded} onChange={e=>setOnlyIncluded(e.target.checked)}/> 포함 상품만</label><select value={sort} onChange={e=>setSort(e.target.value)}><option value="name">상품명순</option><option value="stock">수량 많은순</option><option value="profit">개당 이익 높은순</option><option value="loss">개당 이익 낮은순</option></select><button onClick={()=>update(old=>old.map(r=>({...r,included:true})))}>전체 포함</button><button onClick={()=>update(old=>old.map(r=>({...r,included:false})))}>전체 제외</button></div>
 <div className="simTableWrap"><table className="simTable"><thead><tr><th>포함</th><th>상품명</th><th>재고수량 (계획)</th><th>원가/개</th><th>정가/개</th><th>개별 할인율 %</th><th>개별 할인상한 원</th><th>할인가/개</th><th>이익/개</th><th>총 예상이익</th></tr></thead><tbody>{visible.map(r=>{const p=price(r);return <tr key={r.id} className={!r.included?'simExcluded':''}><td><input type="checkbox" checked={r.included} onChange={e=>change(r.id,{included:e.target.checked})}/></td><td>{r.name}<small>{r.category}</small></td><td><input type="number" min="0" value={r.stock} onChange={e=>change(r.id,{stock:Math.floor(num(e.target.value))})}/></td><td>{won(r.wholesale)}</td><td>{won(r.retail)}</td><td><input type="number" min="0" max="100" placeholder="전체" value={r.rate??''} onChange={e=>change(r.id,{rate:e.target.value===''?null:Math.min(100,num(e.target.value))})}/></td><td><input type="number" min="0" placeholder="전체" value={r.cap??''} onChange={e=>change(r.id,{cap:e.target.value===''?null:num(e.target.value)})}/></td><td>{won(p.sale)}</td><td className={p.profit<0?'simLoss':''}>{won(p.profit)}</td><td className={p.profit<0?'simLoss':''}>{won(p.profit*r.stock)}</td></tr>})}</tbody></table></div><p className="simHint">표에 표시된 상품 {visible.length}종 / 전체 {rows.length}종 · 제외된 상품은 합계에서 빠집니다.</p></section>
}
