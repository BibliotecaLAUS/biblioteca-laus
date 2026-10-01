let catalog=[];
const $=s=>document.querySelector(s);
const norm=s=>(s||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase();
const els={search:$("#search"),tiempo:$("#tiempo"),uso:$("#uso"),formacion:$("#formacion"),works:$("#works"),count:$("#count"),clear:$("#clear"),modal:$("#previewModal"),title:$("#previewTitle"),author:$("#previewAuthor"),img:$("#previewImage"),placeholder:$("#previewPlaceholder")};

function fillSelect(el,key){
  const values=[...new Set(catalog.flatMap(x=>x[key]||[]))].sort((a,b)=>a.localeCompare(b,"es"));
  values.forEach(v=>el.add(new Option(v,v)));
}
function matchesList(values,selected){return !selected||(values||[]).includes(selected)}
function render(){
  const q=norm(els.search.value);
  const rows=catalog.filter(w=>{
    const hay=norm([w.titulo,w.compositor,w.arreglador,w.codigo].join(" "));
    return (!q||hay.includes(q))&&matchesList(w.tiempo,els.tiempo.value)&&matchesList(w.uso,els.uso.value)&&matchesList(w.formacion,els.formacion.value);
  });
  els.count.textContent=rows.length+" "+(rows.length===1?"obra":"obras");
  els.works.innerHTML=rows.length?rows.map(w=>`
    <article class="work">
      <div class="work-main"><h2>${w.titulo}</h2><p class="composer">${w.compositor}</p></div>
      <p class="meta">${w.formacion.join(", ")}${w.instrumentacion?" · "+w.instrumentacion:""}</p>
      <div class="actions">
        <button class="icon-button" type="button" data-preview="${w.codigo}" aria-label="Ver primera página de ${w.titulo}" title="Vista previa">◉</button>
        <a class="material" href="${w.material}" target="_blank" rel="noopener" aria-label="Abrir material musical" title="Material musical">↗</a>
      </div>
    </article>`).join(""):'<p class="empty">No encontramos obras con esos filtros.</p>';
}
function openPreview(code){
  const w=catalog.find(x=>x.codigo===code); if(!w)return;
  els.title.textContent=w.titulo; els.author.textContent=w.compositor;
  if(w.preview){els.img.src=w.preview;els.img.alt="Primera página de "+w.titulo;els.img.hidden=false;els.placeholder.hidden=true}
  else{els.img.hidden=true;els.placeholder.hidden=false}
  els.modal.hidden=false;document.body.classList.add("modal-open");$(".close").focus();
}
function closePreview(){els.modal.hidden=true;document.body.classList.remove("modal-open")}
document.addEventListener("click",e=>{const p=e.target.closest("[data-preview]");if(p)openPreview(p.dataset.preview);if(e.target.closest("[data-close]"))closePreview()});
document.addEventListener("keydown",e=>{if(e.key==="Escape"&&!els.modal.hidden)closePreview()});
[els.search,els.tiempo,els.uso,els.formacion].forEach(el=>el.addEventListener("input",render));
els.clear.addEventListener("click",()=>{els.search.value="";els.tiempo.value="";els.uso.value="";els.formacion.value="";render();els.search.focus()});
fetch("catalogo.json").then(r=>r.json()).then(data=>{catalog=data;fillSelect(els.tiempo,"tiempo");fillSelect(els.uso,"uso");fillSelect(els.formacion,"formacion");render()}).catch(()=>{els.works.innerHTML='<p class="empty">No se pudo cargar el catálogo.</p>'});