const btns=document.querySelectorAll('.fbtn');
const acards=document.querySelectorAll('.acard');
function applyFilter(f){
  btns.forEach(x=>x.classList.toggle('active', x.dataset.f===f));
  acards.forEach(a=>a.classList.toggle('sel', a.dataset.f===f));
  document.querySelectorAll('.card').forEach(c=>{
    c.style.display=(f==='all'||c.dataset.asset===f)?'':'none';
  });
}
btns.forEach(b=>b.addEventListener('click',()=>applyFilter(b.dataset.f)));
acards.forEach(a=>a.addEventListener('click',()=>{
  applyFilter(a.dataset.f);
  document.querySelector('.toolbar').scrollIntoView({behavior:'smooth',block:'start'});
}));
