// MVP: inserta [PPP-YY-NNN] en el asunto del mensaje en COMPOSE.
// En modo lectura Office.js no permite modificar el asunto; eso lo haremos con Graph en el siguiente paso.

const RX_PREFIX = /^\[(\d{3}-\d{2})-(\d{3})\]\s*|^(\d{3}-\d{2})_(\d{3})\s*/i;

Office.onReady(() => {
  const btn = document.getElementById('applyCurrent');
  const pppyy = document.getElementById('pppyy');
  const nnn = document.getElementById('nnn');
  const out = document.getElementById('msg');

  btn.addEventListener('click', () => {
    const id = (pppyy.value || '').trim();
    const seq = (nnn.value || '').trim();
    if (!/^\d{3}-\d{2}$/.test(id)) return show(out, 'Formato PPP-YY inválido (ej. 123-25)', 'err');
    if (!/^\d{1,3}$/.test(seq))   return show(out, 'NNN debe ser 0–999', 'err');

    const n3 = String(parseInt(seq, 10)).padStart(3,'0');
    const prefix = `${id}_${n3}`;
    const tag = `[${id}_${n3}]`;

    const item = Office.context.mailbox.item;

    if (!item || !item.subject || !item.subject.setAsync) {
      return show(out, 'Abre un correo en composición para aplicar el número.', 'warn');
    }

    item.subject.getAsync(res => {
      const current = (res.value || '').replace(RX_PREFIX, '').trim();
      const subject = `${prefix} ${current}`.trim();
      const withTag = subject.includes(tag) ? subject : `${subject} ${tag}`;
      item.subject.setAsync(withTag, () => show(out, `Aplicado: ${withTag}`, 'ok'));
    });
  });
});

function show(el, text, cls){
  el.className = `muted ${cls||''}`;
  el.textContent = text;
}
