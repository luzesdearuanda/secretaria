/* Secretaria Luzes de Aruanda — módulos opcionais (versão 1.5).
   PIX, cobrança, link pessoal com mensalidades, agenda com presença e ata, caminhada espiritual,
   campanhas, assembleias, livro de registro, estoque e cópia de segurança.
   Cada módulo pode ser ligado ou desligado em Configurações; desligado, some das telas. */
'use strict';
Object.assign(TBL, { Eventos: ['eventos', 'id'], Presencas: ['presencas', 'id'], Marcos: ['marcos', 'id'], Pessoas: ['pessoas', 'id'], Emissoes: ['emissoes', 'id'], Entregas: ['entregas', 'id'], Campanhas: ['campanhas', 'id'], Contribuicoes: ['contribuicoes', 'id'],
  Itens: ['itens', 'id'], MovEstoque: ['movEstoque', 'id'], Movimentacoes: ['movimentacoes', 'id'], Cobrancas: ['cobrancas', 'id'] });
Object.assign(I, {
  calendar: '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/>',
  gift: '<rect x="3" y="8" width="18" height="5" rx="1"/><path d="M5 13v8h14v-8M12 8v13M12 8c-2-4-6-4-6-1.5S10 8 12 8zM12 8c2-4 6-4 6-1.5S14 8 12 8z"/>',
  box: '<path d="M3 7l9-4 9 4v10l-9 4-9-4z"/><path d="M3 7l9 4 9-4M12 11v10"/>',
  book: '<path d="M4 4h6a3 3 0 0 1 3 3v13a2 2 0 0 0-2-2H4z"/><path d="M20 4h-6a3 3 0 0 0-3 3v13a2 2 0 0 1 2-2h7z"/>',
  send: '<path d="M21 3L3 10.5l7 2.5 2.5 7z"/><path d="M10 13l4.5-4.5"/>',
  qr: '<rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/><path d="M14 14h3v3h-3zM18 18h3v3M14 21h1M21 14v1"/>',
  vote: '<rect x="4" y="11" width="16" height="10" rx="1"/><path d="M8 11V6l4-3 4 3v5M10 15h4"/>',
  leaf: '<path d="M5 19c0-8 6-14 15-14 0 9-6 15-14 15"/><path d="M5 19l8-8"/>',
  db: '<ellipse cx="12" cy="5" rx="8" ry="3"/><path d="M4 5v14c0 1.7 3.6 3 8 3s8-1.3 8-3V5M4 12c0 1.7 3.6 3 8 3s8-1.3 8-3"/>',
});
const L = k => (S && S[k]) || [];
const MODULOS_INFO = [
  ['pix', 'PIX da mensalidade', 'QR Code e "copia e cola" com o valor dos meses em aberto, para mostrar na tela ou mandar por WhatsApp.'],
  ['cobranca', 'Cobrança e avisos', 'Lista de quem está em aberto com WhatsApp e e-mail prontos; e-mail automático mensal e resumo para a tesouraria (opcional).'],
  ['portal', 'Mensalidades no link pessoal', 'O link pessoal da carteirinha mostra os meses pagos, os em aberto e o PIX para pagar.'],
  ['agenda', 'Agenda, presença e ata', 'Giras e eventos, presença pela leitura do QR da carteirinha, frequência por médium e ata após cada gira ou evento.'],
  ['historico', 'Caminhada na casa', 'Batismo, amaci, cruzamento, coroação e cursos no cadastro de cada associado.'],
  ['campanhas', 'Campanhas e doações', 'Arrecadações com meta, doações em dinheiro (entram no livro-caixa) e em itens, com a lista de quem contribuiu.'],
  ['assembleias', 'Assembleias', 'Quem pode votar, lista de presença para assinar, quórum, edital de convocação e ata.'],
  ['livro', 'Livro de registro de associados', 'Relatório para imprimir com termo de abertura, associados, entradas e mudanças de situação.'],
  ['estoque', 'Estoque e patrimônio', 'Velas, ervas e materiais com alerta de estoque baixo, e o inventário dos bens da casa.'],
  ['documentos', 'Documentos', 'Fichas, termos de consentimento (praia, amaci, corte), certificados de batismo e certidões de casamento, em massa ou para pessoas específicas, inclusive não associados.'],
  ['backup', 'Cópia de segurança', 'Cópia da planilha numa pasta do Drive, na hora ou toda semana, guardando as últimas versões.'],
];
const TIPOS_EVENTO = ['Gira', 'Festa', 'Batismo', 'Reunião', 'Curso', 'Outro'];
// versão 1.9: sessões espirituais têm linha, pública/privada, nº da ata e números
const SESSOES = ['Gira', 'Festa', 'Batismo'];
const NUMS_SESSAO = [['mediuns', 'Médiuns presentes'], ['faltas', 'Faltas'], ['faltasJust', 'Faltas justificadas'], ['consultas', 'Consultas atendidas'], ['passes', 'Passes'], ['encaminhamentos', 'Encaminhamentos para cura']];
const MIME_WORD = { docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', doc: 'application/msword' };
const ehDirigente = a => funcoesAtuais(a.id).some(x => { const f = funcao(x.funcaoId); return f && /dirigente espiritual/i.test(f.nome); });
// ata: PDF/foto (reduzida) ou Word (enviado como está)
function prepAta(f) {
  const ext = (f.name.match(/\.(docx?)$/i) || [])[1];
  if (!ext) return prepComprovante(f, 2400, 0.85);
  if (f.size > 6 * 1024 * 1024) return Promise.reject(new Error('Arquivo maior que 6 MB.'));
  return blobB64(f).then(d => ({ mime: MIME_WORD[ext.toLowerCase()], nome: f.name, dados: d }));
}
const TIPOS_MARCO = ['Batismo', 'Casamento', 'Amaci', 'Cruzamento', 'Coroação', 'Desenvolvimento mediúnico', 'Curso', 'Outro'];
const DIAS_SEMANA = ['domingo', 'segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira', 'sexta-feira', 'sábado'];
const mesFimISO = c => { const [y, m] = c.split('-').map(Number); return isoLocal(new Date(y, m, 0)); };
const num = v => Number(v || 0).toLocaleString('pt-BR', { maximumFractionDigits: 3 });
const ativosOrd = () => S.associados.filter(a => a.situacao === 'Ativo').sort((x, y) => x.nome.localeCompare(y.nome, 'pt-BR'));
const assocOpts = (id = 'm-alist', lista = ativosOrd()) => `<datalist id="${id}">${lista.map(a => `<option value="${fmtMat(a.matricula)} · ${esc(a.nome)}">`).join('')}</datalist>`;
function assocDoTexto(v) { const m = /^(\d{1,6})\s*·/.exec(String(v || '').trim()) || /^(\d{1,6})$/.exec(String(v || '').trim()); return m ? S.associados.find(a => a.matricula === +m[1]) : S.associados.find(a => norm(a.nome) === norm(v)); }

/* ============ PIX (BR Code estático do Banco Central) ============ */
const semAcento = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Za-z0-9 .,\-@+/]/g, '').toUpperCase();
function crc16(s) { let c = 0xFFFF; for (let i = 0; i < s.length; i++) { c ^= s.charCodeAt(i) << 8; for (let j = 0; j < 8; j++) c = c & 0x8000 ? ((c << 1) ^ 0x1021) & 0xFFFF : (c << 1) & 0xFFFF; } return c.toString(16).toUpperCase().padStart(4, '0'); }
function pixChave(cfg) {
  const k = String(cfg.pixChave || '').trim(); if (!k) return '';
  if (cfg.pixTipo === 'Telefone') { const d = digits(k); return d ? (d.length <= 11 ? '+55' + d : '+' + d) : ''; }
  if (cfg.pixTipo === 'CPF/CNPJ') return digits(k);
  if (cfg.pixTipo === 'E-mail') return k.toLowerCase();
  return k;
}
function pixPayload(valor, txid) {
  const cfg = S.config, chave = pixChave(cfg); if (!chave) return '';
  const f = (id, v) => id + String(v.length).padStart(2, '0') + v;
  const nome = semAcento(cfg.pixNome || assocCfg().nome).slice(0, 25).trim(), cidade = semAcento(cfg.pixCidade || 'PASSO FUNDO').slice(0, 15).trim();
  const tx = String(txid || '***').toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 25) || '***';
  const s = f('00', '01') + f('26', f('00', 'br.gov.bcb.pix') + f('01', chave)) + f('52', '0000') + f('53', '986') + (valor > 0 ? f('54', (Math.round(valor * 100) / 100).toFixed(2)) : '') + f('58', 'BR') + f('59', nome) + f('60', cidade) + f('62', f('05', tx)) + '6304';
  return s + crc16(s);
}
const pixPronto = () => modOn('pix') && !!pixChave(S.config);
function qrSvg(texto, tam = 220) {
  if (typeof qrcode !== 'function') return '';
  const q = qrcode(0, 'M'); q.addData(texto); q.make();
  const n = q.getModuleCount(), m = tam / (n + 8); let d = '';
  for (let r = 0; r < n; r++) for (let c = 0; c < n; c++) if (q.isDark(r, c)) d += `M${((c + 4) * m).toFixed(2)} ${((r + 4) * m).toFixed(2)}h${m.toFixed(2)}v${m.toFixed(2)}h-${m.toFixed(2)}z`;
  return `<svg class="pix-qr" viewBox="0 0 ${tam} ${tam}" width="${tam}" height="${tam}" role="img" aria-label="QR Code PIX"><rect width="${tam}" height="${tam}" fill="#fff"/><path d="${d}" fill="#000" shape-rendering="crispEdges"/></svg>`;
}
// Mensagem de cobrança (WhatsApp) com os meses, o total e, se houver, o PIX e o link pessoal
function msgCobranca(a, meses, total, pix) {
  const c = carteirinhaAtiva(a.id), link = modOn('portal') && c ? linkPessoal(c) : '';
  return `Olá, ${a.nome.split(' ')[0]}! Aqui é da secretaria da ${assocCfg().nome}. Consta em aberto ${meses.length === 1 ? 'a mensalidade de' : 'as mensalidades de'} ${meses.map(fmtComp).join(', ')}, no total de ${fmtBRL(total)}.`
    + (pix ? `\n\nPIX copia e cola:\n${pix}` : '') + (link ? `\n\nSituação e PIX no seu link pessoal: ${link}` : '') + '\n\nSe já pagou, desconsidere. Obrigado!';
}

/* ============ Hooks usados pelo app.js ============ */
Object.assign(MOD, {
  // Botões extras no painel de mensalidades do associado
  botoesMens(a, fin) {
    if (a.situacao !== 'Ativo' || !fin.cobrado) return '';
    return (pixPronto() ? `<button class="btn sm" data-act="pixAbrir" data-id="${a.id}">${ic('qr')} PIX</button>` : '')
      + (modOn('cobranca') && fin.aberto.length ? `<button class="btn sm" data-act="cobWhats" data-id="${a.id}" ${a.telefone ? '' : 'disabled title="Sem telefone cadastrado"'}>${ic('whats')} Cobrar</button>` : '');
  },
  // Painel extra no cadastro do associado: caminhada na casa e frequência
  painelAssoc(a) {
    if (!modOn('historico') && !modOn('agenda')) return '';
    if (ehDirigente(a)) return ''; // versão 1.9: dirigentes da casa não têm esta seção
    const ms = L('marcos').filter(m => m.assocId === a.id && !m.excluido).sort((x, y) => y.data.localeCompare(x.data));
    const desde = addDays(TODAY(), -90), evs = L('eventos').filter(e => !e.cancelado && e.tipo === 'Gira' && e.data >= desde && e.data <= TODAY());
    const pres = evs.filter(e => L('presencas').some(p => p.eventoId === e.id && p.assocId === a.id)).length;
    return `<section class="panel"><div class="panel-head"><h2>${modOn('historico') ? 'Caminhada na casa' : 'Frequência'}</h2>${modOn('historico') ? `<button class="btn sm primary" data-act="marcoNovo" data-id="${a.id}">${ic('plus')} Registrar</button>` : ''}</div>
      ${modOn('agenda') ? `<p class="small muted">${evs.length ? `Presença nas giras dos últimos 90 dias: <b>${pres} de ${evs.length}</b> (${Math.round(pres / evs.length * 100)}%).` : 'Nenhuma gira registrada na agenda nos últimos 90 dias.'}</p>` : ''}
      ${modOn('historico') ? (ms.length ? `<ol class="timeline">${ms.map(m => `<li><b>${esc(m.tipo)}</b> <span class="muted">· ${fmtD(m.data)}</span>${m.descricao ? `<div class="small">${esc(m.descricao)}</div>` : ''}<div class="small">${m.arquivoId ? `<button class="btn sm ghost" data-act="marcoVer" data-id="${m.id}">${ic('file')} Ver certidão</button>` : ''}<button class="btn sm ghost" data-act="marcoAnexar" data-id="${m.id}">${ic('clip')} ${m.arquivoId ? 'Trocar certidão' : 'Anexar certidão'}</button></div>${isAdmin() ? `<button class="btn sm ghost icon" data-act="marcoExcluir" data-id="${m.id}" aria-label="Excluir registro" title="Excluir registro">${ic('x')}</button>` : ''}</li>`).join('')}</ol>`
        : '<p class="small muted">Nada registrado ainda. Registre batismo, casamento, amaci, cruzamento, coroação, cursos e outras etapas, com a certidão quando houver.</p>') : ''}
    </section>`;
  },
  kpis() {
    let h = '';
    if (modOn('agenda')) {
      const prox = L('eventos').filter(e => !e.cancelado && e.data >= TODAY()).sort((x, y) => (x.data + x.hora).localeCompare(y.data + y.hora))[0];
      h += kpiBox('Próximo evento', prox ? fmtD(prox.data).slice(0, 5) : '—', prox ? `${esc(prox.titulo)}${prox.hora ? ' · ' + prox.hora : ''}` : 'nada agendado', '', prox ? `data-act="go" data-v="evento" data-id="${prox.id}"` : 'data-act="go" data-v="agenda"');
    }
    if (modOn('estoque')) { const b = estoqueBaixo(); if (b.length) h += kpiBox('Estoque baixo', b.length, b.slice(0, 2).map(i => esc(i.nome)).join(', ') + (b.length > 2 ? '…' : ''), 'warn', 'data-act="go" data-v="estoque"'); }
    return h;
  },
  menuMais() {
    return NAV.filter(n => n.mod && modOn(n.mod) && (!n.admin || isAdmin()) && (!n.fin || canFin())).map(n => `<button data-act="go" data-v="${n.id}">${ic(n.ic)} ${n.label}</button>`).join('');
  },
  configHTML,
  csvExtra: null,
});

/* ============ Configurações dos módulos ============ */
function configHTML() {
  const c = S.config, m = c.modulos || {};
  return `<form class="panel form" data-form="modulos" novalidate><h2>Módulos opcionais</h2>
    <p class="small muted">Ligue só o que a casa vai usar. Desligar um módulo apenas esconde as telas dele: nada é apagado e dá para ligar de novo quando quiser.</p>
    <div class="mod-list">${MODULOS_INFO.map(([k, nome, desc]) => `<label class="mod-item"><input type="checkbox" id="mod-${k}" ${m[k] ? 'checked' : ''}><span><b>${nome}</b><span class="small muted">${desc}</span></span></label>`).join('')}</div>
    <fieldset><legend>PIX</legend><div class="fields">
      <div class="field"><label for="px-tipo">Tipo de chave</label><select id="px-tipo">${['CPF/CNPJ', 'Telefone', 'E-mail', 'Aleatória'].map(t => `<option ${c.pixTipo === t ? 'selected' : ''}>${t}</option>`).join('')}</select></div>
      <div class="field"><label for="px-chave">Chave PIX da associação</label><input id="px-chave" value="${esc(c.pixChave || '')}" maxlength="77" placeholder="Ex.: CNPJ da associação"><span class="hint">Sem chave, o PIX não aparece.</span></div>
      <div class="field"><label for="px-nome">Nome do recebedor</label><input id="px-nome" value="${esc(c.pixNome || '')}" maxlength="25" placeholder="Até 25 letras, como no banco"></div>
      <div class="field"><label for="px-cid">Cidade</label><input id="px-cid" value="${esc(c.pixCidade || '')}" maxlength="15" placeholder="Até 15 letras"></div>
    </div></fieldset>
    <fieldset><legend>Cobrança automática por e-mail</legend><div class="fields">
      <div class="field full"><label class="check"><input type="checkbox" id="cb-auto" ${c.cobrancaAuto ? 'checked' : ''}> Todo mês, enviar e-mail para quem tem mensalidade atrasada (e um resumo para tesouraria e administração, com as isenções que estão terminando)</label></div>
      <div class="field"><label for="cb-dia">Dia do mês</label><input id="cb-dia" type="number" min="1" max="28" value="${c.cobrancaDia || 10}"></div>
    </div></fieldset>
    <fieldset><legend>Cópia de segurança</legend><div class="fields">
      <div class="field full"><label class="check"><input type="checkbox" id="bk-auto" ${c.backupAuto ? 'checked' : ''}> Fazer uma cópia da planilha toda semana</label></div>
      <div class="field"><label for="bk-dia">Dia da semana</label><select id="bk-dia">${DIAS_SEMANA.map((d, i) => `<option value="${i}" ${Number(c.backupDiaSemana) === i ? 'selected' : ''}>${d}</option>`).join('')}</select></div>
      <div class="field"><label for="bk-manter">Guardar as últimas</label><input id="bk-manter" type="number" min="1" max="52" value="${c.backupManter || 8}"><span class="hint">cópias (as mais antigas vão para a lixeira do Drive).</span></div>
      <div class="field full"><div class="inline"><button type="button" class="btn sm" data-act="bkAgora">${ic('db')} Fazer cópia agora</button><button type="button" class="btn sm ghost" data-act="bkListar">Ver cópias</button></div><div id="bk-lista"></div></div>
    </div></fieldset>
    <p class="small muted">As automações (cobrança e cópia semanal) rodam pela rotina diária do Apps Script, criada uma vez na implantação (veja o guia).</p>
    <div class="form-actions"><button class="btn primary">Salvar módulos</button></div></form>`;
}
FORM.modulos = () => {
  const v = id => ($('#' + id)?.value || '').trim(), ck = id => !!$('#' + id)?.checked;
  const modulos = {}; MODULOS_INFO.forEach(([k]) => { modulos[k] = ck('mod-' + k); });
  run(async () => {
    await api('config.modulos', { config: { modulos, pixTipo: v('px-tipo'), pixChave: v('px-chave'), pixNome: v('px-nome'), pixCidade: v('px-cid'), cobrancaAuto: ck('cb-auto'), cobrancaDia: v('cb-dia'), backupAuto: ck('bk-auto'), backupDiaSemana: v('bk-dia'), backupManter: v('bk-manter') } });
    render(); toast('Módulos salvos');
  });
};
Object.assign(ACT, {
  bkAgora() { run(async () => { const r = await api('backup.agora', {}, { msg: 'Copiando a planilha…' }); toast('Cópia feita: ' + r.nome); ACT.bkListar(); }); },
  bkListar() {
    run(async () => {
      const r = await api('backup.listar', {}, { msg: 'Buscando cópias…' }), el = $('#bk-lista'); if (!el) return;
      el.innerHTML = r.arquivos.length ? `<ul class="small bk-list">${r.arquivos.map(f => `<li><a href="${esc(f.url)}" target="_blank" rel="noopener">${esc(f.nome)}</a></li>`).join('')}</ul>` : '<p class="small muted">Nenhuma cópia ainda.</p>';
    });
  },
});

/* ============ PIX e cobrança ============ */
Object.assign(ACT, {
  pixAbrir(t) {
    const a = assoc(t.dataset.id), fin = financeiro(a);
    const futuros = []; for (let i = 1; i <= 3; i++) { const c = addMonthsComp(COMP_ATUAL(), i); if (!fin.pagas[c] && !fin.isentos[c]) futuros.push(c); }
    const opc = [...fin.aberto, ...futuros];
    UI.pix = { id: a.id, valor: fin.valor };
    openModal(`<h2>PIX da mensalidade</h2><p class="muted">${esc(a.nome)} · ${fmtBRL(fin.valor)}/mês</p>
      <div class="month-picks">${opc.map(c => `<label class="mpick ${c < COMP_ATUAL() ? 'late' : c > COMP_ATUAL() ? 'future' : ''}"><input type="checkbox" data-ch="pixMes" value="${c}" ${fin.aberto.includes(c) ? 'checked' : ''}><span>${fmtComp(c)}</span></label>`).join('')}</div>
      <div id="pix-box"></div>
      <p class="small muted">O pagamento cai na conta da associação. Depois de conferir no banco, registre em "Registrar pagamento".</p>
      <div class="modal-actions"><button class="btn" data-act="close">Fechar</button>${a.telefone ? `<button class="btn" data-act="pixWhats">${ic('whats')} Enviar por WhatsApp</button>` : ''}<button class="btn primary" data-act="pixCopiar">Copiar código</button></div>`);
    pixAtualiza();
  },
  pixCopiar() { if (UI.pix && UI.pix.codigo) copyText(UI.pix.codigo, 'Código PIX copiado'); },
  pixWhats() {
    const a = assoc(UI.pix.id); if (!UI.pix.codigo) { toast('Escolha ao menos um mês.', 'bad'); return; }
    window.open(`https://wa.me/${telWhats(a.telefone)}?text=${encodeURIComponent(msgCobranca(a, UI.pix.meses, UI.pix.total, UI.pix.codigo))}`, '_blank', 'noopener');
    if (modOn('cobranca')) api('cobranca.registrar', { assocId: a.id, canal: 'WhatsApp' }, { silent: true }).catch(() => { });
  },
  cobWhats(t) {
    const a = assoc(t.dataset.id), fin = financeiro(a); if (!fin.aberto.length) return;
    const total = Math.round(fin.aberto.length * fin.valor * 100) / 100, pix = pixPronto() ? pixPayload(total, 'MENS' + fmtMat(a.matricula)) : '';
    window.open(`https://wa.me/${telWhats(a.telefone)}?text=${encodeURIComponent(msgCobranca(a, fin.aberto, total, pix))}`, '_blank', 'noopener');
    run(async () => { await api('cobranca.registrar', { assocId: a.id, canal: 'WhatsApp' }, { silent: true }); if (VIEW === 'cobranca') $('#cob-res').innerHTML = cobLista(); });
  },
  cobEmail(t) { const a = assoc(t.dataset.id); run(async () => { const r = await api('cobranca.email', { assocId: a.id }, { msg: 'Enviando e-mail…' }); toast(`Cobrança de ${fmtBRL(r.total)} enviada para ${a.email}`); if (VIEW === 'cobranca') $('#cob-res').innerHTML = cobLista(); }); },
  cobEmailTodos() {
    const lista = cobDados().filter(x => x.a.email && x.fin.atraso);
    if (!lista.length) { toast('Ninguém com atraso e e-mail cadastrado.', 'bad'); return; }
    openModal(`<h2>Enviar cobrança por e-mail</h2><p>Enviar para <b>${plural(lista.length, 'associado', 'associados')}</b> com mensalidade atrasada e e-mail cadastrado?</p><p class="small muted">${lista.map(x => esc(x.a.nome)).join(', ')}</p>
      <div class="modal-actions"><button class="btn" data-act="close">Cancelar</button><button class="btn primary" data-act="cobEmailTodosOk">Enviar ${lista.length}</button></div>`);
  },
  cobEmailTodosOk() {
    const lista = cobDados().filter(x => x.a.email && x.fin.atraso); closeModal();
    run(async () => {
      let ok = 0, falha = 0;
      for (const x of lista) { try { await api('cobranca.email', { assocId: x.a.id }, { msg: `Enviando ${ok + falha + 1} de ${lista.length}…` }); ok++; } catch (e) { falha++; } }
      render(); toast(`${plural(ok, 'e-mail enviado', 'e-mails enviados')}${falha ? ` · ${falha} com erro` : ''}`, falha ? 'bad' : 'ok');
    });
  },
});
CHG.pixMes = () => pixAtualiza();
function pixAtualiza() {
  const meses = $$('[data-ch=pixMes]:checked').map(x => x.value).sort(), a = assoc(UI.pix.id);
  const total = Math.round(meses.length * UI.pix.valor * 100) / 100, codigo = meses.length ? pixPayload(total, 'MENS' + fmtMat(a.matricula)) : '';
  Object.assign(UI.pix, { meses, total, codigo });
  $('#pix-box').innerHTML = codigo ? `<div class="pix-box">${qrSvg(codigo)}<div><div class="k-label">Total</div><div class="pix-total">${fmtBRL(total)}</div><div class="small muted">${meses.map(fmtComp).join(', ')}</div><label class="small" for="pix-cod">PIX copia e cola</label><textarea id="pix-cod" class="mono small" rows="4" readonly onfocus="this.select()">${esc(codigo)}</textarea></div></div>` : '<p class="small muted">Marque os meses para gerar o PIX.</p>';
}
function cobDados() {
  return S.associados.filter(a => a.situacao === 'Ativo').map(a => ({ a, fin: financeiro(a) })).filter(x => x.fin.cobrado && x.fin.aberto.length)
    .map(x => ({ ...x, total: Math.round(x.fin.aberto.length * x.fin.valor * 100) / 100, ult: L('cobrancas').filter(c => c.assocId === x.a.id).sort((p, q) => q.em.localeCompare(p.em))[0] }))
    .filter(x => (P.f || 'atraso') === 'todos' || x.fin.atraso)
    .sort((x, y) => y.fin.atraso - x.fin.atraso || x.a.nome.localeCompare(y.a.nome, 'pt-BR'));
}
function cobLista() {
  const l = cobDados();
  if (!l.length) return '<div class="table-wrap"><div class="empty">Ninguém para cobrar nesta lista.</div></div>';
  return `<div class="table-wrap cards"><table><thead><tr><th>Associado</th><th>Em aberto</th><th>Total</th><th>Última cobrança</th><th><span class="sr">Ações</span></th></tr></thead><tbody>
    ${l.map(x => `<tr><td class="lead"><button class="who-row linklike" data-act="go" data-v="associado" data-id="${x.a.id}">${avatar(x.a)}<div><div class="t">${esc(x.a.nome)}</div><div class="s">${x.fin.atraso ? pill(plural(x.fin.atraso, 'mês atrasado', 'meses atrasados'), x.fin.bloqueia ? 'bad' : 'warn') : pill('só o mês atual', 'info')}</div></div></button></td>
      <td data-label="Em aberto" class="small">${x.fin.aberto.map(fmtComp).join(', ')}</td><td data-label="Total" class="num"><b>${fmtBRL(x.total)}</b></td>
      <td data-label="Última cobrança" class="small">${x.ult ? `${fmtStamp(x.ult.em)} · ${esc(x.ult.canal)}` : '—'}</td>
      <td class="actions">${pixPronto() ? `<button class="btn sm" data-act="pixAbrir" data-id="${x.a.id}">${ic('qr')} PIX</button>` : ''}<button class="btn sm" data-act="cobWhats" data-id="${x.a.id}" ${x.a.telefone ? '' : 'disabled title="Sem telefone"'}>${ic('whats')} WhatsApp</button><button class="btn sm" data-act="cobEmail" data-id="${x.a.id}" ${x.a.email ? '' : 'disabled title="Sem e-mail"'}>${ic('mail')} E-mail</button></td></tr>`).join('')}
  </tbody></table></div>`;
}
VIEWS.cobranca = {
  mod: 'cobranca',
  render() {
    const f = P.f || 'atraso', todos = S.associados.filter(a => a.situacao === 'Ativo').map(a => financeiro(a)).filter(x => x.cobrado && x.aberto.length);
    const tab = (id, label, n) => `<button class="tab ${f === id ? 'on' : ''}" data-act="cobF" data-f="${id}">${label} <span class="count">${n}</span></button>`;
    const c = S.config;
    return `<div class="page-head"><div><h1>Cobrança</h1><p class="sub">Quem tem mensalidade em aberto, com mensagem pronta para WhatsApp e e-mail${pixPronto() ? ' e o PIX do valor' : ''}.</p></div>
      <div class="toolbar"><button class="btn primary" data-act="cobEmailTodos">${ic('send')} E-mail para todos com atraso</button></div></div>
      ${!pixPronto() && modOn('pix') && isAdmin() ? `<div class="note warn"><span>Cadastre a chave PIX em Configurações para incluir o PIX nas cobranças.</span></div>` : ''}
      <p class="small muted">Cobrança automática por e-mail: ${c.cobrancaAuto ? `<b>ligada</b>, todo dia ${c.cobrancaDia} do mês` : 'desligada'}${isAdmin() ? ' (muda em Configurações)' : ''}.</p>
      <div class="tabs" role="tablist">${tab('atraso', 'Com atraso', todos.filter(x => x.atraso).length)}${tab('todos', 'Todos em aberto (inclui o mês atual)', todos.length)}</div>
      <div id="cob-res">${cobLista()}</div>`;
  }
};
ACT.cobF = t => { P.f = t.dataset.f; render(); };

/* ============ Agenda, presença e ata ============ */
const evento = id => L('eventos').find(e => e.id === id);
const presentes = eid => L('presencas').filter(p => p.eventoId === eid).map(p => ({ p, a: assoc(p.assocId) })).filter(x => x.a).sort((x, y) => x.a.nome.localeCompare(y.a.nome, 'pt-BR'));
const ehAssembleia = e => e && e.tipo === 'Assembleia';
function aptosVotar() { return ativosOrd().filter(a => (idade(a.nascimento) ?? 99) >= 18 && !financeiro(a).bloqueia); }
function eventoLinha(e) {
  const n = L('presencas').filter(p => p.eventoId === e.id).length;
  return `<button class="row-item linklike ${e.cancelado ? 'dim' : ''}" data-act="go" data-v="evento" data-id="${e.id}" style="width:100%;text-align:left">
    <span class="ev-data"><b>${fmtD(e.data).slice(0, 5)}</b><span>${DIAS_SEMANA[new Date(e.data + 'T12:00').getDay()].slice(0, 3)}</span></span>
    <span class="who"><b>${esc(e.titulo)}</b><span class="small muted">${e.ataNum ? `Ata ${esc(e.ataNum)} · ` : ''}${esc(e.tipo)}${e.subtipo && !ehAssembleia(e) ? ' · ' + esc(e.subtipo) : ''}${e.acesso ? ' · ' + esc(e.acesso.toLowerCase()) : ''}${e.hora ? ' · ' + e.hora : ''}${e.local ? ' · ' + esc(e.local) : ''}${e.cancelado ? ' · cancelado' : ''}</span></span>
    <span class="small muted" style="margin-left:auto;white-space:nowrap">${n ? plural(n, 'presente', 'presentes') : e.mediuns ? plural(+e.mediuns, 'médium', 'médiuns') : ''}${e.ata || e.ataArquivoId ? ' · ata' : ''}</span></button>`;
}
VIEWS.agenda = {
  mod: 'agenda',
  render() {
    const f = P.f || 'proximos', evs = L('eventos').filter(e => !ehAssembleia(e));
    const prox = evs.filter(e => e.data >= TODAY()).sort((x, y) => (x.data + x.hora).localeCompare(y.data + y.hora)), ant = evs.filter(e => e.data < TODAY()).sort((x, y) => (y.data + y.hora).localeCompare(x.data + x.hora));
    const semAta = ant.filter(e => !e.cancelado && !e.ata && !e.ataArquivoId && SESSOES.includes(e.tipo)).length;
    const ano = TODAY().slice(0, 4), doAno = ant.filter(e => !e.cancelado && e.data.startsWith(ano) && SESSOES.includes(e.tipo)), soma = k => doAno.reduce((t, e) => t + (Number(e[k]) || 0), 0);
    const tab = (id, label, n) => `<button class="tab ${f === id ? 'on' : ''}" data-act="agF" data-f="${id}">${label} <span class="count">${n}</span></button>`;
    const lista = f === 'proximos' ? prox : ant;
    return `<div class="page-head"><div><h1>Agenda</h1><p class="sub">Giras, festas, reuniões e cursos, com presença e ata.</p></div>
      <div class="toolbar"><button class="btn" data-act="freqAbrir">${ic('users')} Frequência</button><button class="btn primary" data-act="evNovo">${ic('plus')} Novo evento</button></div></div>
      ${semAta ? `<div class="note info"><span>${plural(semAta, 'gira já realizada está sem ata', 'giras já realizadas estão sem ata')}. Abra a sessão e registre ou anexe a ata.</span></div>` : ''}
      <div class="tabs" role="tablist">${tab('proximos', 'Próximos', prox.length)}${tab('anteriores', 'Anteriores', ant.length)}</div>
      ${f === 'anteriores' && doAno.length ? `<p class="small muted">${ano}: ${plural(doAno.length, 'sessão realizada', 'sessões realizadas')} · ${NUMS_SESSAO.map(([k, r]) => `${r.toLowerCase()}: <b>${soma(k)}</b>`).join(' · ')}</p>` : ''}
      <section class="panel">${lista.length ? `<div class="list">${lista.slice(0, 200).map(eventoLinha).join('')}</div>` : `<div class="empty">${f === 'proximos' ? 'Nada agendado. Use "Novo evento".' : 'Nenhum evento anterior.'}</div>`}</section>`;
  }
};
ACT.agF = t => { P.f = t.dataset.f; render(); };
VIEWS.assembleias = {
  mod: 'assembleias',
  render() {
    const evs = L('eventos').filter(ehAssembleia).sort((x, y) => y.data.localeCompare(x.data)), ap = aptosVotar().length;
    return `<div class="page-head"><div><h1>Assembleias</h1><p class="sub">${plural(ap, 'associado pode votar hoje', 'associados podem votar hoje')} (ativos, maiores de 18 anos e sem atraso acima da tolerância).</p></div>
      ${isAdmin() ? `<div class="toolbar"><button class="btn primary" data-act="evNovo" data-tipo="Assembleia">${ic('plus')} Nova assembleia</button></div>` : ''}</div>
      <section class="panel">${evs.length ? `<div class="list">${evs.map(eventoLinha).join('')}</div>` : '<div class="empty">Nenhuma assembleia registrada.</div>'}</section>`;
  }
};
function eventoForm(e, tipo) {
  const asm = (e ? e.tipo : tipo) === 'Assembleia';
  const sess = SESSOES.includes(e ? e.tipo : 'Gira') ? '' : 'hidden';
  return `<div class="fields">
    ${asm ? `<div class="field"><label for="e-sub">Tipo</label><select id="e-sub">${['Ordinária', 'Extraordinária'].map(t => `<option ${e && e.subtipo === t ? 'selected' : ''}>${t}</option>`).join('')}</select></div>`
      : `<div class="field"><label for="e-tipo">Sessão / tipo</label><select id="e-tipo" data-ch="evTipo">${TIPOS_EVENTO.map(t => `<option ${e && e.tipo === t ? 'selected' : ''}>${t}</option>`).join('')}</select></div>
         <div class="field"><label for="e-sub">Linha / tema</label><input id="e-sub" value="${esc(e ? e.subtipo : '')}" placeholder="Ex.: Quimbanda, Preto Velho, Ogum e Caboclo"></div>
         <div class="field e-sess" ${sess}><label for="e-acesso">Pública ou privada</label><select id="e-acesso"><option value="Pública" ${e && e.acesso === 'Privada' ? '' : 'selected'}>Pública (aberta)</option><option value="Privada" ${e && e.acesso === 'Privada' ? 'selected' : ''}>Privada (fechada)</option></select></div>
         <div class="field"><label for="e-atan">Nº da ata</label><input id="e-atan" type="number" min="1" inputmode="numeric" value="${esc(e ? e.ataNum : '')}"></div>`}
    <div class="field full"><label for="e-tit">Título</label><input id="e-tit" value="${esc(e ? e.titulo : '')}" placeholder="${asm ? 'Assembleia Geral Ordinária' : 'Opcional. Em branco: sessão + linha (ex.: Gira de Quimbanda)'}"></div>
    <div class="field"><label for="e-data">Data</label><input id="e-data" type="date" value="${e ? e.data : TODAY()}"></div>
    <div class="field"><label for="e-hora">Hora</label><input id="e-hora" type="time" value="${e ? e.hora : ''}"></div>
    <div class="field full"><label for="e-local">Local</label><input id="e-local" value="${esc(e ? e.local : '')}" placeholder="${esc(assocCfg().endereco || 'Terreiro')}"></div>
    ${asm ? `<div class="field full"><label for="e-pauta">Pauta (ordem do dia)</label><textarea id="e-pauta" rows="5" placeholder="1. Prestação de contas&#10;2. Eleição da diretoria">${esc(e ? e.pauta : '')}</textarea></div>`
      : `<div class="field"><label for="e-dir">Dirigente</label><input id="e-dir" value="${esc(e ? e.dirigente : '')}"></div>`}
    <div class="field full"><label for="e-desc">Observações</label><textarea id="e-desc" rows="2">${esc(e ? e.descricao : '')}</textarea></div>
    ${asm ? '' : `<fieldset class="field full e-sess nums" ${sess}><legend>Números da sessão (opcional)</legend><div class="nums-grid">${NUMS_SESSAO.map(([k, r]) => `<label>${r}<input id="e-n-${k}" type="number" min="0" inputmode="numeric" value="${esc(e ? e[k] : '')}"></label>`).join('')}</div></fieldset>`}</div>`;
}
const eventoDados = (tipoFixo) => { const d = { tipo: tipoFixo || $('#e-tipo').value, subtipo: $('#e-sub')?.value || '', titulo: $('#e-tit').value, data: $('#e-data').value, hora: $('#e-hora').value, local: $('#e-local').value, pauta: $('#e-pauta')?.value || '', dirigente: $('#e-dir')?.value || '', descricao: $('#e-desc').value, acesso: $('#e-acesso')?.value || '', ataNum: $('#e-atan')?.value || '' };
  NUMS_SESSAO.forEach(([k]) => { d[k] = $('#e-n-' + k)?.value || ''; }); return d; };
// mostra pública/privada e números só para Gira, Festa e Batismo
CHG.evTipo = el => document.querySelectorAll('.e-sess').forEach(x => { x.hidden = !SESSOES.includes(el.value); });
Object.assign(ACT, {
  evNovo(t) {
    const asm = t.dataset.tipo === 'Assembleia';
    openModal(`<h2>${asm ? 'Nova assembleia' : 'Novo evento'}</h2>${eventoForm(null, asm ? 'Assembleia' : '')}
      <div class="modal-actions"><button class="btn" data-act="close">Cancelar</button><button class="btn primary" data-act="evSalvar" data-tipo="${asm ? 'Assembleia' : ''}">Salvar</button></div>`, 'modal-wide');
  },
  evEditar(t) {
    const e = evento(t.dataset.id);
    openModal(`<h2>Editar ${ehAssembleia(e) ? 'assembleia' : 'evento'}</h2>${eventoForm(e)}
      <div class="modal-actions"><button class="btn" data-act="close">Cancelar</button><button class="btn primary" data-act="evSalvar" data-id="${e.id}" data-tipo="${ehAssembleia(e) ? 'Assembleia' : ''}">Salvar</button></div>`, 'modal-wide');
  },
  evSalvar(t) {
    const d = eventoDados(t.dataset.tipo); if (t.dataset.id) d.id = t.dataset.id;
    run(async () => { const r = await api('evento.salvar', d); closeModal(); toast(d.id ? 'Evento atualizado' : 'Evento criado'); go('evento', { id: r.id }); });
  },
  evCancelar(t) {
    openModal(`<h2>Cancelar evento</h2><div class="field"><label for="e-mot">Motivo</label><input id="e-mot" maxlength="300" placeholder="Ex.: chuva, falta de energia"></div>
      <div class="modal-actions"><button class="btn" data-act="close">Voltar</button><button class="btn danger" data-act="evCancelarOk" data-id="${t.dataset.id}">Cancelar evento</button></div>`);
  },
  evCancelarOk(t) { run(async () => { await api('evento.cancelar', { id: t.dataset.id, motivo: $('#e-mot').value }); closeModal(); render(); toast('Evento cancelado'); }); },
  presAdd(t) {
    const a = assocDoTexto($('#pr-q').value); if (!a) { toast('Escolha um associado da lista.', 'bad'); return; }
    run(async () => { const r = await api('presenca.marcar', { eventoId: t.dataset.id, assocId: a.id }); $('#pr-q').value = ''; render(); toast(r.ja ? `${r.nome} já estava na lista` : `Presença: ${r.nome}`); setTimeout(() => $('#pr-q')?.focus(), 50); });
  },
  presRem(t) { run(async () => { await api('presenca.remover', { id: t.dataset.id }); render(); }); },
  presVarios(t) {
    const e = evento(t.dataset.id), ja = new Set(L('presencas').filter(p => p.eventoId === e.id).map(p => p.assocId)), lista = ativosOrd().filter(a => !ja.has(a.id));
    openModal(`<h2>Marcar presença</h2><p class="small muted">${esc(e.titulo)} · ${fmtD(e.data)}. Marque quem esteve presente.</p>
      <label class="search"><span class="sr">Filtrar</span>${ic('search')}<input data-in="presFiltro" placeholder="Filtrar por nome" autocomplete="off"></label>
      <div class="pres-lista">${lista.map(a => `<label class="pres-item" data-nome="${esc(norm(a.nome))}"><input type="checkbox" value="${a.id}">${avatar(a)}<span>${esc(a.nome)}</span></label>`).join('') || '<p class="small muted">Todos os ativos já estão na lista.</p>'}</div>
      <div class="modal-actions"><button class="btn" data-act="close">Cancelar</button><button class="btn primary" data-act="presVariosOk" data-id="${e.id}">Marcar selecionados</button></div>`, 'modal-wide');
  },
  presVariosOk(t) {
    const ids = $$('.pres-lista input:checked').map(x => x.value); if (!ids.length) { toast('Marque ao menos uma pessoa.', 'bad'); return; }
    closeModal();
    run(async () => { let n = 0; for (const id of ids) { await api('presenca.marcar', { eventoId: t.dataset.id, assocId: id }, { msg: `Marcando ${n + 1} de ${ids.length}…` }); n++; } render(); toast(`${plural(n, 'presença marcada', 'presenças marcadas')}`); });
  },
  presQR(t) {
    UI.qr = { eventoId: t.dataset.id, ultimo: '', em: 0 };
    const temDetector = 'BarcodeDetector' in window && navigator.mediaDevices && navigator.mediaDevices.getUserMedia;
    openModal(`<h2>Ler QR da carteirinha</h2>
      ${temDetector ? '<div class="qr-cam"><video id="qr-video" playsinline muted></video></div><p class="small muted" id="qr-msg">Aponte a câmera para o QR Code da carteirinha (impressa ou no celular).</p>'
        : '<div class="note info"><span>Este navegador não lê QR Code pela câmera. Digite o código impresso abaixo do QR (8 letras e números).</span></div>'}
      <div class="pres-add"><input id="qr-cod" class="mono" placeholder="Código da carteirinha (ex.: ABCD-1234)" autocomplete="off"><button class="btn" data-act="presCodigo">Marcar</button></div>
      <div id="qr-log" class="small"></div>
      <div class="modal-actions"><button class="btn primary" data-act="qrFechar">Concluir</button></div>`);
    if (temDetector) qrIniciar();
  },
  presCodigo() { const c = $('#qr-cod').value.trim(); if (!c) return; qrMarcar(c); $('#qr-cod').value = ''; },
  qrFechar() { qrParar(); closeModal(); render(); },
  ataAbrir(t) {
    const e = evento(t.dataset.id), n = presentes(e.id).length, asm = ehAssembleia(e), ap = asm ? aptosVotar() : [];
    const nAp = asm ? presentes(e.id).filter(x => ap.some(a => a.id === x.a.id)).length : 0;
    const modelo = asm
      ? `Ata da ${e.titulo} da ${assocCfg().nome}, realizada em ${fmtD(e.data)}${e.hora ? ', às ' + e.hora : ''}${e.local ? ', em ' + e.local : ''}, com a presença de ${n} associados(as), sendo ${nAp} aptos(as) a votar de um total de ${ap.length}.\n\nPauta:\n${e.pauta || ''}\n\nDeliberações:\n\n\nNada mais havendo a tratar, a assembleia foi encerrada e eu, secretário(a), lavrei a presente ata.`
      : `Ata da ${e.titulo}, realizada em ${fmtD(e.data)}${e.hora ? ', às ' + e.hora : ''}${e.local ? ', em ' + e.local : ''}.\nDirigente: ${e.dirigente || ''}\nPresentes: ${n} associados(as) (lista anexa).\n\nAbertura:\n\nAtendimentos e ocorrências:\n\nAvisos e encaminhamentos:\n\nEncerramento:\n`;
    openModal(`<h2>Ata · ${esc(e.titulo)}</h2><p class="small muted">${fmtD(e.data)} · ${plural(n, 'presente', 'presentes')}</p>
      ${asm ? '' : `<div class="field"><label for="ata-dir">Dirigente</label><input id="ata-dir" value="${esc(e.dirigente || '')}"></div>`}
      <div class="field"><label for="ata-txt">Texto da ata</label><textarea id="ata-txt" rows="14">${esc(e.ata || modelo)}</textarea><span class="hint">Depois de salvar, dá para imprimir para assinatura e anexar a ata assinada.</span></div>
      <div class="modal-actions"><button class="btn" data-act="close">Cancelar</button><button class="btn primary" data-act="ataSalvar" data-id="${e.id}">Salvar ata</button></div>`, 'modal-wide');
  },
  ataSalvar(t) { const p = { id: t.dataset.id, texto: $('#ata-txt').value }; if ($('#ata-dir')) p.dirigente = $('#ata-dir').value; run(async () => { await api('evento.ata', p); closeModal(); render(); toast('Ata salva'); }); },
  ataImprimir(t) {
    const e = evento(t.dataset.id), pr = presentes(e.id), asm = ehAssembleia(e);
    printHTML(`<div class="doc">${docCab(`Ata · ${e.titulo} · ${fmtD(e.data)}`)}<div class="ata-txt">${esc(e.ata)}</div>
      <h3>Presentes (${pr.length})</h3><p class="small">${pr.map(x => `${fmtMat(x.a.matricula)} ${esc(x.a.nome)}`).join(' · ') || '—'}</p>
      <div class="fass-doc">${(asm ? ['Presidente', 'Secretário(a)'] : ['Dirigente' + (e.dirigente ? ': ' + esc(e.dirigente) : ''), 'Secretário(a)']).map(r => `<div><span></span>${r}</div>`).join('')}</div></div>`, 'doc-print');
  },
  ataAnexar(t) {
    openModal(`<h2>Anexar ata</h2><div class="field"><label for="ata-arq">Arquivo: PDF, foto ou Word</label><input id="ata-arq" type="file" accept="image/*,application/pdf,.doc,.docx"><span class="hint">Ata já digitada ou assinada e escaneada. Até 6 MB.</span></div>
      <div class="modal-actions"><button class="btn" data-act="close">Cancelar</button><button class="btn primary" data-act="ataAnexarOk" data-id="${t.dataset.id}">Anexar</button></div>`);
  },
  ataAnexarOk(t) { const f = $('#ata-arq').files[0]; if (!f) { toast('Escolha o arquivo.', 'bad'); return; } run(async () => { const prep = await prepAta(f); await api('evento.ataArquivo', { id: t.dataset.id, ...prep }, { msg: 'Enviando ata…' }); closeModal(); render(); toast('Ata anexada'); }); },
  ataVer(t) { run(async () => { const r = await api('evento.ataObter', { id: t.dataset.id }, { msg: 'Abrindo ata…' }); mostrarArquivo(r, 'Ata'); }); },
  listaPresenca(t) {
    const e = evento(t.dataset.id), asm = ehAssembleia(e), lista = asm ? aptosVotar() : ativosOrd();
    printHTML(`<div class="doc">${docCab(`Lista de presença · ${e.titulo} · ${fmtD(e.data)}${e.hora ? ' ' + e.hora : ''}`)}
      ${asm ? `<p class="small">Associados aptos a votar: ${lista.length} (ativos, maiores de 18 anos e sem atraso acima da tolerância).</p>` : ''}
      <table class="doc-t cols"><thead><tr><th style="width:12%">Matrícula</th><th>Nome</th><th style="width:40%">Assinatura</th></tr></thead><tbody>
      ${lista.map(a => `<tr><td>${fmtMat(a.matricula)}</td><td>${esc(a.nome)}</td><td></td></tr>`).join('')}</tbody></table></div>`, 'doc-print');
  },
  edital(t) {
    const e = evento(t.dataset.id), c = S.config;
    printHTML(`<div class="doc">${docCab('Edital de convocação')}<h3 style="text-align:center">EDITAL DE CONVOCAÇÃO · ${esc(e.titulo.toUpperCase())}</h3>
      <p style="text-align:justify">A Presidência da ${esc(assocCfg().nome)}${assocCfg().cnpj ? `, CNPJ ${fmtCNPJ(assocCfg().cnpj)},` : ''} no uso de suas atribuições estatutárias, convoca os(as) associados(as) para a <b>${esc(e.titulo)}</b>, a realizar-se em <b>${fmtD(e.data)}</b>${e.hora ? `, às <b>${e.hora}</b>` : ''}${e.local ? `, em <b>${esc(e.local)}</b>` : ''}, para deliberar sobre a seguinte ordem do dia:</p>
      <div class="ata-txt">${esc(e.pauta || '')}</div>
      <p>Podem votar os(as) associados(as) em dia com suas obrigações, conforme o estatuto.</p>
      <p>${esc(c.cidade || '')}, ${fmtD(TODAY())}.</p><div class="fass-doc"><div><span></span>${esc(c.presidenteNome || 'Presidente')}<br>Presidente</div></div></div>`, 'doc-print');
  },
  freqAbrir() {
    const de = P.fde || addDays(TODAY(), -90), ate = P.fate || TODAY();
    openModal(`<h2>Frequência nas giras</h2><div class="fields"><div class="field"><label for="fq-de">De</label><input id="fq-de" type="date" value="${de}"></div><div class="field"><label for="fq-ate">Até</label><input id="fq-ate" type="date" value="${ate}"></div>
      <div class="field"><label for="fq-tipo">Eventos</label><select id="fq-tipo"><option value="Gira">Só giras</option><option value="">Todos os tipos</option></select></div></div>
      <div class="modal-actions"><button class="btn" data-act="close">Fechar</button><button class="btn primary" data-act="freqImprimir">${ic('print')} Ver e imprimir</button></div>`);
  },
  freqImprimir() {
    const de = $('#fq-de').value, ate = $('#fq-ate').value, tipo = $('#fq-tipo').value;
    const evs = L('eventos').filter(e => !e.cancelado && !ehAssembleia(e) && e.data >= de && e.data <= ate && (!tipo || e.tipo === tipo));
    const linhas = ativosOrd().map(a => { const n = evs.filter(e => L('presencas').some(p => p.eventoId === e.id && p.assocId === a.id)).length; return { a, n, pct: evs.length ? Math.round(n / evs.length * 100) : 0 }; }).sort((x, y) => y.n - x.n || x.a.nome.localeCompare(y.a.nome, 'pt-BR'));
    closeModal();
    printHTML(`<div class="doc">${docCab(`Frequência ${tipo ? 'nas giras' : 'nos eventos'} · ${fmtD(de)} a ${fmtD(ate)} · ${plural(evs.length, 'evento', 'eventos')}`)}
      <table class="doc-t cols"><thead><tr><th style="width:12%">Matrícula</th><th>Nome</th><th style="width:16%">Presenças</th><th style="width:12%">%</th></tr></thead><tbody>
      ${linhas.map(x => `<tr><td>${fmtMat(x.a.matricula)}</td><td>${esc(x.a.nome)}</td><td>${x.n} de ${evs.length}</td><td>${x.pct}%</td></tr>`).join('')}</tbody></table></div>`, 'doc-print');
  },
});
INP.presFiltro = el => { const q = norm(el.value); $$('.pres-item').forEach(x => { x.hidden = q && !x.dataset.nome.includes(q); }); };
async function qrIniciar() {
  const v = $('#qr-video'); if (!v) return;
  try {
    UI.qr.stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' }, audio: false });
    v.srcObject = UI.qr.stream; await v.play();
    const det = new BarcodeDetector({ formats: ['qr_code'] });
    const passo = async () => {
      if (!UI.qr || !UI.qr.stream || !document.body.contains(v)) return;
      try { const r = await det.detect(v); if (r.length) { const c = r[0].rawValue; if (c !== UI.qr.ultimo || Date.now() - UI.qr.em > 4000) { UI.qr.ultimo = c; UI.qr.em = Date.now(); qrMarcar(c); } } } catch (e) { }
      UI.qr.raf = requestAnimationFrame(passo);
    };
    passo();
  } catch (e) { const m = $('#qr-msg'); if (m) m.textContent = 'Não foi possível abrir a câmera. Digite o código impresso abaixo do QR.'; }
}
function qrParar() { if (UI.qr) { if (UI.qr.raf) cancelAnimationFrame(UI.qr.raf); if (UI.qr.stream) UI.qr.stream.getTracks().forEach(t => t.stop()); UI.qr.stream = null; } }
async function qrMarcar(codigo) {
  const log = $('#qr-log');
  try { const r = await api('presenca.marcar', { eventoId: UI.qr.eventoId, codigo }, { silent: true }); if (log) log.insertAdjacentHTML('afterbegin', `<div class="${r.ja ? 'muted' : 'ok-text'}">${ic('check')} ${esc(r.nome)}${r.ja ? ' (já estava na lista)' : ''}</div>`); if (navigator.vibrate) navigator.vibrate(80); }
  catch (e) { if (log) log.insertAdjacentHTML('afterbegin', `<div class="bad-text">${esc(e.message)}</div>`); }
}
const _closeModal = closeModal;
VIEWS.evento = {
  render() {
    const e = evento(P.id); if (!e) return '<div class="empty">Evento não encontrado.</div>';
    if (!modOn(ehAssembleia(e) ? 'assembleias' : 'agenda')) return '<div class="empty">Este módulo está desligado em Configurações.</div>';
    const asm = ehAssembleia(e), pr = presentes(e.id), podeEd = !asm || isAdmin(), ap = asm ? aptosVotar() : [];
    const prAptos = asm ? pr.filter(x => ap.some(a => a.id === x.a.id)).length : 0;
    return `<div><button class="back" data-act="go" data-v="${asm ? 'assembleias' : 'agenda'}">${ic('back')} ${asm ? 'Assembleias' : 'Agenda'}</button>
      <div class="page-head"><div><h1>${esc(e.titulo)}</h1><p class="sub">${e.ataNum ? `Ata ${esc(e.ataNum)} · ` : ''}${esc(e.tipo)}${e.subtipo ? ' · ' + esc(e.subtipo) : ''}${e.acesso ? ' · ' + esc(e.acesso) : ''} · ${fmtD(e.data)}${e.hora ? ' às ' + e.hora : ''}${e.local ? ' · ' + esc(e.local) : ''}${e.dirigente ? ' · dirigente: ' + esc(e.dirigente) : ''}</p></div>
      ${podeEd && !e.cancelado ? `<div class="toolbar"><button class="btn sm" data-act="evEditar" data-id="${e.id}">${ic('edit')} Editar</button>${!asm && modOn('documentos') ? `<button class="btn sm" data-act="evTermo" data-id="${e.id}">${ic('file')} Termo de consentimento</button>` : ''}<button class="btn sm ghost danger-text" data-act="evCancelar" data-id="${e.id}">${ic('x')} Cancelar</button></div>` : ''}</div></div>
      ${e.cancelado ? `<div class="note bad"><span>Evento cancelado: ${esc(e.motivoCancelamento)}</span></div>` : ''}
      ${e.descricao ? `<p class="small">${esc(e.descricao)}</p>` : ''}
      ${NUMS_SESSAO.some(([k]) => e[k] !== '' && e[k] != null) ? `<section class="panel"><div class="panel-head"><h2>Números da sessão</h2></div><div class="kpis">${NUMS_SESSAO.filter(([k]) => e[k] !== '' && e[k] != null).map(([k, r]) => kpiBox(r, e[k], '')).join('')}</div></section>` : ''}
      ${asm ? `<section class="panel"><div class="panel-head"><h2>Quórum</h2><div class="toolbar"><button class="btn sm" data-act="edital" data-id="${e.id}">${ic('print')} Edital de convocação</button><button class="btn sm" data-act="listaPresenca" data-id="${e.id}">${ic('print')} Lista para assinar</button></div></div>
        <div class="kpis">${kpiBox('Podem votar', ap.length, 'ativos, 18+ e em dia')}${kpiBox('Presentes aptos', prAptos, ap.length ? `${Math.round(prAptos / ap.length * 100)}% dos aptos` : '—', ap.length && prAptos * 2 > ap.length ? 'ok' : '')}${kpiBox('Presentes', pr.length, 'total registrado')}</div>
        ${e.pauta ? `<h3 class="small" style="margin:8px 0 4px">Pauta</h3><div class="ata-txt small">${esc(e.pauta)}</div>` : ''}
        <p class="small muted">Confira no estatuto o quórum exigido em primeira e segunda chamada.</p></section>` : ''}
      <section class="panel"><div class="panel-head"><h2>Presença <span class="count">${pr.length}</span></h2>
        ${!e.cancelado ? `<div class="toolbar"><button class="btn sm primary" data-act="presQR" data-id="${e.id}">${ic('qr')} Ler QR da carteirinha</button><button class="btn sm" data-act="presVarios" data-id="${e.id}">${ic('check')} Marcar vários</button>${asm ? '' : `<button class="btn sm" data-act="listaPresenca" data-id="${e.id}">${ic('print')} Lista para assinar</button>`}</div>` : ''}</div>
        ${!e.cancelado ? `<div class="pres-add"><input id="pr-q" list="m-alist" placeholder="Digite o nome ou a matrícula" autocomplete="off">${assocOpts()}<button class="btn" data-act="presAdd" data-id="${e.id}">${ic('plus')} Marcar</button></div>` : ''}
        ${pr.length ? `<div class="pres-grid">${pr.map(x => `<span class="pres-chip">${avatar(x.a)}<span>${esc(x.a.nome)}${x.p.modo === 'Carteirinha' ? ` <span class="muted small">${ic('qr')}</span>` : ''}</span>${!e.cancelado && podeEd ? `<button class="btn sm ghost icon" data-act="presRem" data-id="${x.p.id}" aria-label="Remover ${esc(x.a.nome)}" title="Remover">${ic('x')}</button>` : ''}</span>`).join('')}</div>` : '<p class="small muted">Ninguém marcado ainda.</p>'}
      </section>
      <section class="panel"><div class="panel-head"><h2>Ata</h2>${podeEd && !e.cancelado ? `<div class="toolbar">${e.ata ? `<button class="btn sm" data-act="ataImprimir" data-id="${e.id}">${ic('print')} Imprimir</button>` : ''}<button class="btn sm ${e.ata || e.ataArquivoId ? '' : 'primary'}" data-act="ataAbrir" data-id="${e.id}">${ic('edit')} ${e.ata ? 'Editar ata' : 'Digitar ata'}</button><button class="btn sm" data-act="ataAnexar" data-id="${e.id}">${ic('clip')} ${e.ataArquivoId ? 'Trocar ata anexada' : 'Anexar ata'}</button></div>` : ''}</div>
        ${e.ata ? `<div class="ata-txt">${esc(e.ata)}</div><p class="small muted">Registrada por ${esc(e.ataPor)} em ${fmtStamp(e.ataEm)}.${e.ataArquivoId ? ` <button class="btn sm ghost" data-act="ataVer" data-id="${e.id}">${ic('file')} Abrir ata anexada</button>` : ''}</p>`
          : e.ataArquivoId ? `<p class="small">Ata anexada: ${esc(e.ataArquivoNome)} <button class="btn sm" data-act="ataVer" data-id="${e.id}">${ic('file')} Abrir</button></p>`
          : `<p class="small muted">${e.data <= TODAY() ? 'Digite a ata aqui ou anexe a ata já digitada (PDF, foto ou Word).' : 'A ata pode ser registrada depois do evento.'}</p>`}
      </section>`;
  }
};
PARENT.evento = 'agenda';
// Fecha a câmera se o modal for fechado por outro caminho
window.closeModal = function () { qrParar(); _closeModal(); };

/* ============ Caminhada na casa ============ */
Object.assign(ACT, {
  marcoNovo(t) {
    const a = assoc(t.dataset.id);
    openModal(`<h2>Caminhada na casa</h2><p class="muted">${esc(a.nome)}</p><div class="fields">
      <div class="field"><label for="mc-tipo">Etapa</label><select id="mc-tipo">${TIPOS_MARCO.map(x => `<option>${x}</option>`).join('')}</select></div>
      <div class="field"><label for="mc-data">Data</label><input id="mc-data" type="date" value="${TODAY()}" max="${TODAY()}"></div>
      <div class="field full"><label for="mc-desc">Detalhes</label><input id="mc-desc" maxlength="500" placeholder="Ex.: orixá de cabeça, quem conduziu, nome do curso"></div>
      <div class="field full"><label for="mc-arq">Certidão (opcional)</label><input id="mc-arq" type="file" accept="image/*,application/pdf"><span class="hint">Certidão de batismo, de casamento ou outro comprovante. Foto ou PDF.</span></div></div>
      <div class="modal-actions"><button class="btn" data-act="close">Cancelar</button><button class="btn primary" data-act="marcoSalvar" data-id="${a.id}">Registrar</button></div>`);
  },
  marcoSalvar(t) { const f = $('#mc-arq').files[0]; run(async () => { const prep = f ? await prepComprovante(f, 2400, 0.85) : {}; await api('marco.salvar', { assocId: t.dataset.id, tipo: $('#mc-tipo').value, data: $('#mc-data').value, descricao: $('#mc-desc').value, ...prep }, { msg: 'Registrando…' }); closeModal(); render(); toast('Registrado na caminhada'); }); },
  marcoAnexar(t) {
    openModal(`<h2>Anexar certidão</h2><div class="field"><label for="mc-arq">Arquivo (foto ou PDF)</label><input id="mc-arq" type="file" accept="image/*,application/pdf"></div>
      <div class="modal-actions"><button class="btn" data-act="close">Cancelar</button><button class="btn primary" data-act="marcoAnexarOk" data-id="${t.dataset.id}">Anexar</button></div>`);
  },
  marcoAnexarOk(t) { const f = $('#mc-arq').files[0]; if (!f) { toast('Escolha o arquivo.', 'bad'); return; } run(async () => { const prep = await prepComprovante(f, 2400, 0.85); await api('marco.anexar', { id: t.dataset.id, ...prep }, { msg: 'Enviando certidão…' }); closeModal(); render(); toast('Certidão anexada'); }); },
  marcoVer(t) { const m = L('marcos').find(x => x.id === t.dataset.id); run(async () => { const r = await api('marco.arquivo', { id: t.dataset.id }, { msg: 'Abrindo certidão…' }); mostrarArquivo(r, `Certidão${m ? ' · ' + m.tipo : ''}`); }); },
  marcoExcluir(t) { run(async () => { await api('marco.excluir', { id: t.dataset.id }); render(); toast('Registro excluído'); }); },
});

/* ============ Campanhas e doações ============ */
const campanha = id => L('campanhas').find(c => c.id === id);
const contribs = cid => L('contribuicoes').filter(x => x.campanhaId === cid && !x.cancelado);
function campanhaTotais(c) { const l = contribs(c.id); const valor = l.reduce((s, x) => s + (x.tipo === 'Dinheiro' ? x.valor : 0), 0), itens = l.reduce((s, x) => s + (x.tipo === 'Item' ? x.quantidade : 0), 0); const atual = c.tipoMeta === 'Valor' ? valor : c.tipoMeta === 'Itens' ? itens : 0; return { valor: Math.round(valor * 100) / 100, itens, n: l.length, pct: c.meta ? Math.min(100, Math.round(atual / c.meta * 100)) : null }; }
const barra = pct => pct == null ? '' : `<div class="barra" role="progressbar" aria-valuenow="${pct}" aria-valuemin="0" aria-valuemax="100"><span style="width:${pct}%"></span></div>`;
VIEWS.campanhas = {
  mod: 'campanhas',
  render() {
    const l = L('campanhas').slice().sort((x, y) => (y.ativa - x.ativa) || y.inicio.localeCompare(x.inicio));
    return `<div class="page-head"><div><h1>Campanhas</h1><p class="sub">Arrecadações da casa: cestas básicas, festas, reformas.</p></div>${canFin() ? `<div class="toolbar"><button class="btn primary" data-act="cmpNova">${ic('plus')} Nova campanha</button></div>` : ''}</div>
      <section class="panel">${l.length ? `<div class="list">${l.map(c => { const t = campanhaTotais(c); return `<button class="row-item linklike ${c.ativa ? '' : 'dim'}" data-act="go" data-v="campanha" data-id="${c.id}" style="width:100%;text-align:left;display:grid;gap:4px">
        <span class="who"><b>${esc(c.nome)}</b><span class="small muted">${c.ativa ? 'Ativa' : 'Encerrada'} · desde ${fmtD(c.inicio)}${c.fim ? ' até ' + fmtD(c.fim) : ''} · ${plural(t.n, 'contribuição', 'contribuições')}${t.valor ? ' · ' + fmtBRL(t.valor) : ''}${t.itens ? ` · ${num(t.itens)} ${esc(c.unidade || 'itens')}` : ''}${c.meta ? ` · meta ${c.tipoMeta === 'Valor' ? fmtBRL(c.meta) : num(c.meta) + ' ' + esc(c.unidade || 'itens')}` : ''}</span></span>${barra(t.pct)}</button>`; }).join('')}</div>`
        : `<div class="empty">Nenhuma campanha.${canFin() ? ' Use "Nova campanha".' : ''}</div>`}</section>`;
  }
};
VIEWS.campanha = {
  mod: 'campanhas',
  render() {
    const c = campanha(P.id); if (!c) return '<div class="empty">Campanha não encontrada.</div>';
    const t = campanhaTotais(c), l = L('contribuicoes').filter(x => x.campanhaId === c.id).sort((x, y) => y.data.localeCompare(x.data) || y.registradoEm.localeCompare(x.registradoEm));
    return `<div><button class="back" data-act="go" data-v="campanhas">${ic('back')} Campanhas</button>
      <div class="page-head"><div><h1>${esc(c.nome)}</h1><p class="sub">${c.ativa ? 'Ativa' : 'Encerrada'} · desde ${fmtD(c.inicio)}${c.fim ? ' até ' + fmtD(c.fim) : ''}</p></div>
      <div class="toolbar"><button class="btn sm" data-act="cmpImprimir" data-id="${c.id}">${ic('print')} Imprimir</button>${canFin() ? `<button class="btn sm" data-act="cmpEditar" data-id="${c.id}">${ic('edit')} Editar</button><button class="btn sm" data-act="cmpAtiva" data-id="${c.id}">${c.ativa ? 'Encerrar' : 'Reabrir'}</button>` : ''}${c.ativa ? `<button class="btn sm primary" data-act="ctbNova" data-id="${c.id}">${ic('plus')} Registrar contribuição</button>` : ''}</div></div></div>
      ${c.descricao ? `<p class="small">${esc(c.descricao)}</p>` : ''}
      <div class="kpis">${kpiBox('Arrecadado', fmtBRL(t.valor), 'em dinheiro', 'money')}${kpiBox('Itens', num(t.itens), esc(c.unidade || 'unidades'))}${kpiBox('Contribuições', t.n, 'registradas')}${c.meta ? kpiBox('Meta', c.tipoMeta === 'Valor' ? fmtBRL(c.meta) : num(c.meta) + ' ' + esc(c.unidade || ''), `${t.pct}% alcançado`, t.pct >= 100 ? 'ok' : '') : ''}</div>
      ${barra(t.pct)}
      <section class="panel"><div class="panel-head"><h2>Contribuições</h2></div>
        ${l.length ? `<div class="table-wrap cards"><table><thead><tr><th>Quem</th><th>Contribuição</th><th>Data</th><th><span class="sr">Ações</span></th></tr></thead><tbody>
        ${l.map(x => `<tr class="${x.cancelado ? 'dim' : ''}"><td class="lead"><b>${esc(x.pessoa)}</b></td><td data-label="Contribuição">${x.tipo === 'Dinheiro' ? fmtBRL(x.valor) : `${num(x.quantidade)} ${esc(x.descricao)}`}${x.tipo === 'Dinheiro' && x.descricao ? `<div class="s">${esc(x.descricao)}</div>` : ''}${x.cancelado ? `<div class="s" style="color:var(--bad)">Cancelada: ${esc(x.motivoCancelamento)}</div>` : ''}</td>
          <td data-label="Data" class="num">${fmtD(x.data)}</td><td class="actions">${!x.cancelado && (canFin() || (x.tipo === 'Item' && x.registradoPor === S.session.email)) ? `<button class="btn sm ghost" data-act="ctbCancelar" data-id="${x.id}">Cancelar</button>` : ''}</td></tr>`).join('')}</tbody></table></div>` : '<div class="empty">Nenhuma contribuição ainda.</div>'}
      </section>`;
  }
};
PARENT.campanha = 'campanhas';
function campanhaForm(c) {
  return `<div class="fields"><div class="field full"><label for="cp-nome">Nome</label><input id="cp-nome" value="${esc(c ? c.nome : '')}" placeholder="Ex.: Cestas de Natal"></div>
    <div class="field full"><label for="cp-desc">Descrição</label><textarea id="cp-desc" rows="2">${esc(c ? c.descricao : '')}</textarea></div>
    <div class="field"><label for="cp-tm">Meta</label><select id="cp-tm"><option value="">Sem meta</option><option value="Valor" ${c && c.tipoMeta === 'Valor' ? 'selected' : ''}>Em dinheiro</option><option value="Itens" ${c && c.tipoMeta === 'Itens' ? 'selected' : ''}>Em itens</option></select></div>
    <div class="field"><label for="cp-meta">Quanto</label><input id="cp-meta" type="number" min="0" step="0.01" value="${c && c.meta ? c.meta : ''}"></div>
    <div class="field"><label for="cp-un">Unidade dos itens</label><input id="cp-un" value="${esc(c ? c.unidade : '')}" placeholder="Ex.: cestas, kg, peças"></div>
    <div class="field"><label for="cp-ini">Início</label><input id="cp-ini" type="date" value="${c ? c.inicio : TODAY()}"></div>
    <div class="field"><label for="cp-fim">Término (opcional)</label><input id="cp-fim" type="date" value="${c ? c.fim : ''}"></div></div>`;
}
const campanhaDados = () => ({ nome: $('#cp-nome').value, descricao: $('#cp-desc').value, tipoMeta: $('#cp-tm').value, meta: $('#cp-meta').value, unidade: $('#cp-un').value, inicio: $('#cp-ini').value, fim: $('#cp-fim').value });
Object.assign(ACT, {
  cmpNova() { openModal(`<h2>Nova campanha</h2>${campanhaForm(null)}<div class="modal-actions"><button class="btn" data-act="close">Cancelar</button><button class="btn primary" data-act="cmpSalvar">Criar campanha</button></div>`, 'modal-wide'); },
  cmpEditar(t) { const c = campanha(t.dataset.id); openModal(`<h2>Editar campanha</h2>${campanhaForm(c)}<div class="modal-actions"><button class="btn" data-act="close">Cancelar</button><button class="btn primary" data-act="cmpSalvar" data-id="${c.id}">Salvar</button></div>`, 'modal-wide'); },
  cmpSalvar(t) { const d = campanhaDados(); if (t.dataset.id) { d.id = t.dataset.id; d.ativa = campanha(d.id).ativa; } run(async () => { const r = await api('campanha.salvar', d); closeModal(); toast('Campanha salva'); go('campanha', { id: r.id }); }); },
  cmpAtiva(t) { const c = campanha(t.dataset.id); run(async () => { await api('campanha.salvar', { ...c, ativa: !c.ativa }); render(); toast(c.ativa ? 'Campanha encerrada' : 'Campanha reaberta'); }); },
  ctbNova(t) {
    const c = campanha(t.dataset.id);
    openModal(`<h2>Registrar contribuição</h2><p class="muted">${esc(c.nome)}</p><div class="fields">
      <div class="field"><label for="ct-tipo">Tipo</label><select id="ct-tipo" data-ch="ctbTipo"><option value="Item">Itens (alimentos, materiais…)</option>${canFin() ? '<option value="Dinheiro">Dinheiro</option>' : ''}</select>${canFin() ? '' : '<span class="hint">Doações em dinheiro são registradas pela tesouraria.</span>'}</div>
      <div class="field"><label for="ct-quem">Quem doou</label><input id="ct-quem" list="m-alist" placeholder="Associado ou nome de quem doou" autocomplete="off">${assocOpts('m-alist', S.associados.slice().sort((x, y) => x.nome.localeCompare(y.nome, 'pt-BR')))}<span class="hint">Deixe em branco para doação anônima.</span></div>
      <div class="field ct-item"><label for="ct-qtd">Quantidade</label><input id="ct-qtd" type="number" min="0" step="0.001" value="1"></div>
      <div class="field ct-din" hidden><label for="ct-val">Valor (R$)</label><input id="ct-val" type="number" min="0" step="0.01"></div>
      <div class="field ct-din" hidden><label for="ct-forma">Forma</label><select id="ct-forma">${['PIX', 'Dinheiro', 'Transferência', 'Cartão', 'Outro'].map(f => `<option>${f}</option>`).join('')}</select></div>
      <div class="field full"><label for="ct-desc">O que foi doado / observação</label><input id="ct-desc" maxlength="300" placeholder="Ex.: cestas básicas, 5 kg de arroz"></div>
      <div class="field"><label for="ct-data">Data</label><input id="ct-data" type="date" value="${TODAY()}" max="${TODAY()}"></div></div>
      <div class="modal-actions"><button class="btn" data-act="close">Cancelar</button><button class="btn primary" data-act="ctbSalvar" data-id="${c.id}">Registrar</button></div>`, 'modal-wide');
  },
  ctbSalvar(t) {
    const quem = $('#ct-quem').value.trim(), a = quem ? assocDoTexto(quem) : null, tipo = $('#ct-tipo').value;
    run(async () => { await api('contribuicao.registrar', { campanhaId: t.dataset.id, tipo, assocId: a ? a.id : '', pessoa: a ? '' : quem, quantidade: $('#ct-qtd').value, valor: $('#ct-val').value, forma: $('#ct-forma').value, descricao: $('#ct-desc').value, data: $('#ct-data').value }); closeModal(); render(); toast(tipo === 'Dinheiro' ? 'Doação registrada e lançada no livro-caixa' : 'Doação registrada'); });
  },
  ctbCancelar(t) {
    openModal(`<h2>Cancelar contribuição</h2><div class="field"><label for="ct-mot">Motivo</label><input id="ct-mot" maxlength="300" placeholder="Ex.: lançada em dobro"></div><p class="small muted">Se for em dinheiro, o lançamento no livro-caixa também é cancelado.</p>
      <div class="modal-actions"><button class="btn" data-act="close">Voltar</button><button class="btn danger" data-act="ctbCancelarOk" data-id="${t.dataset.id}">Cancelar contribuição</button></div>`);
  },
  ctbCancelarOk(t) { run(async () => { await api('contribuicao.cancelar', { id: t.dataset.id, motivo: $('#ct-mot').value }); closeModal(); render(); toast('Contribuição cancelada'); }); },
  cmpImprimir(t) {
    const c = campanha(t.dataset.id), tt = campanhaTotais(c), l = contribs(c.id).sort((x, y) => x.data.localeCompare(y.data));
    printHTML(`<div class="doc">${docCab(`Campanha · ${c.nome}`)}<p>${esc(c.descricao || '')}</p><p><b>Arrecadado:</b> ${fmtBRL(tt.valor)} em dinheiro${tt.itens ? ` e ${num(tt.itens)} ${esc(c.unidade || 'itens')}` : ''}${c.meta ? ` · meta ${c.tipoMeta === 'Valor' ? fmtBRL(c.meta) : num(c.meta) + ' ' + esc(c.unidade || '')} (${tt.pct}%)` : ''}.</p>
      <table class="doc-t cols"><thead><tr><th style="width:14%">Data</th><th>Quem</th><th>Contribuição</th></tr></thead><tbody>${l.map(x => `<tr><td>${fmtD(x.data)}</td><td>${esc(x.pessoa)}</td><td>${x.tipo === 'Dinheiro' ? fmtBRL(x.valor) : `${num(x.quantidade)} ${esc(x.descricao)}`}</td></tr>`).join('')}</tbody></table></div>`, 'doc-print');
  },
});
CHG.ctbTipo = el => { const din = el.value === 'Dinheiro'; $$('.ct-din').forEach(x => { x.hidden = !din; }); $$('.ct-item').forEach(x => { x.hidden = din; }); };

/* ============ Estoque e patrimônio ============ */
const item = id => L('itens').find(i => i.id === id);
const estoqueBaixo = () => L('itens').filter(i => i.ativo && i.tipo === 'Consumo' && i.minimo > 0 && i.quantidade <= i.minimo);
VIEWS.estoque = {
  mod: 'estoque',
  render() {
    const f = P.f || 'Consumo', q = norm(P.q || ''), inat = !!P.inat;
    const l = L('itens').filter(i => i.tipo === f && (inat || i.ativo) && (!q || norm(i.nome + ' ' + i.categoria + ' ' + i.local).includes(q))).sort((x, y) => x.nome.localeCompare(y.nome, 'pt-BR'));
    const tab = (id, label, n) => `<button class="tab ${f === id ? 'on' : ''}" data-act="estF" data-f="${id}">${label} <span class="count">${n}</span></button>`;
    const baixo = estoqueBaixo(), totPat = L('itens').filter(i => i.ativo && i.tipo === 'Patrimônio').reduce((s, i) => s + i.valor * (i.quantidade || 1), 0);
    return `<div class="page-head"><div><h1>Estoque e patrimônio</h1><p class="sub">Materiais de consumo com alerta de estoque baixo e o inventário dos bens da casa.</p></div>
      <div class="toolbar"><button class="btn sm" data-act="estImprimir">${ic('print')} Imprimir ${f === 'Consumo' ? 'estoque' : 'inventário'}</button><button class="btn primary" data-act="itemNovo" data-tipo="${f}">${ic('plus')} Novo item</button></div></div>
      ${baixo.length ? `<div class="note warn"><b>Estoque baixo (${baixo.length})</b><span>${baixo.map(i => `${esc(i.nome)}: ${num(i.quantidade)} ${esc(i.unidade)}`).join(' · ')}</span></div>` : ''}
      <div class="tabs" role="tablist">${tab('Consumo', 'Consumo', L('itens').filter(i => i.ativo && i.tipo === 'Consumo').length)}${tab('Patrimônio', 'Patrimônio', L('itens').filter(i => i.ativo && i.tipo === 'Patrimônio').length)}</div>
      <div class="toolbar"><label class="search"><span class="sr">Buscar</span>${ic('search')}<input data-in="estQ" value="${esc(P.q || '')}" placeholder="Nome, categoria ou local" autocomplete="off"></label><label class="check small"><input type="checkbox" data-ch="estInat" ${inat ? 'checked' : ''}> Mostrar desativados</label></div>
      ${f === 'Patrimônio' && totPat ? `<p class="small muted">Valor estimado do patrimônio: <b>${fmtBRL(totPat)}</b>.</p>` : ''}
      <div id="est-res">${estLista(l, f)}</div>`;
  }
};
function estLista(l, f) {
  if (!l.length) return '<div class="table-wrap"><div class="empty">Nenhum item.</div></div>';
  return `<div class="table-wrap cards"><table><thead><tr><th>Item</th><th>Quantidade</th>${f === 'Consumo' ? '<th>Mínimo</th>' : '<th>Valor</th>'}<th>Local</th><th><span class="sr">Ações</span></th></tr></thead><tbody>
    ${l.map(i => { const b = i.tipo === 'Consumo' && i.minimo > 0 && i.quantidade <= i.minimo; return `<tr class="${i.ativo ? '' : 'dim'}"><td class="lead"><b>${esc(i.nome)}</b>${i.categoria ? `<div class="s">${esc(i.categoria)}</div>` : ''}</td>
      <td data-label="Quantidade" class="num">${num(i.quantidade)} ${esc(i.unidade)} ${b ? pill('baixo', 'warn') : ''}</td>
      ${f === 'Consumo' ? `<td data-label="Mínimo" class="num">${i.minimo ? num(i.minimo) : '—'}</td>` : `<td data-label="Valor" class="num">${i.valor ? fmtBRL(i.valor) : '—'}</td>`}
      <td data-label="Local" class="small">${esc(i.local || '—')}</td>
      <td class="actions">${i.ativo ? `<button class="btn sm" data-act="movNovo" data-id="${i.id}" data-tipo="Entrada">${ic('plus')} Entrada</button><button class="btn sm" data-act="movNovo" data-id="${i.id}" data-tipo="Saída">${ic('minus')} Saída</button>` : ''}<button class="btn sm ghost" data-act="itemVer" data-id="${i.id}">Detalhes</button></td></tr>`; }).join('')}
  </tbody></table></div>`;
}
INP.estQ = el => { P.q = el.value; const f = P.f || 'Consumo'; const q = norm(el.value); $('#est-res').innerHTML = estLista(L('itens').filter(i => i.tipo === f && (P.inat || i.ativo) && (!q || norm(i.nome + ' ' + i.categoria + ' ' + i.local).includes(q))).sort((x, y) => x.nome.localeCompare(y.nome, 'pt-BR')), f); };
CHG.estInat = el => { P.inat = el.checked ? '1' : ''; render(); };
function itemForm(i, tipo) {
  const t = i ? i.tipo : tipo;
  return `<div class="fields"><div class="field full"><label for="it-nome">Nome</label><input id="it-nome" value="${esc(i ? i.nome : '')}" placeholder="${t === 'Patrimônio' ? 'Ex.: Atabaque, mesa, caixa de som' : 'Ex.: Vela branca 7 dias'}"></div>
    <div class="field"><label for="it-tipo">Tipo</label><select id="it-tipo">${['Consumo', 'Patrimônio'].map(x => `<option ${t === x ? 'selected' : ''}>${x}</option>`).join('')}</select></div>
    <div class="field"><label for="it-cat">Categoria</label><input id="it-cat" value="${esc(i ? i.categoria : '')}" placeholder="Ex.: velas, ervas, limpeza, instrumentos"></div>
    <div class="field"><label for="it-un">Unidade</label><input id="it-un" value="${esc(i ? i.unidade : 'un')}" placeholder="un, cx, kg, maço"></div>
    ${i ? '' : `<div class="field"><label for="it-qtd">Quantidade atual</label><input id="it-qtd" type="number" min="0" step="0.001" value="0"></div>`}
    <div class="field"><label for="it-min">Estoque mínimo</label><input id="it-min" type="number" min="0" step="0.001" value="${i ? i.minimo || '' : ''}"><span class="hint">Abaixo disso aparece o alerta.</span></div>
    <div class="field"><label for="it-local">Local</label><input id="it-local" value="${esc(i ? i.local : '')}" placeholder="Ex.: almoxarifado, congá"></div>
    <div class="field"><label for="it-val">Valor estimado (R$)</label><input id="it-val" type="number" min="0" step="0.01" value="${i && i.valor ? i.valor : ''}"></div>
    <div class="field"><label for="it-aq">Data de aquisição</label><input id="it-aq" type="date" value="${i ? i.aquisicao : ''}"></div>
    <div class="field full"><label for="it-obs">Observações</label><input id="it-obs" maxlength="500" value="${esc(i ? i.obs : '')}"></div></div>`;
}
const itemDados = () => ({ nome: $('#it-nome').value, tipo: $('#it-tipo').value, categoria: $('#it-cat').value, unidade: $('#it-un').value, quantidade: $('#it-qtd') ? $('#it-qtd').value : undefined, minimo: $('#it-min').value, local: $('#it-local').value, valor: $('#it-val').value, aquisicao: $('#it-aq').value, obs: $('#it-obs').value });
Object.assign(ACT, {
  estF(t) { P.f = t.dataset.f; P.q = ''; render(); },
  itemNovo(t) { openModal(`<h2>Novo item</h2>${itemForm(null, t.dataset.tipo)}<div class="modal-actions"><button class="btn" data-act="close">Cancelar</button><button class="btn primary" data-act="itemSalvar">Cadastrar</button></div>`, 'modal-wide'); },
  itemEditar(t) { const i = item(t.dataset.id); openModal(`<h2>Editar item</h2>${itemForm(i)}<div class="modal-actions"><button class="btn" data-act="close">Cancelar</button><button class="btn primary" data-act="itemSalvar" data-id="${i.id}">Salvar</button></div>`, 'modal-wide'); },
  itemSalvar(t) { const d = itemDados(); if (t.dataset.id) d.id = t.dataset.id; run(async () => { await api('item.salvar', d); closeModal(); P.f = d.tipo; render(); toast('Item salvo'); }); },
  itemVer(t) {
    const i = item(t.dataset.id), movs = L('movEstoque').filter(m => m.itemId === i.id).sort((x, y) => y.em.localeCompare(x.em)).slice(0, 30);
    openModal(`<h2>${esc(i.nome)}</h2><p class="muted">${esc(i.tipo)}${i.categoria ? ' · ' + esc(i.categoria) : ''} · ${num(i.quantidade)} ${esc(i.unidade)}${i.local ? ' · ' + esc(i.local) : ''}</p>
      ${i.obs ? `<p class="small">${esc(i.obs)}</p>` : ''}${i.aquisicao || i.valor ? `<p class="small muted">${i.aquisicao ? 'Adquirido em ' + fmtD(i.aquisicao) : ''}${i.valor ? ' · valor estimado ' + fmtBRL(i.valor) : ''}</p>` : ''}
      <h3 class="small">Movimentações</h3>${movs.length ? `<div class="list">${movs.map(m => `<div class="row-item"><span class="who"><b>${esc(m.tipo)}: ${num(m.quantidade)}</b><span class="small muted">${fmtD(m.data)} · saldo ${num(m.saldo)}${m.obs ? ' · ' + esc(m.obs) : ''} · ${esc(m.por)}</span></span></div>`).join('')}</div>` : '<p class="small muted">Nenhuma.</p>'}
      <div class="modal-actions"><button class="btn ghost" data-act="itemAtivo" data-id="${i.id}">${i.ativo ? 'Desativar' : 'Reativar'}</button>${i.ativo ? `<button class="btn" data-act="movNovo" data-id="${i.id}" data-tipo="Ajuste">Ajustar (inventário)</button><button class="btn" data-act="itemEditar" data-id="${i.id}">${ic('edit')} Editar</button>` : ''}<button class="btn primary" data-act="close">Fechar</button></div>`, 'modal-wide');
  },
  itemAtivo(t) { const i = item(t.dataset.id); run(async () => { await api('item.ativar', { id: i.id, ativo: !i.ativo }); closeModal(); render(); toast(i.ativo ? 'Item desativado' : 'Item reativado'); }); },
  movNovo(t) {
    const i = item(t.dataset.id), tipo = t.dataset.tipo;
    openModal(`<h2>${tipo === 'Ajuste' ? 'Ajustar estoque' : tipo === 'Entrada' ? 'Entrada no estoque' : 'Saída do estoque'}</h2><p class="muted">${esc(i.nome)} · atual: ${num(i.quantidade)} ${esc(i.unidade)}</p>
      <div class="fields"><div class="field"><label for="mv-q">${tipo === 'Ajuste' ? 'Quantidade contada' : 'Quantidade'}</label><input id="mv-q" type="number" min="0" step="0.001" value="${tipo === 'Ajuste' ? i.quantidade : ''}"></div>
      <div class="field"><label for="mv-data">Data</label><input id="mv-data" type="date" value="${TODAY()}" max="${TODAY()}"></div>
      <div class="field full"><label for="mv-obs">Observação</label><input id="mv-obs" maxlength="300" placeholder="${tipo === 'Saída' ? 'Ex.: gira de sábado' : tipo === 'Entrada' ? 'Ex.: compra, doação' : 'Ex.: contagem de inventário'}"></div></div>
      <div class="modal-actions"><button class="btn" data-act="close">Cancelar</button><button class="btn primary" data-act="movSalvar" data-id="${i.id}" data-tipo="${tipo}">Registrar</button></div>`);
  },
  movSalvar(t) { run(async () => { const r = await api('item.movimentar', { itemId: t.dataset.id, tipo: t.dataset.tipo, quantidade: $('#mv-q').value, data: $('#mv-data').value, obs: $('#mv-obs').value }); closeModal(); render(); toast(`Saldo: ${num(r.saldo)}`); }); },
  estImprimir() {
    const f = P.f || 'Consumo', l = L('itens').filter(i => i.ativo && i.tipo === f).sort((x, y) => (x.categoria + x.nome).localeCompare(y.categoria + y.nome, 'pt-BR'));
    printHTML(`<div class="doc">${docCab(f === 'Consumo' ? 'Estoque de materiais' : 'Inventário do patrimônio')}
      <table class="doc-t cols"><thead><tr><th>Item</th><th>Categoria</th><th style="width:14%">Quantidade</th>${f === 'Consumo' ? '<th style="width:12%">Mínimo</th>' : '<th style="width:16%">Valor</th>'}<th>Local</th></tr></thead><tbody>
      ${l.map(i => `<tr><td>${esc(i.nome)}</td><td>${esc(i.categoria)}</td><td>${num(i.quantidade)} ${esc(i.unidade)}</td>${f === 'Consumo' ? `<td>${i.minimo ? num(i.minimo) : ''}</td>` : `<td>${i.valor ? fmtBRL(i.valor) : ''}</td>`}<td>${esc(i.local)}</td></tr>`).join('')}</tbody></table>
      ${f === 'Patrimônio' ? `<p><b>Valor estimado total:</b> ${fmtBRL(l.reduce((s, i) => s + i.valor * (i.quantidade || 1), 0))}</p>` : ''}</div>`, 'doc-print');
  },
});

/* ============ Livro de registro de associados ============ */
function movimentosLivro() {
  const reg = L('movimentacoes').slice(), com = new Set(reg.filter(m => m.tipo === 'Ingresso').map(m => m.assocId));
  S.associados.forEach(a => { if (!com.has(a.id)) reg.push({ id: 'ing-' + a.id, assocId: a.id, tipo: 'Ingresso', de: '', para: 'Ativo', data: a.dataIngresso, derivado: true }); });
  return reg.filter(m => assoc(m.assocId)).sort((x, y) => x.data.localeCompare(y.data) || (x.em || '').localeCompare(y.em || ''));
}
VIEWS.livro = {
  mod: 'livro',
  render() {
    const as = S.associados.slice().sort((x, y) => x.matricula - y.matricula), mv = movimentosLivro();
    return `<div class="page-head"><div><h1>Livro de registro de associados</h1><p class="sub">${plural(as.length, 'associado registrado', 'associados registrados')} · ${plural(mv.length, 'movimentação', 'movimentações')}.</p></div>
      <div class="toolbar"><button class="btn primary" data-act="livroImprimir">${ic('print')} Imprimir livro</button></div></div>
      <p class="small muted">O livro traz termo de abertura, a relação de associados por matrícula, as entradas e mudanças de situação e o termo de encerramento, com espaço para as assinaturas.${isAdmin() ? '' : ' Para operadores, o CPF sai parcialmente oculto.'}</p>
      <section class="panel"><div class="panel-head"><h2>Últimas movimentações</h2></div>
        ${mv.length ? `<div class="list">${mv.slice(-15).reverse().map(m => { const a = assoc(m.assocId); return `<div class="row-item"><span class="who"><b>${esc(a.nome)}</b><span class="small muted">${fmtD(m.data)} · ${m.tipo === 'Ingresso' ? 'Ingresso' : `${esc(m.de)} → ${esc(m.para)}`}</span></span></div>`; }).join('')}</div>` : '<div class="empty">Nenhuma.</div>'}
      </section>`;
  }
};
ACT.livroImprimir = () => {
  const ac = assocCfg(), c = S.config, as = S.associados.slice().sort((x, y) => x.matricula - y.matricula), mv = movimentosLivro();
  const assin = `<div class="fass-doc"><div><span></span>${esc(c.presidenteNome || 'Presidente')}<br>Presidente</div><div><span></span>Secretário(a)</div></div>`;
  printHTML(`<div class="doc">${docCab('Livro de registro de associados')}
    <h3 style="text-align:center">TERMO DE ABERTURA</h3>
    <p style="text-align:justify">Este livro contém o registro dos associados da <b>${esc(ac.nome)}</b>${ac.cnpj ? `, inscrita no CNPJ sob o nº ${fmtCNPJ(ac.cnpj)}` : ''}${ac.endereco ? `, com sede em ${esc(ac.endereco)}` : ''}, com ${plural(as.length, 'associado registrado', 'associados registrados')} até ${fmtD(TODAY())}, emitido pelo sistema da secretaria.</p>
    <p>${esc(c.cidade || '')}, ${fmtD(TODAY())}.</p>${assin}
    <h3 style="break-before:page">RELAÇÃO DE ASSOCIADOS</h3>
    <table class="doc-t cols small"><thead><tr><th>Matr.</th><th>Nome</th><th>CPF</th><th>Nascimento</th><th>Endereço</th><th>Ingresso</th><th>Categoria</th><th>Situação</th></tr></thead><tbody>
    ${as.map(a => `<tr><td>${fmtMat(a.matricula)}</td><td>${esc(a.nome)}</td><td>${a.cpf ? (isAdmin() ? fmtCPF(a.cpf) : maskCPF(a.cpf)) : ''}</td><td>${a.nascimento ? fmtD(a.nascimento) : ''}</td><td>${esc(enderecoDe(a))}</td><td>${fmtD(a.dataIngresso)}</td><td>${esc(catNome(a.categoriaId))}</td><td>${esc(a.situacao)}</td></tr>`).join('')}</tbody></table>
    <h3 style="break-before:page">ENTRADAS E MUDANÇAS DE SITUAÇÃO</h3>
    <table class="doc-t cols small"><thead><tr><th style="width:13%">Data</th><th style="width:9%">Matr.</th><th>Nome</th><th>Movimentação</th></tr></thead><tbody>
    ${mv.map(m => { const a = assoc(m.assocId); return `<tr><td>${fmtD(m.data)}</td><td>${fmtMat(a.matricula)}</td><td>${esc(a.nome)}</td><td>${m.tipo === 'Ingresso' ? 'Ingresso na associação' : `Situação: ${esc(m.de)} → ${esc(m.para)}`}</td></tr>`; }).join('')}</tbody></table>
    <h3 style="text-align:center;break-before:page">TERMO DE ENCERRAMENTO</h3>
    <p style="text-align:justify">Encerra-se nesta data a presente emissão do livro de registro de associados da ${esc(ac.nome)}, contendo ${plural(as.length, 'associado', 'associados')} e ${plural(mv.length, 'movimentação', 'movimentações')}.</p>
    <p>${esc(c.cidade || '')}, ${fmtD(TODAY())}.</p>${assin}</div>`, 'doc-print');
};

/* ============ Documentos (versão 2.0): termos, certificados e certidões ============ */
const DOC_NOMES = { termo: 'Termo de consentimento', batismo: 'Certificado de batismo', casamento: 'Certidão de casamento religioso', ficha: 'Ficha de inscrição' };
const ATIVIDADES_TERMO = ['Viagem de praia', 'Amaci', 'Corte de quimbanda'];
const DOC_PADRAO = {
  termo: `Eu, {nome}, CPF {cpf}, nascido(a) em {nascimento}, declaro que participo, por livre e espontânea vontade, da atividade {atividade}, a realizar-se em {data}, em {local}, promovida pela {associacao} sob a orientação da direção espiritual da casa.

Declaro que:
1) conheço o caráter religioso da atividade, cujos fundamentos me foram explicados previamente;
2) seguirei as orientações dos dirigentes e da equipe responsável, inclusive quanto a horários, deslocamento e segurança;
3) sou responsável pelos meus pertences pessoais;
4) informei abaixo as minhas condições de saúde, alergias e medicamentos de uso contínuo.

Saúde/alergias: {saude}
Medicamentos: ________________________________________
Contato de emergência: {emergencia}

Em caso de necessidade, autorizo a equipe a me encaminhar ao serviço de saúde mais próximo e a avisar o meu contato de emergência.

[menor] Eu, {responsavel}, CPF {cpfResponsavel}, responsável legal, autorizo a participação do(a) menor acima e assumo as responsabilidades deste termo.`,
  batismo: `Certificamos que {nome}, nascido(a) em {nascimento}, filho(a) de {filiacao}, recebeu o Batismo na Umbanda em {data}, na {associacao}, sob a direção espiritual de {dirigente}, tendo como padrinho {padrinho} e madrinha {madrinha}.`,
  casamento: `Certificamos que, em {data}, na {associacao}, perante o(a) celebrante {celebrante} e as testemunhas {testemunha1} e {testemunha2}, uniram-se pelo rito do casamento na Umbanda {nome1}, CPF {cpf1}, e {nome2}, CPF {cpf2}.

Este ato tem caráter exclusivamente religioso e não substitui o casamento civil.`,
};
const DOC_CAMPOS = { termo: '{nome} {cpf} {nascimento} {atividade} {data} {local} {associacao} {saude} {emergencia} {responsavel} {cpfResponsavel} · linha iniciada com [menor] só aparece para menores de 18 anos',
  batismo: '{nome} {nascimento} {filiacao} {data} {associacao} {dirigente} {padrinho} {madrinha}', casamento: '{data} {associacao} {celebrante} {testemunha1} {testemunha2} {nome1} {cpf1} {nome2} {cpf2}' };
const modeloTxt = m => (S.config[{ termo: 'modeloTermo', batismo: 'modeloBatismo', casamento: 'modeloCasamento' }[m]] || '').trim() || DOC_PADRAO[m];
const dataExt = d => { if (!d) return ''; const [y, m, dd] = d.split('-').map(Number); return `${dd} de ${MESES[m - 1]} de ${y}`; };
const pessoaExt = id => L('pessoas').find(x => x.id === id);
// pessoa de um documento: "a:ID" (associado) ou "p:ID" (pessoa externa)
function pessoaDoc(ref) {
  const [t, id] = String(ref).split(':');
  if (t === 'a') { const a = assoc(id); if (!a) return null; return { ref, tipo: 'Associado', nome: a.nome, cpf: a.cpf, nascimento: a.nascimento, saude: a.saudeObs, emergencia: [a.emergNome, a.emergParentesco, a.emergTelefone].filter(Boolean).join(' · '), responsavel: a.responsavelNome, cpfResp: a.responsavelCpf, menor: ehMenor(a), filiacao: '', a }; }
  const p = pessoaExt(id); if (!p) return null;
  return { ref, tipo: 'Externa', nome: p.nome, cpf: p.cpf, nascimento: p.nascimento, saude: '', emergencia: p.telefone && p.responsavelNome ? `${p.responsavelNome} · ${p.telefone}` : '', responsavel: p.responsavelNome, cpfResp: p.responsavelCpf, menor: !!p.nascimento && idade(p.nascimento) < 18, filiacao: p.filiacao, p };
}
const BR = '____________________';
function preenche(txt, v, menor) {
  return txt.split('\n').filter(l => menor || !/^\s*\[menor\]/i.test(l)).map(l => l.replace(/^\s*\[menor\]\s*/i, '')).join('\n')
    .split(/\n\s*\n/).map(par => `<p>${esc(par).replace(/\{(\w+)\}/g, (m, k) => k in v ? (v[k] ? `<b>${esc(v[k])}</b>` : BR) : m).replace(/\n/g, '<br>')}</p>`).join('');
}
const assinaturas = l => `<div class="fass-doc">${l.map(x => `<div><span></span>${x}</div>`).join('')}</div>`;
function docHTML(em, ref) {
  const ac = assocCfg(), c = S.config, d = em.dados || {}, cid = esc(c.cidade || '');
  const reg = em.livro ? `<p class="doc-reg">Livro ${esc(em.livro)} · folha ${esc(em.folha)} · termo ${esc(em.termo)}</p>` : '';
  const base = { associacao: ac.nome, data: dataExt(em.data), local: em.local, atividade: em.atividade, dirigente: d.dirigente, padrinho: d.padrinho, madrinha: d.madrinha, celebrante: d.celebrante, testemunha1: d.testemunha1, testemunha2: d.testemunha2 };
  if (em.modelo === 'casamento') {
    const [p1, p2] = em.pessoas.map(pessoaDoc);
    const v = Object.assign(base, { nome1: p1 && p1.nome, cpf1: p1 && p1.cpf ? fmtCPF(p1.cpf) : '', nome2: p2 && p2.nome, cpf2: p2 && p2.cpf ? fmtCPF(p2.cpf) : '' });
    return `<div class="doc doc-cert">${docCab(DOC_NOMES.casamento)}${reg}<h3>CERTIDÃO DE CASAMENTO RELIGIOSO</h3>${preenche(modeloTxt('casamento'), v)}
      <p class="doc-place">${cid}, ${dataExt(TODAY())}.</p>${assinaturas(['Celebrante', 'Testemunha', 'Testemunha', 'Secretário(a)'])}</div>`;
  }
  const p = pessoaDoc(ref); if (!p) return '';
  const v = Object.assign(base, { nome: p.nome, cpf: p.cpf ? fmtCPF(p.cpf) : '', nascimento: p.nascimento ? (em.modelo === 'batismo' ? dataExt(p.nascimento) : fmtD(p.nascimento)) : '', saude: p.saude, emergencia: p.emergencia,
    responsavel: p.responsavel, cpfResponsavel: p.cpfResp ? fmtCPF(p.cpfResp) : '', filiacao: d.filiacao || p.filiacao });
  if (em.modelo === 'batismo') return `<div class="doc doc-cert">${docCab(DOC_NOMES.batismo)}${reg}<h3>CERTIFICADO DE BATISMO</h3>${preenche(modeloTxt('batismo'), v)}
      <p class="doc-place">${cid}, ${dataExt(TODAY())}.</p>${assinaturas(['Dirigente espiritual', 'Secretário(a)'])}</div>`;
  return `<div class="doc doc-termo-c">${docCab(`${DOC_NOMES.termo} · ${em.atividade}`)}<h3>TERMO DE CONSENTIMENTO E RESPONSABILIDADE</h3>${preenche(modeloTxt('termo'), v, p.menor)}
    <p class="doc-place">${cid}, ____/____/________.</p>${assinaturas(['Participante'].concat(p.menor ? ['Responsável legal'] : []))}</div>`;
}
async function imprimirEmissao(em) {
  if (em.modelo === 'ficha') { printHTML(em.pessoas.map(r => pessoaDoc(r)).filter(x => x && x.a).map(x => `<div class="ficha-uma">${fichaHTML(x.a, x.a.fotoMini || '')}</div>`).join(''), 'ficha-print', FICHA_CSS + '.ficha-uma + .ficha-uma{break-before:page}'); return; }
  const html = em.modelo === 'casamento' ? docHTML(em) : em.pessoas.map(r => docHTML(em, r)).join('');
  printHTML(html, 'doc-print');
}
const entregaDe = (em, ref) => L('entregas').find(x => x.emissaoId === em.id && x.pessoa === ref);
const emissao = id => L('emissoes').find(x => x.id === id);
function emResumo(em) {
  const n = em.pessoas.length, ent = L('entregas').filter(x => x.emissaoId === em.id).length;
  const quem = em.modelo === 'casamento' || n <= 2 ? em.pessoas.map(r => (pessoaDoc(r) || { nome: '?' }).nome).join(' e ') : plural(n, 'pessoa', 'pessoas');
  return { n, ent, quem };
}

VIEWS.documentos = {
  mod: 'documentos',
  render() {
    const f = P.f || 'emitir', tab = (id, label, n) => `<button class="tab ${f === id ? 'on' : ''}" data-act="docsF" data-f="${id}">${label}${n != null ? ` <span class="count">${n}</span>` : ''}</button>`;
    const ems = L('emissoes').filter(e => !e.cancelado);
    return `<div class="page-head"><div><h1>Documentos</h1><p class="sub">Termos de consentimento, certificados de batismo, certidões de casamento e fichas, em massa ou para pessoas específicas.</p></div></div>
      <div class="tabs" role="tablist">${tab('emitir', 'Emitir')}${tab('emitidos', 'Emitidos', ems.length)}${tab('pessoas', 'Pessoas externas', L('pessoas').length)}${tab('modelos', 'Modelos')}</div>
      ${f === 'emitidos' ? docsEmitidos() : f === 'pessoas' ? docsPessoas() : f === 'modelos' ? docsModelos() : docsEmitir()}`;
  }
};
function docsEmitir() {
  const m = UI.dcMod || 'termo', ev = P.ev || UI.dcEv || '';
  if (!UI.dcSel) UI.dcSel = new Set();
  const evs = modOn('agenda') ? L('eventos').filter(e => !e.cancelado && e.data >= TODAY()).sort((x, y) => x.data.localeCompare(y.data)) : [];
  const e = ev ? evento(ev) : null, atv = UI.dcAtv || ATIVIDADES_TERMO[0];
  const inp = (id, label, val, extra = '') => `<div class="field"><label for="${id}">${label}</label><input id="${id}" value="${esc(val || '')}" ${extra}></div>`;
  const reg = `${inp('dc-livro', 'Livro', '', 'inputmode="numeric"')}${inp('dc-folha', 'Folha', '', 'inputmode="numeric"')}${inp('dc-termo', 'Termo', '', 'inputmode="numeric"')}`;
  const campos = m === 'termo' ? `<div class="field"><label for="dc-atv">Atividade</label><select id="dc-atv" data-ch="dcAtv">${ATIVIDADES_TERMO.concat('Outra').map(x => `<option ${x === atv ? 'selected' : ''}>${x}</option>`).join('')}</select></div>
      ${atv === 'Outra' ? inp('dc-atvo', 'Qual atividade?', '', 'maxlength="80"') : ''}
      ${evs.length ? `<div class="field"><label for="dc-ev">Sessão da agenda (opcional)</label><select id="dc-ev" data-ch="dcEv"><option value="">— sem ligação —</option>${evs.map(x => `<option value="${x.id}" ${x.id === ev ? 'selected' : ''}>${fmtD(x.data)} · ${esc(x.titulo)}</option>`).join('')}</select><span class="hint">Permite controlar quem devolveu o termo assinado.</span></div>` : ''}
      ${inp('dc-data', 'Data da atividade', e ? e.data : '', 'type="date"')}${inp('dc-local', 'Local', e ? e.local : '', 'maxlength="150" placeholder="Ex.: Praia de Tramandaí"')}`
    : m === 'batismo' ? `${inp('dc-data', 'Data do batismo', '', 'type="date"')}${reg}${inp('dc-dir', 'Dirigente espiritual', UI.dcDir || '', 'maxlength="150"')}${inp('dc-pad', 'Padrinho', '', 'maxlength="150"')}${inp('dc-mad', 'Madrinha', '', 'maxlength="150"')}
      <div class="field full"><label for="dc-fil">Filiação</label><input id="dc-fil" maxlength="250" placeholder="Ex.: Maria da Silva e João da Silva (se vazio, usa o cadastro da pessoa externa)"></div>`
    : m === 'casamento' ? `${inp('dc-data', 'Data do casamento', '', 'type="date"')}${reg}${inp('dc-cel', 'Celebrante', '', 'maxlength="150"')}${inp('dc-t1', 'Testemunha 1', '', 'maxlength="150"')}${inp('dc-t2', 'Testemunha 2', '', 'maxlength="150"')}`
    : '<p class="small muted full">A ficha sai preenchida com os dados do cadastro, uma por associado.</p>';
  const lim = m === 'batismo' ? 'Escolha 1 pessoa.' : m === 'casamento' ? 'Escolha as 2 pessoas do casal.' : m === 'ficha' ? 'Somente associados.' : 'Escolha quantas pessoas quiser.';
  const pres = e ? presentes(e.id) : [];
  return `<section class="panel"><div class="fields">
      <div class="field"><label for="dc-mod">Documento</label><select id="dc-mod" data-ch="dcMod">${Object.keys(DOC_NOMES).map(k => `<option value="${k}" ${k === m ? 'selected' : ''}>${DOC_NOMES[k]}</option>`).join('')}</select></div>
      ${campos}</div></section>
    <section class="panel"><div class="panel-head"><h2>Para quem <span class="count" id="dc-n">${UI.dcSel.size}</span></h2>
      <div class="toolbar"><button class="btn sm" data-act="dcTodos">${ic('users')} Todos os ativos</button>${pres.length ? `<button class="btn sm" data-act="dcPres">${ic('check')} Presentes na sessão (${pres.length})</button>` : ''}<button class="btn sm ghost" data-act="dcLimpar">Limpar</button></div></div>
      <p class="small muted">${lim}</p>
      <div class="toolbar"><label class="search"><span class="sr">Buscar</span>${ic('search')}<input data-in="dcQ" value="${esc(UI.dcQ || '')}" placeholder="Nome ou matrícula" autocomplete="off"></label>
        <select data-ch="dcCat" aria-label="Categoria"><option value="">Todas as categorias</option>${S.categorias.map(c => `<option value="${c.id}" ${UI.dcCat === c.id ? 'selected' : ''}>${esc(c.nome)}</option>`).join('')}</select>
        <button class="btn sm ghost" data-act="dcVisiveis">Marcar os da lista</button></div>
      <div id="dc-lista" class="dc-lista">${dcLista()}</div></section>
    <div class="toolbar" style="justify-content:flex-end"><button class="btn primary" data-act="dcEmitir">${ic('print')} Emitir e imprimir</button></div>`;
}
function dcCandidatos() {
  const m = UI.dcMod || 'termo', q = norm(UI.dcQ || '');
  const as = S.associados.filter(a => a.situacao !== 'Desligado' && (!UI.dcCat || a.categoriaId === UI.dcCat)).map(a => ({ ref: 'a:' + a.id, nome: a.nome, s: `Matrícula ${fmtMat(a.matricula)} · ${esc(a.situacao)}`, b: norm(a.nome + ' ' + fmtMat(a.matricula)) }));
  const ps = m === 'ficha' || UI.dcCat ? [] : L('pessoas').map(p => ({ ref: 'p:' + p.id, nome: p.nome, s: 'Pessoa externa', b: norm(p.nome) }));
  return as.concat(ps).filter(x => !q || x.b.includes(q)).sort((x, y) => x.nome.localeCompare(y.nome, 'pt-BR'));
}
function dcLista() {
  const l = dcCandidatos(); if (!l.length) return '<div class="empty">Ninguém encontrado.</div>';
  return l.map(x => `<label class="check dc-item"><input type="checkbox" data-ch="dcMarca" value="${x.ref}" ${UI.dcSel.has(x.ref) ? 'checked' : ''}><span><b>${esc(x.nome)}</b> <span class="small muted">${x.s}</span></span></label>`).join('');
}
const dcAtualiza = () => { $('#dc-lista').innerHTML = dcLista(); $('#dc-n').textContent = UI.dcSel.size; };
INP.dcQ = el => { UI.dcQ = el.value; $('#dc-lista').innerHTML = dcLista(); };
CHG.dcCat = el => { UI.dcCat = el.value; dcAtualiza(); };
CHG.dcMarca = el => { el.checked ? UI.dcSel.add(el.value) : UI.dcSel.delete(el.value); $('#dc-n').textContent = UI.dcSel.size; };
CHG.dcMod = el => { UI.dcMod = el.value; if (el.value === 'ficha') [...UI.dcSel].forEach(r => r[0] === 'p' && UI.dcSel.delete(r)); render(); };
CHG.dcAtv = el => { UI.dcAtv = el.value; render(); };
CHG.dcEv = el => { UI.dcEv = el.value; P.ev = ''; render(); };
function docsEmitidos() {
  const l = L('emissoes').slice().sort((x, y) => y.emitidoEm.localeCompare(x.emitidoEm));
  if (!l.length) return '<div class="table-wrap"><div class="empty">Nenhum documento emitido ainda.</div></div>';
  return `<div class="table-wrap cards"><table><thead><tr><th>Documento</th><th>Para</th><th>Termos assinados</th><th><span class="sr">Ações</span></th></tr></thead><tbody>
    ${l.map(em => { const r = emResumo(em), e = em.eventoId ? evento(em.eventoId) : null; return `<tr class="${em.cancelado ? 'dim' : ''}"><td class="lead"><b>${DOC_NOMES[em.modelo]}${em.atividade && em.modelo === 'termo' ? ' · ' + esc(em.atividade) : ''}</b><div class="s">${em.data ? fmtD(em.data) : ''}${em.livro ? ` · L${esc(em.livro)} F${esc(em.folha)} T${esc(em.termo)}` : ''}${e ? ' · ' + esc(e.titulo) : ''} · emitido em ${fmtD(em.emitidoEm.slice(0, 10))}${em.cancelado ? ' · cancelado: ' + esc(em.motivoCancelamento) : ''}</div></td>
      <td data-label="Para">${esc(r.quem)}</td>
      <td data-label="Termos assinados">${em.modelo === 'termo' ? (r.ent >= r.n ? pill(`${r.ent}/${r.n} entregues`, 'ok') : pill(`${r.ent}/${r.n} · faltam ${r.n - r.ent}`, 'warn')) : '—'}</td>
      <td class="actions">${em.cancelado ? '' : `${em.modelo === 'termo' ? `<button class="btn sm" data-act="emControle" data-id="${em.id}">${ic('check')} Controle</button>` : ''}<button class="btn sm ghost" data-act="emImprimir" data-id="${em.id}">${ic('print')} Imprimir</button><button class="btn sm ghost danger-text" data-act="emCancelar" data-id="${em.id}">${ic('x')}<span class="sr">Cancelar</span></button>`}</td></tr>`; }).join('')}
  </tbody></table></div>`;
}
function emControleHTML(em) {
  const r = emResumo(em);
  return `<h2>Termos assinados</h2><p class="muted">${esc(em.atividade)}${em.data ? ' · ' + fmtD(em.data) : ''} · ${r.ent} de ${r.n} entregues</p>
    <div class="list">${em.pessoas.map(ref => { const p = pessoaDoc(ref) || { nome: '(excluída)', tipo: '' }, x = entregaDe(em, ref);
      return `<div class="row-item"><span class="who"><b>${esc(p.nome)}</b><span class="small muted">${p.tipo === 'Externa' ? 'Pessoa externa · ' : ''}${x ? `Entregue em ${fmtD(x.entregueEm.slice(0, 10))}${x.arquivoId ? ' · com arquivo' : ''}` : 'Falta entregar'}</span></span>
        <span class="toolbar">${x ? `${x.arquivoId ? `<button class="btn sm ghost" data-act="entVer" data-id="${x.id}">${ic('file')} Ver</button>` : ''}<button class="btn sm ghost" data-act="entDesfazer" data-id="${x.id}" data-em="${em.id}">Desfazer</button>` : `<button class="btn sm" data-act="entMarcar" data-em="${em.id}" data-p="${ref}">${ic('check')} Entregue</button>`}
        <label class="btn sm ghost">${ic('clip')} Anexar<input type="file" accept="image/*,application/pdf" hidden data-ch="entAnexar" data-em="${em.id}" data-p="${ref}"></label></span></div>`; }).join('')}</div>
    <div class="modal-actions"><button class="btn" data-act="close">Fechar</button></div>`;
}
const emControleAbre = id => openModal(emControleHTML(emissao(id)), 'modal-wide');
CHG.entAnexar = el => { const f = el.files[0]; if (!f) return; run(async () => { const prep = await prepComprovante(f, 2400, 0.85); await api('entrega.registrar', { emissaoId: el.dataset.em, pessoa: el.dataset.p, ...prep }, { msg: 'Enviando o termo…' }); emControleAbre(el.dataset.em); render(); toast('Termo assinado anexado'); }); };
function docsPessoas() {
  const l = L('pessoas').slice().sort((x, y) => x.nome.localeCompare(y.nome, 'pt-BR'));
  return `<div class="toolbar" style="justify-content:space-between"><p class="small muted">Pessoas que não são associadas (visitantes, batizandos, noivos). Não geram mensalidade nem carteirinha.</p><button class="btn primary" data-act="pesNovo">${ic('plus')} Nova pessoa</button></div>
    ${l.length ? `<div class="list">${l.map(p => `<div class="row-item"><span class="who"><b>${esc(p.nome)}</b><span class="small muted">${[p.cpf ? fmtCPF(p.cpf) : '', p.nascimento ? fmtD(p.nascimento) : '', p.telefone, p.responsavelNome ? 'resp.: ' + p.responsavelNome : ''].filter(Boolean).map(esc).join(' · ')}</span></span>
      <span class="toolbar"><button class="btn sm ghost" data-act="pesEditar" data-id="${p.id}">${ic('edit')} Editar</button></span></div>`).join('')}</div>` : '<div class="table-wrap"><div class="empty">Nenhuma pessoa externa cadastrada.</div></div>'}`;
}
function pesForm(p) {
  p = p || {}; const i = (id, l, v, x = '') => `<div class="field"><label for="${id}">${l}</label><input id="${id}" value="${esc(v || '')}" ${x}></div>`;
  return `<div class="fields"><div class="field full"><label for="pe-nome">Nome completo</label><input id="pe-nome" value="${esc(p.nome || '')}" maxlength="150"></div>
    ${i('pe-cpf', 'CPF', p.cpf ? fmtCPF(p.cpf) : '', 'inputmode="numeric"')}${i('pe-nasc', 'Nascimento', p.nascimento, `type="date" max="${TODAY()}"`)}${i('pe-tel', 'Telefone', p.telefone, 'data-fmt="tel" placeholder="(54) 9 0000-0000"')}${i('pe-email', 'E-mail', p.email, 'type="email"')}
    <div class="field full"><label for="pe-fil">Filiação</label><input id="pe-fil" value="${esc(p.filiacao || '')}" maxlength="250" placeholder="Nome da mãe e do pai (usado no certificado de batismo)"></div>
    ${i('pe-rn', 'Responsável legal (menores)', p.responsavelNome, 'data-fmt="nome" maxlength="150"')}${i('pe-rc', 'CPF do responsável', p.responsavelCpf ? fmtCPF(p.responsavelCpf) : '', 'inputmode="numeric"')}
    <div class="field full"><label for="pe-obs">Observações</label><input id="pe-obs" value="${esc(p.obs || '')}" maxlength="500"></div></div>`;
}
function docsModelos() {
  const adm = isAdmin();
  return `<p class="small muted">Textos usados nos documentos. Os campos entre chaves são preenchidos automaticamente; o que ficar sem dado sai como linha para preencher à mão.${adm ? '' : ' Somente o administrador altera os modelos.'}</p>
    ${['termo', 'batismo', 'casamento'].map(m => `<section class="panel"><div class="panel-head"><h2>${DOC_NOMES[m]}</h2>${S.config[{ termo: 'modeloTermo', batismo: 'modeloBatismo', casamento: 'modeloCasamento' }[m]] ? pill('personalizado', 'ok') : pill('padrão', '')}</div>
      <textarea id="md-${m}" rows="${m === 'termo' ? 16 : 6}" maxlength="6000" ${adm ? '' : 'readonly'} style="width:100%">${esc(modeloTxt(m))}</textarea>
      <p class="small muted">Campos: ${esc(DOC_CAMPOS[m])}</p>
      ${adm ? `<div class="toolbar"><button class="btn sm ghost" data-act="mdPadrao" data-m="${m}">Voltar ao padrão</button><button class="btn sm primary" data-act="mdSalvar" data-m="${m}">Salvar</button></div>` : ''}</section>`).join('')}`;
}
Object.assign(ACT, {
  docsF(t) { P.f = t.dataset.f; render(); },
  dcTodos() { S.associados.filter(a => a.situacao === 'Ativo' && (!UI.dcCat || a.categoriaId === UI.dcCat)).forEach(a => UI.dcSel.add('a:' + a.id)); dcAtualiza(); },
  dcPres() { const e = evento(P.ev || UI.dcEv); presentes(e.id).forEach(x => UI.dcSel.add('a:' + x.a.id)); dcAtualiza(); },
  dcVisiveis() { dcCandidatos().forEach(x => UI.dcSel.add(x.ref)); dcAtualiza(); },
  dcLimpar() { UI.dcSel.clear(); dcAtualiza(); },
  dcEmitir() {
    const m = UI.dcMod || 'termo', v = id => ($(id) ? $(id).value : ''), atv = m === 'termo' ? (v('#dc-atv') === 'Outra' ? v('#dc-atvo') : v('#dc-atv')) : '';
    const d = { modelo: m, atividade: atv, eventoId: m === 'termo' ? v('#dc-ev') : '', data: v('#dc-data'), local: v('#dc-local'), livro: v('#dc-livro'), folha: v('#dc-folha'), termo: v('#dc-termo'), pessoas: [...UI.dcSel],
      dados: { dirigente: v('#dc-dir'), padrinho: v('#dc-pad'), madrinha: v('#dc-mad'), filiacao: v('#dc-fil'), celebrante: v('#dc-cel'), testemunha1: v('#dc-t1'), testemunha2: v('#dc-t2') } };
    if (m === 'batismo') UI.dcDir = d.dados.dirigente;
    run(async () => { const r = await api('emissao.criar', d, { msg: 'Emitindo…' }); UI.dcSel.clear(); toast(`${DOC_NOMES[m]} emitido`); P.f = 'emitidos'; render(); imprimirEmissao(emissao(r.id)); });
  },
  emImprimir(t) { imprimirEmissao(emissao(t.dataset.id)); },
  emControle(t) { emControleAbre(t.dataset.id); },
  emCancelar(t) { openModal(`<h2>Cancelar documento</h2><p class="muted">O registro continua no histórico, marcado como cancelado.</p><div class="field"><label for="em-mot">Motivo</label><input id="em-mot" maxlength="300"></div>
    <div class="modal-actions"><button class="btn" data-act="close">Voltar</button><button class="btn danger" data-act="emCancelarOk" data-id="${t.dataset.id}">Cancelar documento</button></div>`); },
  emCancelarOk(t) { run(async () => { await api('emissao.cancelar', { id: t.dataset.id, motivo: $('#em-mot').value }); closeModal(); render(); toast('Documento cancelado'); }); },
  entMarcar(t) { run(async () => { await api('entrega.registrar', { emissaoId: t.dataset.em, pessoa: t.dataset.p }); emControleAbre(t.dataset.em); render(); }); },
  entDesfazer(t) { run(async () => { await api('entrega.desfazer', { id: t.dataset.id }); emControleAbre(t.dataset.em); render(); }); },
  entVer(t) { run(async () => { const r = await api('entrega.arquivo', { id: t.dataset.id }, { msg: 'Abrindo…' }); mostrarArquivo(r, 'Termo assinado'); }); },
  pesNovo() { openModal(`<h2>Nova pessoa externa</h2>${pesForm()}<div class="modal-actions"><button class="btn" data-act="close">Cancelar</button><button class="btn primary" data-act="pesSalvar">Cadastrar</button></div>`, 'modal-wide'); },
  pesEditar(t) { const p = pessoaExt(t.dataset.id); openModal(`<h2>Editar pessoa externa</h2>${pesForm(p)}<div class="modal-actions"><button class="btn ghost danger-text" data-act="pesExcluir" data-id="${p.id}">Excluir</button><button class="btn" data-act="close">Cancelar</button><button class="btn primary" data-act="pesSalvar" data-id="${p.id}">Salvar</button></div>`, 'modal-wide'); },
  pesSalvar(t) { const d = { id: t.dataset.id || '', nome: $('#pe-nome').value, cpf: $('#pe-cpf').value, nascimento: $('#pe-nasc').value, telefone: $('#pe-tel').value, email: $('#pe-email').value, filiacao: $('#pe-fil').value, responsavelNome: $('#pe-rn').value, responsavelCpf: $('#pe-rc').value, obs: $('#pe-obs').value };
    run(async () => { await api('pessoa.salvar', d); closeModal(); render(); toast('Pessoa salva'); }); },
  pesExcluir(t) { run(async () => { await api('pessoa.excluir', { id: t.dataset.id }); closeModal(); render(); toast('Pessoa excluída'); }); },
  mdSalvar(t) { const m = t.dataset.m, txt = $('#md-' + m).value; run(async () => { await api('modelo.salvar', { modelo: m, texto: txt.trim() === DOC_PADRAO[m] ? '' : txt }); render(); toast('Modelo salvo'); }); },
  mdPadrao(t) { const m = t.dataset.m; run(async () => { await api('modelo.salvar', { modelo: m, texto: '' }); render(); toast('Modelo voltou ao padrão'); }); },
  evTermo(t) { UI.dcMod = 'termo'; UI.dcEv = t.dataset.id; P.ev = ''; go('documentos', { f: 'emitir' }); },
});
