// =========================================================================
// Painel FTC — app.js
// Vanilla JS + Firebase (modular SDK via CDN). Sem build step.
// =========================================================================

import { firebaseConfig } from "./firebase-config.js";

import { initializeApp } from "https://www.gstatic.com/firebasejs/10.13.0/firebase-app.js";
import {
  getAuth,
  signInAnonymously,
  onAuthStateChanged
} from "https://www.gstatic.com/firebasejs/10.13.0/firebase-auth.js";
import {
  getFirestore,
  collection,
  addDoc,
  setDoc,
  updateDoc,
  deleteDoc,
  doc,
  arrayUnion,
  arrayRemove,
  getDoc,
  Bytes,
  onSnapshot,
  query,
  orderBy,
  serverTimestamp
} from "https://www.gstatic.com/firebasejs/10.13.0/firebase-firestore.js";

// ---------------------------------------------------------------------
// Firebase init
// ---------------------------------------------------------------------

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

let currentUid = null;

const connEl = document.getElementById("connStatus");

onAuthStateChanged(auth, (user) => {
  if (user) {
    currentUid = user.uid; // identificador técnico interno; nunca exibido na UI
    setConn("online");
    startListeners();
  }
});

signInAnonymously(auth).catch((err) => {
  console.error("Falha na autenticação anônima:", err);
  setConn("offline");
});

function setConn(state) {
  connEl.classList.remove("conn--connecting", "conn--online", "conn--offline");
  connEl.classList.add(`conn--${state}`);
  connEl.title =
    state === "online" ? "Conectado" :
    state === "offline" ? "Sem conexão" : "Conectando...";
}

// ---------------------------------------------------------------------
// Constantes de domínio
// ---------------------------------------------------------------------

const CELULAS = [
  { id: "d1", nome: "D-1 Pessoal", abrev: "D1" },
  { id: "d2", nome: "D-2 Inteligência", abrev: "D2" },
  { id: "d3", nome: "D-3 Operações", abrev: "D3" },
  { id: "d4", nome: "D-4 Logística", abrev: "D4" },
  { id: "d5", nome: "D-5 Planejamento", abrev: "D5" },
  { id: "d6", nome: "D-6 C2/Ciber", abrev: "D6" },
  { id: "d7d8", nome: "D-7/D-8 ComSoc/OpInfo", abrev: "D7/D8" },
  { id: "d9", nome: "D-9 Assuntos Civis", abrev: "D9" },
  { id: "d10", nome: "D-10 Finanças", abrev: "D10" }
];

const CELULA_MAP = Object.fromEntries(CELULAS.map((c) => [c.id, c]));

const URGENCIA_LABEL = {
  critico: "Crítico",
  importante: "Importante",
  informativo: "Informativo"
};

const STATUS_LABEL = {
  pendente: "Pendente",
  pronta: "Pronto"
};

// ---------------------------------------------------------------------
// Status do PMS — derivado das respostas, nunca editado à mão:
// fica "pronta" só quando TODAS as células roteadas anexaram resposta.
// ---------------------------------------------------------------------

function celulasRoteadas(it) {
  return CELULAS.filter((c) => it.celulas && it.celulas[c.id] && it.celulas[c.id].sel);
}

function respostasDaCelula(it, celulaId) {
  return (it.respostas || []).filter((r) => r.celula === celulaId);
}

function statusPms(it) {
  const roteadas = celulasRoteadas(it);
  if (roteadas.length === 0) return "pendente";
  return roteadas.every((c) => respostasDaCelula(it, c.id).length > 0) ? "pronta" : "pendente";
}

// ---------------------------------------------------------------------
// Estado local (dados em cache vindos do Firestore)
// ---------------------------------------------------------------------

let injecoesCache = [];
let simulasCache = [];

// ---------------------------------------------------------------------
// Navegação por abas
// ---------------------------------------------------------------------

const tabs = document.querySelectorAll(".tab");
const views = document.querySelectorAll(".view");

function ativarAba(target) {
  tabs.forEach((t) => t.classList.toggle("tab--active", t.dataset.tab === target));
  views.forEach((v) => v.classList.toggle("view--active", v.id === `view-${target}`));
}

tabs.forEach((tab) => {
  tab.addEventListener("click", () => ativarAba(tab.dataset.tab));
});

// ---------------------------------------------------------------------
// Seletor global "Ver como"
// ---------------------------------------------------------------------

const VIEWAS_KEY = "painelEm3_verComo";
const viewAsSelect = document.getElementById("viewAsSelect");

function getViewAs() {
  return localStorage.getItem(VIEWAS_KEY) || "todas";
}

function setViewAs(value) {
  localStorage.setItem(VIEWAS_KEY, value);
}

viewAsSelect.value = getViewAs();
viewAsSelect.addEventListener("change", () => {
  setViewAs(viewAsSelect.value);
  renderDirex();
  renderSimulas();
});

function isCommandView(viewAs) {
  return viewAs === "todas" || viewAs === "comto";
}

// ---------------------------------------------------------------------
// Validação de formulários — mensagem de erro visível (nunca só foco
// silencioso) quando campos obrigatórios estão vazios.
// ---------------------------------------------------------------------

function buildRequiredFieldsMessage(labels) {
  if (labels.length === 1) return `Preencha ${labels[0]} antes de registrar.`;
  const allButLast = labels.slice(0, -1).join(", ");
  const last = labels[labels.length - 1];
  return `Preencha ${allButLast} e ${last} antes de registrar.`;
}

// fields: [{ el, label }] em ordem de exibição no formulário.
// Retorna null se tudo válido, ou { message, firstEmptyEl } caso falte algo.
function validateRequiredFields(fields) {
  const missing = fields.filter((f) => !f.el.value.trim());
  if (missing.length === 0) return null;
  return {
    message: buildRequiredFieldsMessage(missing.map((f) => f.label)),
    firstEmptyEl: missing[0].el
  };
}

// ---------------------------------------------------------------------
// Anexos (DOCX/PDF) — guardados no próprio Firestore (plano gratuito),
// fatiados em partes de até 900 KB, pois cada documento tem limite de 1 MiB:
//   arquivos/<id>               { nome, tipo, tamanho, partes }
//   arquivos/<id>/partes/<n>    { dados: Bytes }
// O PMS guarda só a referência: { nome, arquivoId, tamanho }.
// ---------------------------------------------------------------------

const ANEXO_MAX_BYTES = 10 * 1024 * 1024; // 10 MB
const ANEXO_PARTE_BYTES = 900 * 1024;
const ANEXO_TIPOS = {
  pdf: "application/pdf",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
};
const ANEXO_ACCEPT = ".pdf,.docx," + Object.values(ANEXO_TIPOS).join(",");

// Retorna mensagem de erro, ou null se o arquivo é aceitável.
function validateAnexo(file) {
  const ext = file.name.split(".").pop().toLowerCase();
  if (!ANEXO_TIPOS[ext]) return "Apenas arquivos DOCX ou PDF são aceitos.";
  if (file.size > ANEXO_MAX_BYTES) return "Arquivo maior que 10 MB.";
  return null;
}

async function uploadAnexo(file) {
  const ext = file.name.split(".").pop().toLowerCase();
  const bytes = new Uint8Array(await file.arrayBuffer());
  const partes = Math.max(1, Math.ceil(bytes.length / ANEXO_PARTE_BYTES));
  const arquivoRef = doc(collection(db, "arquivos"));

  for (let i = 0; i < partes; i++) {
    const fatia = bytes.subarray(i * ANEXO_PARTE_BYTES, (i + 1) * ANEXO_PARTE_BYTES);
    await setDoc(doc(arquivoRef, "partes", String(i)), { dados: Bytes.fromUint8Array(fatia) });
  }
  // Metadados por último: só existem quando todas as partes já foram gravadas.
  await setDoc(arquivoRef, {
    nome: file.name,
    tipo: ANEXO_TIPOS[ext],
    tamanho: file.size,
    partes,
    createdAt: serverTimestamp()
  });

  return { nome: file.name, arquivoId: arquivoRef.id, tamanho: file.size };
}

async function downloadAnexo(arquivoId) {
  const arquivoRef = doc(db, "arquivos", arquivoId);
  const meta = await getDoc(arquivoRef);
  if (!meta.exists()) throw new Error("Arquivo não encontrado.");
  const { nome, tipo, partes } = meta.data();

  const snaps = await Promise.all(
    Array.from({ length: partes }, (_, i) => getDoc(doc(arquivoRef, "partes", String(i))))
  );
  const blob = new Blob(snaps.map((p) => p.data().dados.toUint8Array()), { type: tipo });

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = nome;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}

async function deleteAnexo(anexo) {
  const arquivoRef = doc(db, "arquivos", anexo.arquivoId);
  const meta = await getDoc(arquivoRef);
  const partes = meta.exists() ? meta.data().partes : 0;
  await deleteDoc(arquivoRef);
  await Promise.all(
    Array.from({ length: partes }, (_, i) => deleteDoc(doc(arquivoRef, "partes", String(i))))
  );
}

function formatBytes(n) {
  if (!n) return "";
  if (n < 1024 * 1024) return `${Math.max(1, Math.round(n / 1024))} KB`;
  return `${(n / (1024 * 1024)).toFixed(1).replace(".", ",")} MB`;
}

function anexoLinkHtml(anexo) {
  return `<span><button type="button" class="anexo-link" data-arquivo-id="${escapeHtml(anexo.arquivoId)}">` +
    `${escapeHtml(anexo.nome)}</button> <span class="anexo-link__tam">${formatBytes(anexo.tamanho)}</span></span>`;
}

// Delegação: qualquer link de anexo na página dispara o download.
document.addEventListener("click", async (e) => {
  const link = e.target.closest(".anexo-link[data-arquivo-id]");
  if (!link || link.disabled) return;
  const textoOriginal = link.textContent;
  link.disabled = true;
  link.textContent = "Baixando...";
  try {
    await downloadAnexo(link.dataset.arquivoId);
  } catch (err) {
    console.error("Erro ao baixar anexo:", err);
    window.alert("Erro ao baixar o arquivo. Tente novamente.");
  } finally {
    link.disabled = false;
    link.textContent = textoOriginal;
  }
});

// ---------------------------------------------------------------------
// Formulário: Novo PMS (coleção "injecoes")
// ---------------------------------------------------------------------

const celulasGridNova = document.getElementById("celulasGridNova");
const tplCelulaBlock = document.getElementById("tpl-celula-block");
const celulaBlockRefs = {}; // id -> { checkbox, textarea }

CELULAS.forEach((celula) => {
  const node = tplCelulaBlock.content.cloneNode(true);
  const checkbox = node.querySelector(".celula-check");
  const abrevEl = node.querySelector(".celula-block__abrev");
  const nomeEl = node.querySelector(".celula-block__nome");
  const textarea = node.querySelector(".celula-block__texto");

  abrevEl.textContent = celula.abrev;
  nomeEl.textContent = celula.nome;
  checkbox.addEventListener("change", () => {
    textarea.hidden = !checkbox.checked;
    if (!checkbox.checked) textarea.value = "";
  });

  celulaBlockRefs[celula.id] = { checkbox, textarea };
  celulasGridNova.appendChild(node);
});

const formNovaInjecao = document.getElementById("formNovaInjecao");
const novaInjecaoMsg = document.getElementById("novaInjecaoMsg");
const inAnexo = document.getElementById("inAnexo");
const inAnexoNome = document.getElementById("inAnexoNome");
const inAnexoLimpar = document.getElementById("inAnexoLimpar");
const inAnexoBtn = document.getElementById("inAnexoBtn");
const formNovaTitulo = document.getElementById("formNovaTitulo");
const btnCancelarEdicao = document.getElementById("btnCancelarEdicao");
const btnSubmitPms = formNovaInjecao.querySelector('button[type="submit"]');

// Edição: o mesmo formulário do Novo PMS é reaproveitado para alterar um
// PMS já registrado. `edicao` guarda o id e o anexo atual desse PMS.
let edicao = null; // { id, gdh, anexo }
let removerAnexoAtual = false;

function anexoAtualEdicao() {
  return edicao && edicao.anexo && !removerAnexoAtual ? edicao.anexo : null;
}

function resetAnexoPicker() {
  inAnexo.value = "";
  const atual = anexoAtualEdicao();
  inAnexoNome.textContent = atual
    ? `Atual: ${atual.nome} (${formatBytes(atual.tamanho)})`
    : "Nenhum arquivo selecionado";
  inAnexoLimpar.hidden = !atual;
  inAnexoBtn.textContent = atual ? "Substituir arquivo (DOCX/PDF, até 10 MB)" : "Anexar arquivo (DOCX/PDF, até 10 MB)";
}

function limparFormPms() {
  formNovaInjecao.reset();
  CELULAS.forEach((c) => {
    const ref = celulaBlockRefs[c.id];
    ref.textarea.hidden = true;
    ref.textarea.value = "";
  });
  formNovaInjecao.querySelector('input[name="urgencia"][value="critico"]').checked = true;
  edicao = null;
  removerAnexoAtual = false;
  resetAnexoPicker();
  formNovaTitulo.textContent = "Registrar PMS";
  btnSubmitPms.textContent = "Registrar e Rotear";
  btnCancelarEdicao.hidden = true;
}

function iniciarEdicao(it) {
  limparFormPms();
  edicao = { id: it.id, gdh: it.gdh || "", anexo: it.anexo || null };

  document.getElementById("inGdh").value = it.gdh || "";
  document.getElementById("inOrigem").value = it.origem || "";
  document.getElementById("inAutor").value = it.autor || "";
  document.getElementById("inResumo").value = it.resumo || "";
  const urg = formNovaInjecao.querySelector(`input[name="urgencia"][value="${it.urgencia}"]`);
  if (urg) urg.checked = true;
  CELULAS.forEach((c) => {
    const ref = celulaBlockRefs[c.id];
    const cel = it.celulas && it.celulas[c.id];
    ref.checkbox.checked = !!(cel && cel.sel);
    ref.textarea.hidden = !ref.checkbox.checked;
    ref.textarea.value = ref.checkbox.checked ? cel.texto || "" : "";
  });

  resetAnexoPicker();
  formNovaTitulo.textContent = `Editar PMS ${it.gdh || ""}`.trim();
  btnSubmitPms.textContent = "Salvar alterações";
  btnCancelarEdicao.hidden = false;
  novaInjecaoMsg.textContent = "";
  novaInjecaoMsg.classList.remove("error");
  ativarAba("nova");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

btnCancelarEdicao.addEventListener("click", () => {
  limparFormPms();
  novaInjecaoMsg.textContent = "";
  novaInjecaoMsg.classList.remove("error");
});

inAnexo.addEventListener("change", () => {
  novaInjecaoMsg.textContent = "";
  novaInjecaoMsg.classList.remove("error");
  const file = inAnexo.files[0];
  if (!file) return resetAnexoPicker();
  const erro = validateAnexo(file);
  if (erro) {
    resetAnexoPicker();
    novaInjecaoMsg.textContent = erro;
    novaInjecaoMsg.classList.add("error");
    return;
  }
  inAnexoNome.textContent = `${file.name} (${formatBytes(file.size)})`;
  inAnexoLimpar.hidden = false;
});

// "remover": descarta o arquivo recém-escolhido; sem ele, na edição,
// marca o arquivo atual do PMS para ser apagado ao salvar.
inAnexoLimpar.addEventListener("click", () => {
  if (!inAnexo.files[0] && edicao && edicao.anexo) removerAnexoAtual = true;
  resetAnexoPicker();
});

formNovaInjecao.addEventListener("submit", async (e) => {
  e.preventDefault();
  novaInjecaoMsg.textContent = "";
  novaInjecaoMsg.classList.remove("error");

  const gdhEl = document.getElementById("inGdh");
  const origemEl = document.getElementById("inOrigem");
  const resumoEl = document.getElementById("inResumo");

  const requiredFieldsValidation = validateRequiredFields([
    { el: gdhEl, label: "o GDH" },
    { el: origemEl, label: "a Origem" },
    { el: resumoEl, label: "o resumo da situação" }
  ]);

  if (requiredFieldsValidation) {
    novaInjecaoMsg.textContent = requiredFieldsValidation.message;
    novaInjecaoMsg.classList.add("error");
    requiredFieldsValidation.firstEmptyEl.focus();
    return;
  }

  const gdh = gdhEl.value.trim();
  const origem = origemEl.value.trim();
  const autor = document.getElementById("inAutor").value.trim();
  const urgencia = formNovaInjecao.querySelector('input[name="urgencia"]:checked').value;
  const resumo = resumoEl.value.trim();

  const celulas = {};
  let algumaSelecionada = false;
  CELULAS.forEach((c) => {
    const ref = celulaBlockRefs[c.id];
    const sel = ref.checkbox.checked;
    if (sel) algumaSelecionada = true;
    celulas[c.id] = { sel, texto: sel ? ref.textarea.value.trim() : "" };
  });

  if (!algumaSelecionada) {
    novaInjecaoMsg.textContent = "Selecione ao menos uma célula para rotear.";
    novaInjecaoMsg.classList.add("error");
    return;
  }

  if (edicao) return salvarEdicao({ gdh, origem, autor, urgencia, resumo, celulas });

  const submitBtn = btnSubmitPms;
  submitBtn.disabled = true;

  try {
    let anexo = null;
    if (inAnexo.files[0]) {
      novaInjecaoMsg.textContent = "Enviando arquivo...";
      anexo = await uploadAnexo(inAnexo.files[0]);
    }

    await addDoc(collection(db, "injecoes"), {
      gdh,
      origem,
      autor,
      urgencia,
      resumo,
      celulas,
      anexo,
      respostas: [],
      createdAt: serverTimestamp()
    });

    novaInjecaoMsg.textContent = "PMS registrado e roteado.";
    limparFormPms();
  } catch (err) {
    console.error("Erro ao registrar PMS:", err);
    novaInjecaoMsg.textContent = "Erro ao registrar. Tente novamente.";
    novaInjecaoMsg.classList.add("error");
  } finally {
    submitBtn.disabled = false;
  }
});

async function salvarEdicao(dados) {
  const atual = injecoesCache.find((x) => x.id === edicao.id);
  if (!atual) {
    novaInjecaoMsg.textContent = "Este PMS não existe mais (pode ter sido excluído).";
    novaInjecaoMsg.classList.add("error");
    return;
  }

  // Células desmarcadas que já responderam: as respostas delas são apagadas.
  const orfas = (atual.respostas || []).filter((r) => !(dados.celulas[r.celula] && dados.celulas[r.celula].sel));
  if (orfas.length > 0) {
    const abrevs = [...new Set(orfas.map((r) => (CELULA_MAP[r.celula] ? CELULA_MAP[r.celula].abrev : r.celula)))];
    const ok = window.confirm(
      `${abrevs.join(", ")} ${abrevs.length > 1 ? "foram desmarcadas mas já responderam" : "foi desmarcada mas já respondeu"}. ` +
      "Salvar vai apagar essas respostas. Continuar?"
    );
    if (!ok) return;
  }

  btnSubmitPms.disabled = true;
  try {
    const anexoAntigo = edicao.anexo;
    let anexo = anexoAtualEdicao();
    if (inAnexo.files[0]) {
      novaInjecaoMsg.textContent = "Enviando arquivo...";
      anexo = await uploadAnexo(inAnexo.files[0]);
    }

    const alteracoes = { ...dados, anexo, editadoEm: serverTimestamp() };
    if (orfas.length > 0) alteracoes.respostas = arrayRemove(...orfas);
    await updateDoc(doc(db, "injecoes", edicao.id), alteracoes);

    // Melhor esforço: apaga o arquivo substituído/removido e os das respostas descartadas.
    const lixo = [...orfas];
    if (anexoAntigo && (!anexo || anexo.arquivoId !== anexoAntigo.arquivoId)) lixo.push(anexoAntigo);
    await Promise.allSettled(lixo.map(deleteAnexo));

    limparFormPms();
    novaInjecaoMsg.textContent = "PMS atualizado.";
  } catch (err) {
    console.error("Erro ao salvar PMS:", err);
    novaInjecaoMsg.textContent = "Erro ao salvar. Tente novamente.";
    novaInjecaoMsg.classList.add("error");
  } finally {
    btnSubmitPms.disabled = false;
  }
}

// ---------------------------------------------------------------------
// Controle PMS — listagem + filtros
// ---------------------------------------------------------------------

const direxList = document.getElementById("direxList");
const statusFilter = document.getElementById("statusFilter");

statusFilter.addEventListener("change", renderDirex);

function formatTimestamp(ts) {
  if (!ts || typeof ts.toDate !== "function") return "—";
  const d = ts.toDate();
  return d.toLocaleString("pt-BR", {
    day: "2-digit", month: "2-digit", year: "2-digit",
    hour: "2-digit", minute: "2-digit"
  });
}

// Seções abertas/recolhidas (por navegador) e rascunhos das caixas de
// texto de resposta: sobrevivem às re-renderizações em tempo real.
const SECOES_KEY = "painelFtc_secoesAbertas";
const secoesAbertas = new Set(lerSecoesAbertas());
const rascunhosResposta = new Map(); // "<pmsId>:<celulaId>" -> texto

function lerSecoesAbertas() {
  try {
    return JSON.parse(localStorage.getItem(SECOES_KEY)) || [];
  } catch {
    return [];
  }
}

function salvarSecoesAbertas() {
  try {
    localStorage.setItem(SECOES_KEY, JSON.stringify([...secoesAbertas]));
  } catch {
    // sem armazenamento local: as seções só não ficam lembradas
  }
}

function celulaRespondeu(it, celulaId) {
  return respostasDaCelula(it, celulaId).length > 0;
}

function renderDirex() {
  const viewAs = getViewAs();
  const statusVal = statusFilter.value;
  const secoes = isCommandView(viewAs) ? CELULAS : [CELULA_MAP[viewAs]];

  // Preserva o foco e o cursor de quem está escrevendo uma resposta.
  const ativo = document.activeElement;
  const focoChave = ativo && ativo.classList.contains("resposta-texto") ? ativo.dataset.chave : null;
  const cursor = focoChave ? [ativo.selectionStart, ativo.selectionEnd] : null;

  direxList.innerHTML = "";
  secoes.forEach((c) => direxList.appendChild(buildSecao(c, statusVal, secoes.length === 1)));

  if (focoChave) {
    const el = direxList.querySelector(`.resposta-texto[data-chave="${focoChave}"]`);
    if (el) {
      el.focus();
      el.setSelectionRange(cursor[0], cursor[1]);
    }
  }
}

function buildSecao(celula, statusVal, unica) {
  const doCelula = injecoesCache.filter((it) => it.celulas && it.celulas[celula.id] && it.celulas[celula.id].sel);
  const pendentes = doCelula.filter((it) => !celulaRespondeu(it, celula.id)).length;

  let items = doCelula;
  if (statusVal !== "todos") {
    items = items.filter((it) => (celulaRespondeu(it, celula.id) ? "pronta" : "pendente") === statusVal);
  }

  const secao = document.createElement("details");
  secao.className = "secao";
  secao.open = unica || secoesAbertas.has(celula.id);
  secao.innerHTML = `
    <summary class="secao__head">
      <span class="secao__abrev mono">${celula.abrev}</span>
      <span class="secao__nome">${escapeHtml(celula.nome)}</span>
      <span class="secao__contagem${pendentes > 0 ? " secao__contagem--pendente" : ""}">
        ${pendentes} pendente${pendentes === 1 ? "" : "s"} · ${doCelula.length} PMS
      </span>
    </summary>
    <div class="secao__body cards-list"></div>
  `;

  if (!unica) {
    secao.addEventListener("toggle", () => {
      if (secao.open) secoesAbertas.add(celula.id);
      else secoesAbertas.delete(celula.id);
      salvarSecoesAbertas();
    });
  }

  const body = secao.querySelector(".secao__body");
  if (items.length === 0) {
    body.innerHTML = `<p class="empty-state">${doCelula.length === 0
      ? "Nenhum PMS roteado para esta seção."
      : "Nenhum PMS desta seção para este filtro."}</p>`;
  } else {
    items.forEach((it) => body.appendChild(buildInjecaoCard(it, celula.id)));
  }
  return secao;
}

// Cartão de um PMS dentro da seção de uma célula: mostra o prazo, as
// respostas e o envio de resposta dessa célula.
function buildInjecaoCard(it, celulaId) {
  const card = document.createElement("article");
  card.className = "injecao-card";
  card.dataset.urgencia = it.urgencia;

  const roteadas = celulasRoteadas(it);
  const chipsHtml = roteadas
    .map((c) => `<span class="chip${c.id === celulaId ? " chip--atual" : ""}">${c.abrev}</span>`)
    .join("");

  const cel = it.celulas[celulaId];
  const prazoHtml = cel.texto
    ? `<div class="celula-texto-block"><span class="celula-texto-block__label">${CELULA_MAP[celulaId].abrev}</span>Prazo: ${escapeHtml(cel.texto)}</div>`
    : "";

  const anexoHtml = it.anexo
    ? `<div class="anexo-block"><span class="anexo-block__label">Arquivo do PMS</span>${anexoLinkHtml(it.anexo)}</div>`
    : "";

  const respostasItensHtml = respostasDaCelula(it, celulaId).map((r) => `
    <div class="resposta-item">
      <div class="resposta-item__linha">
        ${anexoLinkHtml(r)}
        <button type="button" class="impacto-row__remove resposta-remover" data-arquivo-id="${escapeHtml(r.arquivoId)}">remover</button>
      </div>
      ${r.texto ? `<div class="resposta-item__texto">${escapeHtml(r.texto)}</div>` : ""}
    </div>
  `).join("");

  const chave = `${it.id}:${celulaId}`;
  const respostaHtml = `
    <div class="resposta-block">
      <span class="anexo-block__label">Resposta da ${CELULA_MAP[celulaId].abrev}</span>
      ${respostasItensHtml || '<span class="resposta-vazia">Nenhuma resposta anexada.</span>'}
      <div class="resposta-upload">
        <textarea class="resposta-texto" rows="3" data-chave="${chave}"
          placeholder="Para quem:&#10;Assunto:&#10;Outras informações:"></textarea>
        <div class="resposta-upload__acoes">
          <label class="btn btn--ghost btn--small anexo-picker__btn">
            Anexar resposta (DOCX/PDF)
            <input type="file" class="anexo-picker__input resposta-file" accept="${ANEXO_ACCEPT}">
          </label>
          <span class="form-msg resposta-msg"></span>
        </div>
      </div>
    </div>
  `;

  const respondeu = celulaRespondeu(it, celulaId);
  const status = respondeu ? "pronta" : "pendente";
  const qtdRespondidas = roteadas.filter((c) => celulaRespondeu(it, c.id)).length;

  card.innerHTML = `
    <div class="injecao-card__head">
      <span class="pill pill--${it.urgencia}">${URGENCIA_LABEL[it.urgencia] || it.urgencia}</span>
      <span class="pill pill--status-${status}">${STATUS_LABEL[status]}</span>
      <span class="injecao-card__gdh mono">${escapeHtml(it.gdh || "—")}</span>
      <span class="injecao-card__origem">${escapeHtml(it.origem || "")}</span>
    </div>
    <div class="injecao-card__resumo">${escapeHtml(it.resumo || "")}</div>
    <div class="chips">${chipsHtml}<span class="chips__progresso">${qtdRespondidas}/${roteadas.length} células responderam</span></div>
    ${prazoHtml}
    ${anexoHtml}
    ${respostaHtml}
    <div class="injecao-card__footer">
      <span class="injecao-card__autor">${it.autor ? "Autor: " + escapeHtml(it.autor) + " · " : ""}${formatTimestamp(it.createdAt)}${it.editadoEm ? " · editado " + formatTimestamp(it.editadoEm) : ""}</span>
      <button type="button" class="btn btn--ghost btn--small btn-editar">Editar</button>
      <button type="button" class="btn btn--danger-ghost btn--small btn-excluir">Excluir</button>
    </div>
  `;

  card.querySelectorAll(".resposta-remover").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const resposta = (it.respostas || []).find((r) => r.arquivoId === btn.dataset.arquivoId);
      if (!resposta) return;
      const abrev = CELULA_MAP[resposta.celula] ? CELULA_MAP[resposta.celula].abrev : resposta.celula;
      const ok = window.confirm(`Remover a resposta "${resposta.nome}" da ${abrev}? O PMS volta a ficar pendente para esta célula se não houver outra resposta dela.`);
      if (!ok) return;
      btn.disabled = true;
      try {
        // arrayRemove exige o objeto idêntico ao gravado.
        await updateDoc(doc(db, "injecoes", it.id), { respostas: arrayRemove(resposta) });
        await deleteAnexo(resposta).catch((err) => console.error("Erro ao apagar arquivo da resposta:", err));
      } catch (err) {
        console.error("Erro ao remover resposta:", err);
        btn.disabled = false;
        window.alert("Erro ao remover. Tente novamente.");
      }
    });
  });

  const respostaTexto = card.querySelector(".resposta-texto");
  respostaTexto.value = rascunhosResposta.get(chave) || "";
  respostaTexto.addEventListener("input", () => {
    rascunhosResposta.set(chave, respostaTexto.value);
  });

  const respostaFile = card.querySelector(".resposta-file");
  const respostaMsg = card.querySelector(".resposta-msg");
  respostaFile.addEventListener("change", async () => {
    const file = respostaFile.files[0];
    if (!file) return;
    respostaMsg.classList.remove("error");

    const texto = respostaTexto.value.trim();
    const erro = !texto
      ? "Preencha para quem, o assunto e as demais informações antes de anexar."
      : validateAnexo(file);
    if (erro) {
      respostaMsg.textContent = erro;
      respostaMsg.classList.add("error");
      respostaFile.value = "";
      if (!texto) respostaTexto.focus();
      return;
    }

    respostaFile.disabled = true;
    respostaTexto.disabled = true;
    respostaMsg.textContent = "Enviando arquivo...";
    try {
      const anexo = await uploadAnexo(file);
      // Limpa o rascunho antes de gravar: o listener re-renderiza o cartão
      // já com a nova resposta e a caixa vazia.
      rascunhosResposta.delete(chave);
      await updateDoc(doc(db, "injecoes", it.id), {
        respostas: arrayUnion({ ...anexo, celula: celulaId, texto, enviadoEm: Date.now() })
      });
    } catch (err) {
      console.error("Erro ao anexar resposta:", err);
      rascunhosResposta.set(chave, texto);
      respostaMsg.textContent = "Erro ao enviar. Tente novamente.";
      respostaMsg.classList.add("error");
      respostaFile.disabled = false;
      respostaTexto.disabled = false;
      respostaFile.value = "";
    }
  });

  card.querySelector(".btn-editar").addEventListener("click", () => iniciarEdicao(it));

  const btnExcluir = card.querySelector(".btn-excluir");
  btnExcluir.addEventListener("click", async () => {
    const ok = window.confirm(
      `Excluir o PMS ${it.gdh || ""}? Os arquivos anexados também serão apagados. Esta ação não pode ser desfeita.`
    );
    if (!ok) return;
    btnExcluir.disabled = true;
    try {
      await deleteDoc(doc(db, "injecoes", it.id));
      const arquivos = [it.anexo, ...(it.respostas || [])].filter((a) => a && a.arquivoId);
      // Melhor esforço: falha ao apagar um arquivo não impede a exclusão do PMS.
      await Promise.allSettled(arquivos.map(deleteAnexo));
    } catch (err) {
      console.error("Erro ao excluir PMS:", err);
      btnExcluir.disabled = false;
      window.alert("Erro ao excluir. Tente novamente.");
    }
  });

  return card;
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str == null ? "" : String(str);
  return div.innerHTML;
}

// ---------------------------------------------------------------------
// Súmulas de Reunião
// ---------------------------------------------------------------------

const impactosList = document.getElementById("impactosList");
const btnAddImpacto = document.getElementById("btnAddImpacto");
const formSimula = document.getElementById("formSimula");
const simulaMsg = document.getElementById("simulaMsg");
const simulasListEl = document.getElementById("simulasList");

function buildImpactoRow() {
  const row = document.createElement("div");
  row.className = "impacto-row";

  const select = document.createElement("select");
  CELULAS.forEach((c) => {
    const opt = document.createElement("option");
    opt.value = c.id;
    opt.textContent = `${c.abrev} — ${c.nome}`;
    select.appendChild(opt);
  });

  const textarea = document.createElement("textarea");
  textarea.rows = 1;
  textarea.placeholder = "Impacto para esta célula...";

  const removeBtn = document.createElement("button");
  removeBtn.type = "button";
  removeBtn.className = "impacto-row__remove";
  removeBtn.textContent = "remover";
  removeBtn.addEventListener("click", () => row.remove());

  row.appendChild(select);
  row.appendChild(textarea);
  row.appendChild(removeBtn);
  return row;
}

btnAddImpacto.addEventListener("click", () => {
  impactosList.appendChild(buildImpactoRow());
});

// Começa com uma linha de impacto pronta
impactosList.appendChild(buildImpactoRow());

formSimula.addEventListener("submit", async (e) => {
  e.preventDefault();
  simulaMsg.textContent = "";
  simulaMsg.classList.remove("error");

  const reuniaoEl = document.getElementById("simReuniao");
  const gdhEl = document.getElementById("simGdh");
  const decisaoEl = document.getElementById("simDecisao");

  const requiredFieldsValidation = validateRequiredFields([
    { el: reuniaoEl, label: "a Reunião" },
    { el: gdhEl, label: "o GDH" },
    { el: decisaoEl, label: "a Decisão-chave" }
  ]);

  if (requiredFieldsValidation) {
    simulaMsg.textContent = requiredFieldsValidation.message;
    simulaMsg.classList.add("error");
    requiredFieldsValidation.firstEmptyEl.focus();
    return;
  }

  const reuniao = reuniaoEl.value.trim();
  const gdh = gdhEl.value.trim();
  const decisao = decisaoEl.value.trim();
  const autor = document.getElementById("simAutor").value.trim();

  const impactos = [];
  impactosList.querySelectorAll(".impacto-row").forEach((row) => {
    const celula = row.querySelector("select").value;
    const texto = row.querySelector("textarea").value.trim();
    if (texto) impactos.push({ celula, texto });
  });

  const submitBtn = formSimula.querySelector('button[type="submit"]');
  submitBtn.disabled = true;

  try {
    await addDoc(collection(db, "simulas"), {
      reuniao,
      gdh,
      decisao,
      autor,
      impactos,
      createdAt: serverTimestamp()
    });

    simulaMsg.textContent = "Súmula registrada.";
    formSimula.reset();
    impactosList.innerHTML = "";
    impactosList.appendChild(buildImpactoRow());
  } catch (err) {
    console.error("Erro ao registrar súmula:", err);
    simulaMsg.textContent = "Erro ao registrar. Tente novamente.";
    simulaMsg.classList.add("error");
  } finally {
    submitBtn.disabled = false;
  }
});

function renderSimulas() {
  const viewAs = getViewAs();
  let items = simulasCache.slice();

  if (!isCommandView(viewAs)) {
    items = items.filter((s) => (s.impactos || []).some((imp) => imp.celula === viewAs));
  }

  simulasListEl.innerHTML = "";

  if (items.length === 0) {
    simulasListEl.innerHTML = '<p class="empty-state">Nenhuma súmula encontrada para este filtro.</p>';
    return;
  }

  items.forEach((s) => {
    simulasListEl.appendChild(buildSimulaCard(s, viewAs));
  });
}

function buildSimulaCard(s, viewAs) {
  const card = document.createElement("article");
  card.className = "simula-card";

  let impactos = s.impactos || [];
  if (!isCommandView(viewAs)) {
    impactos = impactos.filter((imp) => imp.celula === viewAs);
  }

  const impactosHtml = impactos.map((imp) => `
    <div class="impacto-item">
      <span class="impacto-item__celula mono">${CELULA_MAP[imp.celula] ? CELULA_MAP[imp.celula].abrev : imp.celula}</span>
      <span>${escapeHtml(imp.texto)}</span>
    </div>
  `).join("");

  card.innerHTML = `
    <div class="simula-card__head">
      <span class="simula-card__reuniao">${escapeHtml(s.reuniao || "")}</span>
      <span class="simula-card__gdh mono">${escapeHtml(s.gdh || "—")}</span>
    </div>
    <div class="simula-card__decisao">${escapeHtml(s.decisao || "")}</div>
    ${impactosHtml}
    <div class="injecao-card__footer">
      <span class="injecao-card__autor">${s.autor ? "Autor: " + escapeHtml(s.autor) : ""} · ${formatTimestamp(s.createdAt)}</span>
    </div>
  `;

  return card;
}

// ---------------------------------------------------------------------
// Painel do Comando — KPIs + Kanban
// ---------------------------------------------------------------------

const kpiTotal = document.getElementById("kpiTotal");
const kpiPendentes = document.getElementById("kpiPendentes");
const kpiProntas = document.getElementById("kpiProntas");

const kanbanCols = {
  pendente: document.getElementById("kanbanPendente"),
  pronta: document.getElementById("kanbanPronta")
};

function renderComando() {
  const total = injecoesCache.length;
  const pendentes = injecoesCache.filter((it) => statusPms(it) === "pendente").length;
  const prontas = total - pendentes;

  kpiTotal.textContent = total;
  kpiPendentes.textContent = pendentes;
  kpiProntas.textContent = prontas;

  Object.values(kanbanCols).forEach((col) => (col.innerHTML = ""));

  injecoesCache.forEach((it) => {
    kanbanCols[statusPms(it)].appendChild(buildKanbanCard(it));
  });
}

function buildKanbanCard(it) {
  const card = document.createElement("div");
  card.className = "kanban-card";
  card.dataset.urgencia = it.urgencia;

  // Situação por célula: quem já respondeu (com link para baixar) e quem
  // ainda está pendente.
  const celulasHtml = celulasRoteadas(it).map((c) => {
    const respostas = respostasDaCelula(it, c.id);
    const respondeu = respostas.length > 0;
    return `
      <div class="kanban-celula kanban-celula--${respondeu ? "ok" : "pendente"}">
        <span class="kanban-celula__status">${respondeu ? "✓" : "⏳"}</span>
        <span class="kanban-celula__abrev mono">${c.abrev}</span>
        <span class="kanban-celula__info">${respondeu
          ? respostas.map(anexoLinkHtml).join("")
          : '<span class="kanban-celula__pendente">pendente</span>'}</span>
      </div>
    `;
  }).join("");

  card.innerHTML = `
    <div class="kanban-card__head">
      <span class="pill pill--${it.urgencia}">${URGENCIA_LABEL[it.urgencia] || it.urgencia}</span>
      <span class="kanban-card__gdh mono">${escapeHtml(it.gdh || "—")}</span>
    </div>
    <div class="kanban-card__resumo">${escapeHtml(it.resumo || "")}</div>
    ${it.anexo ? `<div class="kanban-card__anexo">PMS: ${anexoLinkHtml(it.anexo)}</div>` : ""}
    <div class="kanban-celulas">${celulasHtml}</div>
  `;

  return card;
}

// ---------------------------------------------------------------------
// Listeners Firestore (tempo real)
// ---------------------------------------------------------------------

let listenersStarted = false;

function startListeners() {
  if (listenersStarted) return;
  listenersStarted = true;

  const injecoesQuery = query(collection(db, "injecoes"), orderBy("createdAt", "desc"));
  onSnapshot(
    injecoesQuery,
    (snapshot) => {
      injecoesCache = snapshot.docs.map((d) => ({ id: d.id, ...d.data() }));
      renderDirex();
      renderComando();
    },
    (err) => {
      console.error("Erro no listener de PMS:", err);
      setConn("offline");
    }
  );

  const simulasQuery = query(collection(db, "simulas"), orderBy("createdAt", "desc"));
  onSnapshot(
    simulasQuery,
    (snapshot) => {
      simulasCache = snapshot.docs.map((d) => ({ id: d.id, ...d.data() }));
      renderSimulas();
    },
    (err) => {
      console.error("Erro no listener de súmulas:", err);
      setConn("offline");
    }
  );
}
