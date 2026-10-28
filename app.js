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
// fica "pronta" só quando TODAS as células roteadas enviaram resposta.
// ---------------------------------------------------------------------

function celulasRoteadas(it) {
  return CELULAS.filter((c) => it.celulas && it.celulas[c.id] && it.celulas[c.id].sel);
}

// Respostas antigas são identificadas pelo arquivo; as novas têm "id"
// próprio, porque podem ser só texto.
function idResposta(r) {
  return r.id || r.arquivoId;
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
let mensagensCache = [];

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
  renderMensagens();
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
  const ext = (anexo.nome || "").split(".").pop().toUpperCase();
  return `<span class="arquivo"><span class="arquivo__tipo">${escapeHtml(ext)}</span>` +
    `<button type="button" class="anexo-link" data-arquivo-id="${escapeHtml(anexo.arquivoId)}">${escapeHtml(anexo.nome)}</button>` +
    `<span class="anexo-link__tam">${formatBytes(anexo.tamanho)}</span></span>`;
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
// Seletor de anexo (DOCX/PDF) dos formulários. Na edição, `carregar`
// mostra o arquivo atual, que pode ser mantido, substituído ou removido.
// ---------------------------------------------------------------------

function criarAnexoPicker(prefixo, msgEl) {
  const input = document.getElementById(`${prefixo}Anexo`);
  const nomeEl = document.getElementById(`${prefixo}AnexoNome`);
  const limparBtn = document.getElementById(`${prefixo}AnexoLimpar`);
  const botao = document.getElementById(`${prefixo}AnexoBtn`);
  const picker = { atual: null, removerAtual: false };

  picker.arquivo = () => input.files[0] || null;
  picker.mantido = () => (picker.atual && !picker.removerAtual ? picker.atual : null);
  picker.reset = () => {
    input.value = "";
    const atual = picker.mantido();
    nomeEl.textContent = atual
      ? `Atual: ${atual.nome} (${formatBytes(atual.tamanho)})`
      : "Nenhum arquivo selecionado";
    limparBtn.hidden = !atual;
    botao.textContent = atual ? "Substituir arquivo (DOCX/PDF, até 10 MB)" : "Anexar arquivo (DOCX/PDF, até 10 MB)";
  };
  picker.carregar = (anexo) => {
    picker.atual = anexo || null;
    picker.removerAtual = false;
    picker.reset();
  };

  input.addEventListener("change", () => {
    msgEl.textContent = "";
    msgEl.classList.remove("error");
    const file = input.files[0];
    if (!file) return picker.reset();
    const erro = validateAnexo(file);
    if (erro) {
      picker.reset();
      msgEl.textContent = erro;
      msgEl.classList.add("error");
      return;
    }
    nomeEl.textContent = `${file.name} (${formatBytes(file.size)})`;
    limparBtn.hidden = false;
  });

  // "remover": descarta o arquivo recém-escolhido; sem ele, na edição,
  // marca o arquivo atual para ser apagado ao salvar.
  limparBtn.addEventListener("click", () => {
    if (!input.files[0] && picker.atual) picker.removerAtual = true;
    picker.reset();
  });

  return picker;
}

// Anexo final ao salvar: novo upload, o atual mantido ou nenhum. Devolve
// também o anexo anterior que deixou de ser usado, para ser apagado.
async function resolverAnexo(picker, msgEl) {
  let anexo = picker.mantido();
  if (picker.arquivo()) {
    msgEl.textContent = "Enviando arquivo...";
    anexo = await uploadAnexo(picker.arquivo());
  }
  const descartado = picker.atual && (!anexo || anexo.arquivoId !== picker.atual.arquivoId)
    ? picker.atual
    : null;
  return { anexo, descartado };
}

function limparMsg(msgEl) {
  msgEl.textContent = "";
  msgEl.classList.remove("error");
}

function mostrarErro(msgEl, texto) {
  msgEl.textContent = texto;
  msgEl.classList.add("error");
}

// ---------------------------------------------------------------------
// Novo Evento: escolha entre PMS e Mensagem
// ---------------------------------------------------------------------

const formNovaInjecao = document.getElementById("formNovaInjecao");
const formMensagem = document.getElementById("formMensagem");
const tipoEventoRadios = document.querySelectorAll('input[name="tipoEvento"]');

function escolherTipoEvento(tipo) {
  tipoEventoRadios.forEach((r) => (r.checked = r.value === tipo));
  formNovaInjecao.hidden = tipo !== "pms";
  formMensagem.hidden = tipo !== "mensagem";
}

tipoEventoRadios.forEach((r) => {
  r.addEventListener("change", () => escolherTipoEvento(r.value));
});

function abrirFormulario(tipo) {
  escolherTipoEvento(tipo);
  ativarAba("nova");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

// ---------------------------------------------------------------------
// Formulário: PMS (coleção "injecoes")
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

const novaInjecaoMsg = document.getElementById("novaInjecaoMsg");
const formNovaTitulo = document.getElementById("formNovaTitulo");
const btnCancelarEdicao = document.getElementById("btnCancelarEdicao");
const btnSubmitPms = formNovaInjecao.querySelector('button[type="submit"]');
const pickerPms = criarAnexoPicker("in", novaInjecaoMsg);

// Edição: o mesmo formulário é reaproveitado para alterar um PMS já
// registrado; `edicao` guarda o id desse PMS.
let edicao = null; // { id }

function limparFormPms() {
  formNovaInjecao.reset();
  CELULAS.forEach((c) => {
    const ref = celulaBlockRefs[c.id];
    ref.textarea.hidden = true;
    ref.textarea.value = "";
  });
  formNovaInjecao.querySelector('input[name="urgencia"][value="critico"]').checked = true;
  edicao = null;
  pickerPms.carregar(null);
  formNovaTitulo.textContent = "Registrar PMS";
  btnSubmitPms.textContent = "Enviar PMS";
  btnCancelarEdicao.hidden = true;
}

function iniciarEdicao(it) {
  limparFormPms();
  edicao = { id: it.id };

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

  pickerPms.carregar(it.anexo);
  formNovaTitulo.textContent = `Editar PMS ${it.gdh || ""}`.trim();
  btnSubmitPms.textContent = "Salvar alterações";
  btnCancelarEdicao.hidden = false;
  limparMsg(novaInjecaoMsg);
  abrirFormulario("pms");
}

btnCancelarEdicao.addEventListener("click", () => {
  limparFormPms();
  limparMsg(novaInjecaoMsg);
});

formNovaInjecao.addEventListener("submit", async (e) => {
  e.preventDefault();
  limparMsg(novaInjecaoMsg);

  const gdhEl = document.getElementById("inGdh");
  const origemEl = document.getElementById("inOrigem");
  const resumoEl = document.getElementById("inResumo");

  const requiredFieldsValidation = validateRequiredFields([
    { el: gdhEl, label: "o GDH" },
    { el: origemEl, label: "a Origem" },
    { el: resumoEl, label: "o resumo da situação" }
  ]);

  if (requiredFieldsValidation) {
    mostrarErro(novaInjecaoMsg, requiredFieldsValidation.message);
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
    mostrarErro(novaInjecaoMsg, "Selecione ao menos uma célula para rotear.");
    return;
  }

  if (edicao) return salvarEdicao({ gdh, origem, autor, urgencia, resumo, celulas });

  btnSubmitPms.disabled = true;
  try {
    const { anexo } = await resolverAnexo(pickerPms, novaInjecaoMsg);

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

    limparFormPms();
    novaInjecaoMsg.textContent = "PMS enviado.";
  } catch (err) {
    console.error("Erro ao registrar PMS:", err);
    mostrarErro(novaInjecaoMsg, "Erro ao enviar. Tente novamente.");
  } finally {
    btnSubmitPms.disabled = false;
  }
});

async function salvarEdicao(dados) {
  const atual = injecoesCache.find((x) => x.id === edicao.id);
  if (!atual) {
    mostrarErro(novaInjecaoMsg, "Este PMS não existe mais (pode ter sido excluído).");
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
    const { anexo, descartado } = await resolverAnexo(pickerPms, novaInjecaoMsg);

    const alteracoes = { ...dados, anexo, editadoEm: serverTimestamp() };
    if (orfas.length > 0) alteracoes.respostas = arrayRemove(...orfas);
    await updateDoc(doc(db, "injecoes", edicao.id), alteracoes);

    // Melhor esforço: apaga o arquivo substituído/removido e os das respostas descartadas.
    const lixo = (descartado ? [...orfas, descartado] : orfas).filter((a) => a.arquivoId);
    await Promise.allSettled(lixo.map(deleteAnexo));

    limparFormPms();
    novaInjecaoMsg.textContent = "PMS atualizado.";
  } catch (err) {
    console.error("Erro ao salvar PMS:", err);
    mostrarErro(novaInjecaoMsg, "Erro ao salvar. Tente novamente.");
  } finally {
    btnSubmitPms.disabled = false;
  }
}

async function excluirPms(it, btn) {
  const ok = window.confirm(
    `Excluir o PMS ${it.gdh || ""}? Os arquivos anexados também serão apagados. Esta ação não pode ser desfeita.`
  );
  if (!ok) return;
  btn.disabled = true;
  try {
    await deleteDoc(doc(db, "injecoes", it.id));
    const arquivos = [it.anexo, ...(it.respostas || [])].filter((a) => a && a.arquivoId);
    // Melhor esforço: falha ao apagar um arquivo não impede a exclusão do PMS.
    await Promise.allSettled(arquivos.map(deleteAnexo));
  } catch (err) {
    console.error("Erro ao excluir PMS:", err);
    btn.disabled = false;
    window.alert("Erro ao excluir. Tente novamente.");
  }
}

// ---------------------------------------------------------------------
// Formulário: Mensagem (coleção "mensagens") — só para consulta das
// seções de destino; sem prazo, urgência ou resposta.
// ---------------------------------------------------------------------

const mensagemMsg = document.getElementById("mensagemMsg");
const formMensagemTitulo = document.getElementById("formMensagemTitulo");
const btnCancelarEdicaoMsg = document.getElementById("btnCancelarEdicaoMsg");
const btnSubmitMsg = formMensagem.querySelector('button[type="submit"]');
const btnSelecionarTodas = document.getElementById("btnSelecionarTodas");
const destinosGrid = document.getElementById("destinosGrid");
const destinoChecks = {}; // id -> checkbox
const pickerMsg = criarAnexoPicker("msg", mensagemMsg);

let edicaoMsg = null; // { id }

CELULAS.forEach((c) => {
  const opt = document.createElement("label");
  opt.className = "celula-block celula-block__check";
  opt.innerHTML = `
    <input type="checkbox" class="celula-check">
    <span class="celula-block__abrev">${c.abrev}</span>
    <span class="celula-block__nome">${escapeHtml(c.nome)}</span>
  `;
  const checkbox = opt.querySelector("input");
  checkbox.addEventListener("change", atualizarSelecionarTodas);
  destinoChecks[c.id] = checkbox;
  destinosGrid.appendChild(opt);
});

function todasSelecionadas() {
  return CELULAS.every((c) => destinoChecks[c.id].checked);
}

function atualizarSelecionarTodas() {
  btnSelecionarTodas.textContent = todasSelecionadas() ? "Desmarcar todas" : "Selecionar todas";
}

btnSelecionarTodas.addEventListener("click", () => {
  const marcar = !todasSelecionadas();
  CELULAS.forEach((c) => (destinoChecks[c.id].checked = marcar));
  atualizarSelecionarTodas();
});

function limparFormMensagem() {
  formMensagem.reset();
  edicaoMsg = null;
  pickerMsg.carregar(null);
  atualizarSelecionarTodas();
  formMensagemTitulo.textContent = "Registrar Mensagem";
  btnSubmitMsg.textContent = "Enviar Mensagem";
  btnCancelarEdicaoMsg.hidden = true;
}

function iniciarEdicaoMensagem(m) {
  limparFormMensagem();
  edicaoMsg = { id: m.id };

  document.getElementById("msgGdh").value = m.gdh || "";
  document.getElementById("msgOrigem").value = m.origem || "";
  document.getElementById("msgAutor").value = m.autor || "";
  document.getElementById("msgAssunto").value = m.assunto || "";
  document.getElementById("msgTexto").value = m.texto || "";
  CELULAS.forEach((c) => (destinoChecks[c.id].checked = (m.destinos || []).includes(c.id)));
  atualizarSelecionarTodas();

  pickerMsg.carregar(m.anexo);
  formMensagemTitulo.textContent = `Editar Mensagem ${m.gdh || ""}`.trim();
  btnSubmitMsg.textContent = "Salvar alterações";
  btnCancelarEdicaoMsg.hidden = false;
  limparMsg(mensagemMsg);
  abrirFormulario("mensagem");
}

btnCancelarEdicaoMsg.addEventListener("click", () => {
  limparFormMensagem();
  limparMsg(mensagemMsg);
});

formMensagem.addEventListener("submit", async (e) => {
  e.preventDefault();
  limparMsg(mensagemMsg);

  const gdhEl = document.getElementById("msgGdh");
  const origemEl = document.getElementById("msgOrigem");
  const assuntoEl = document.getElementById("msgAssunto");

  const requiredFieldsValidation = validateRequiredFields([
    { el: gdhEl, label: "o GDH" },
    { el: origemEl, label: "a Origem" },
    { el: assuntoEl, label: "o Assunto" }
  ]);

  if (requiredFieldsValidation) {
    mostrarErro(mensagemMsg, requiredFieldsValidation.message);
    requiredFieldsValidation.firstEmptyEl.focus();
    return;
  }

  const texto = document.getElementById("msgTexto").value.trim();
  if (!texto && !pickerMsg.mantido() && !pickerMsg.arquivo()) {
    mostrarErro(mensagemMsg, "Escreva o texto da mensagem ou anexe um arquivo.");
    return;
  }

  const destinos = CELULAS.filter((c) => destinoChecks[c.id].checked).map((c) => c.id);
  if (destinos.length === 0) {
    mostrarErro(mensagemMsg, "Selecione ao menos uma seção de destino.");
    return;
  }

  const dados = {
    gdh: gdhEl.value.trim(),
    origem: origemEl.value.trim(),
    autor: document.getElementById("msgAutor").value.trim(),
    assunto: assuntoEl.value.trim(),
    texto,
    destinos
  };

  btnSubmitMsg.disabled = true;
  try {
    const { anexo, descartado } = await resolverAnexo(pickerMsg, mensagemMsg);
    if (edicaoMsg) {
      await updateDoc(doc(db, "mensagens", edicaoMsg.id), { ...dados, anexo, editadoEm: serverTimestamp() });
      if (descartado) await deleteAnexo(descartado).catch((err) => console.error("Erro ao apagar arquivo antigo:", err));
      limparFormMensagem();
      mensagemMsg.textContent = "Mensagem atualizada.";
    } else {
      await addDoc(collection(db, "mensagens"), { ...dados, anexo, createdAt: serverTimestamp() });
      limparFormMensagem();
      mensagemMsg.textContent = "Mensagem enviada.";
    }
  } catch (err) {
    console.error("Erro ao registrar mensagem:", err);
    mostrarErro(mensagemMsg, edicaoMsg ? "Erro ao salvar. Tente novamente." : "Erro ao enviar. Tente novamente.");
  } finally {
    btnSubmitMsg.disabled = false;
  }
});

async function excluirMensagem(m, btn) {
  const ok = window.confirm(
    `Excluir a mensagem ${m.gdh || ""}${m.assunto ? ` (${m.assunto})` : ""}? O arquivo anexado também será apagado. Esta ação não pode ser desfeita.`
  );
  if (!ok) return;
  btn.disabled = true;
  try {
    await deleteDoc(doc(db, "mensagens", m.id));
    if (m.anexo) await deleteAnexo(m.anexo).catch((err) => console.error("Erro ao apagar arquivo da mensagem:", err));
  } catch (err) {
    console.error("Erro ao excluir mensagem:", err);
    btn.disabled = false;
    window.alert("Erro ao excluir. Tente novamente.");
  }
}

// Aceita Timestamp do Firestore ou milissegundos (enviadoEm das respostas).
function formatTimestamp(ts) {
  const d = typeof ts === "number" ? new Date(ts) : ts && typeof ts.toDate === "function" ? ts.toDate() : null;
  if (!d) return "—";
  return d.toLocaleString("pt-BR", {
    day: "2-digit", month: "2-digit", year: "2-digit",
    hour: "2-digit", minute: "2-digit"
  });
}

function autorDataHtml(item) {
  return (item.autor ? "Autor: " + escapeHtml(item.autor) + " · " : "") +
    formatTimestamp(item.createdAt) +
    (item.editadoEm ? " · editado " + formatTimestamp(item.editadoEm) : "");
}

// Peças comuns a todos os cartões (Controle PMS, Controle Mensagens,
// listas e Painel do Comando), para manter o mesmo padrão visual.

function urgenciaPill(urgencia) {
  return `<span class="pill pill--${urgencia}">${URGENCIA_LABEL[urgencia] || escapeHtml(urgencia)}</span>`;
}

function statusPill(status) {
  return `<span class="pill pill--status-${status}">${STATUS_LABEL[status]}</span>`;
}

function cabecalhoHtml(pillsHtml, gdh, origem) {
  return `
    <div class="injecao-card__head">
      ${pillsHtml}
      <span class="injecao-card__gdh mono">${escapeHtml(gdh || "—")}</span>
      ${origem ? `<span class="injecao-card__origem">${escapeHtml(origem)}</span>` : ""}
    </div>
  `;
}

function infoHtml(rotulo, valorHtml) {
  return `<div class="info"><span class="rotulo">${rotulo}</span><div class="info__valor">${valorHtml}</div></div>`;
}

function rodapeHtml(item, comAcoes) {
  return `
    <div class="injecao-card__footer">
      <span class="injecao-card__autor">${autorDataHtml(item)}</span>
      ${comAcoes ? `
        <span class="injecao-card__acoes">
          <button type="button" class="btn btn--ghost btn--small btn-editar">Editar</button>
          <button type="button" class="btn btn--danger-ghost btn--small btn-excluir">Excluir</button>
        </span>
      ` : ""}
    </div>
  `;
}

// Resposta de uma célula: texto (para quem, assunto...) e/ou arquivo.
function respostaHtml(r, comRemover) {
  return `
    <div class="resposta">
      ${r.texto ? `<div class="resposta__texto">${escapeHtml(r.texto)}</div>` : ""}
      <div class="resposta__rodape">
        ${r.arquivoId ? anexoLinkHtml(r) : ""}
        ${r.enviadoEm ? `<span class="resposta__data">enviada ${formatTimestamp(r.enviadoEm)}</span>` : ""}
        ${comRemover ? `<button type="button" class="btn-remover resposta-remover" data-resposta-id="${escapeHtml(idResposta(r))}">remover</button>` : ""}
      </div>
    </div>
  `;
}

// ---------------------------------------------------------------------
// Controles (PMS e Mensagens) em formato de dashboard: um quadrado por
// seção; clicar abre a seção inteira; "Listar" mostra todos os
// lançamentos, do mais recente para o mais antigo. Na visão de uma
// célula ("Ver como"), a seção dela abre direto.
// ---------------------------------------------------------------------

const TILE_MAX_ITENS = 5;

function criarControle(cfg) {
  const conteudo = document.getElementById(`${cfg.prefixo}Conteudo`);
  const btnListar = document.getElementById(`${cfg.prefixo}BtnListar`);
  const btnVoltar = document.getElementById(`${cfg.prefixo}BtnVoltar`);
  const titulo = document.getElementById(`${cfg.prefixo}Titulo`);
  const estado = { modo: "grade", celula: null };

  function ir(modo, celula = null) {
    estado.modo = modo;
    estado.celula = celula;
    render();
    window.scrollTo({ top: 0 });
  }

  btnListar.addEventListener("click", () => ir("lista"));
  btnVoltar.addEventListener("click", () => ir("grade"));

  function render() {
    const viewAs = getViewAs();
    const unica = !isCommandView(viewAs);
    let { modo, celula } = estado;
    if (unica && modo !== "lista") {
      modo = "secao";
      celula = viewAs;
    }

    btnListar.hidden = modo === "lista";
    btnVoltar.hidden = modo === "grade" || (modo === "secao" && unica);
    btnVoltar.textContent = unica ? "← Voltar à seção" : "← Voltar às seções";
    titulo.textContent =
      modo === "lista" ? cfg.tituloLista :
      modo === "secao" ? `${cfg.titulo} · ${CELULA_MAP[celula].nome}` : cfg.titulo;

    // Preserva o foco e o cursor de quem está escrevendo numa caixa de texto.
    const ativo = document.activeElement;
    const focoChave = ativo && conteudo.contains(ativo) && ativo.dataset.chave ? ativo.dataset.chave : null;
    const cursor = focoChave ? [ativo.selectionStart, ativo.selectionEnd] : null;

    conteudo.innerHTML = "";
    if (modo === "grade") {
      conteudo.appendChild(buildGrade());
    } else if (modo === "secao") {
      conteudo.appendChild(buildLista(cfg.itensDaCelula(celula), (it) => cfg.cardSecao(it, celula), cfg.vazioSecao));
    } else {
      conteudo.appendChild(buildLista(cfg.itensLista(), cfg.linhaLista, cfg.vazioLista));
    }

    if (focoChave) {
      const el = conteudo.querySelector(`[data-chave="${focoChave}"]`);
      if (el) {
        el.focus();
        el.setSelectionRange(cursor[0], cursor[1]);
      }
    }
  }

  function buildLista(items, build, vazio) {
    const lista = document.createElement("div");
    lista.className = "cards-list";
    if (items.length === 0) {
      lista.innerHTML = `<p class="empty-state">${vazio}</p>`;
    } else {
      items.forEach((it) => lista.appendChild(build(it)));
    }
    return lista;
  }

  function buildGrade() {
    const grade = document.createElement("div");
    grade.className = "secoes-grade";
    CELULAS.forEach((c) => {
      const { contagem, alerta, linhas } = cfg.tile(c.id);
      const extras = linhas.length - TILE_MAX_ITENS;
      const tile = document.createElement("button");
      tile.type = "button";
      tile.className = "secao-tile";
      tile.innerHTML = `
        <span class="secao-tile__head">
          <span class="secao-tile__abrev mono">${c.abrev}</span>
          <span class="secao-tile__nome">${escapeHtml(c.nome)}</span>
        </span>
        <span class="secao-tile__contagem${alerta ? " secao-tile__contagem--alerta" : ""}">${contagem}</span>
        <span class="secao-tile__itens">
          ${linhas.length === 0 ? `<span class="secao-tile__vazio">${cfg.vazioTile}</span>` : linhas.slice(0, TILE_MAX_ITENS).join("")}
          ${extras > 0 ? `<span class="secao-tile__mais">+${extras} mais</span>` : ""}
        </span>
        <span class="secao-tile__abrir">Abrir seção →</span>
      `;
      tile.addEventListener("click", () => ir("secao", c.id));
      grade.appendChild(tile);
    });
    return grade;
  }

  return { render };
}

function plural(n, singular, pluralTxt) {
  return `${n} ${n === 1 ? singular : pluralTxt}`;
}

// ---------------------------------------------------------------------
// Controle PMS
// ---------------------------------------------------------------------

const statusFilter = document.getElementById("statusFilter");

// Rascunhos das respostas (texto e arquivo escolhido, ainda não
// enviados): sobrevivem às re-renderizações em tempo real.
const rascunhosResposta = new Map(); // "<pmsId>:<celulaId>" -> { texto, arquivo }
const respostasEnviando = new Set(); // chaves com envio em andamento

function celulaRespondeu(it, celulaId) {
  return respostasDaCelula(it, celulaId).length > 0;
}

function pmsDaCelula(celulaId) {
  return injecoesCache.filter((it) => it.celulas && it.celulas[celulaId] && it.celulas[celulaId].sel);
}

// Na seção, o filtro de status vale para a resposta daquela célula.
function pmsDaCelulaFiltrados(celulaId) {
  const statusVal = statusFilter.value;
  const items = pmsDaCelula(celulaId);
  if (statusVal === "todos") return items;
  return items.filter((it) => (celulaRespondeu(it, celulaId) ? "pronta" : "pendente") === statusVal);
}

const controlePms = criarControle({
  prefixo: "direx",
  titulo: "Controle PMS",
  tituloLista: "Controle PMS · PMS lançados",
  vazioTile: "Nenhum PMS",
  vazioSecao: "Nenhum PMS desta seção para este filtro.",
  vazioLista: "Nenhum PMS encontrado para este filtro.",
  itensDaCelula: pmsDaCelulaFiltrados,
  cardSecao: buildInjecaoCard,
  itensLista: () => statusFilter.value === "todos"
    ? injecoesCache
    : injecoesCache.filter((it) => statusPms(it) === statusFilter.value),
  linhaLista: buildPmsLinha,
  tile: (celulaId) => {
    const pendentes = pmsDaCelula(celulaId).filter((it) => !celulaRespondeu(it, celulaId)).length;
    return {
      contagem: `${plural(pendentes, "pendente", "pendentes")} · ${pmsDaCelula(celulaId).length} PMS`,
      alerta: pendentes > 0,
      linhas: pmsDaCelulaFiltrados(celulaId).map((it) => {
        const ok = celulaRespondeu(it, celulaId);
        return `
          <span class="secao-tile__item">
            <span class="secao-tile__status secao-tile__status--${ok ? "ok" : "pendente"}">${ok ? "✓" : "⏳"}</span>
            <span class="secao-tile__gdh mono">${escapeHtml(it.gdh || "—")}</span>
            <span class="secao-tile__resumo">${escapeHtml(it.resumo || "")}</span>
          </span>
        `;
      })
    };
  }
});

statusFilter.addEventListener("change", renderDirex);

function renderDirex() {
  controlePms.render();
}

// Linha do "Listar PMS": resumo do lançamento e situação de cada célula.
function buildPmsLinha(it) {
  const row = document.createElement("article");
  row.className = "injecao-card lancamento";
  row.dataset.urgencia = it.urgencia;

  const celulasHtml = celulasRoteadas(it).map((c) => {
    const ok = celulaRespondeu(it, c.id);
    return `<span class="chip chip--${ok ? "ok" : "pendente"}">${ok ? "✓" : "⏳"} ${c.abrev}</span>`;
  }).join("");

  row.innerHTML = `
    ${cabecalhoHtml(urgenciaPill(it.urgencia) + statusPill(statusPms(it)), it.gdh, it.origem)}
    <div class="injecao-card__resumo lancamento__texto">${escapeHtml(it.resumo || "")}</div>
    <div class="chips">${celulasHtml}</div>
    ${it.anexo ? `<div class="info-grid">${infoHtml("Arquivo do PMS", anexoLinkHtml(it.anexo))}</div>` : ""}
    ${rodapeHtml(it, true)}
  `;

  row.querySelector(".btn-editar").addEventListener("click", () => iniciarEdicao(it));
  const btnExcluir = row.querySelector(".btn-excluir");
  btnExcluir.addEventListener("click", () => excluirPms(it, btnExcluir));
  return row;
}

// Cartão de um PMS dentro da seção de uma célula: mostra o prazo, as
// respostas e o envio de resposta dessa célula (texto e/ou arquivo, que
// só são gravados ao clicar em "Enviar resposta").
function buildInjecaoCard(it, celulaId) {
  const card = document.createElement("article");
  card.className = "injecao-card";
  card.dataset.urgencia = it.urgencia;

  const celula = CELULA_MAP[celulaId];
  const roteadas = celulasRoteadas(it);
  const qtdRespondidas = roteadas.filter((c) => celulaRespondeu(it, c.id)).length;
  const status = celulaRespondeu(it, celulaId) ? "pronta" : "pendente";
  const chipsHtml = roteadas
    .map((c) => `<span class="chip${c.id === celulaId ? " chip--atual" : ""}">${c.abrev}</span>`)
    .join("");

  const prazo = it.celulas[celulaId].texto;
  const infos =
    (prazo ? infoHtml(`Prazo da ${celula.abrev}`, escapeHtml(prazo)) : "") +
    (it.anexo ? infoHtml("Arquivo do PMS", anexoLinkHtml(it.anexo)) : "");

  const respostas = respostasDaCelula(it, celulaId);
  const chave = `${it.id}:${celulaId}`;

  card.innerHTML = `
    ${cabecalhoHtml(urgenciaPill(it.urgencia) + statusPill(status), it.gdh, it.origem)}
    <div class="injecao-card__resumo">${escapeHtml(it.resumo || "")}</div>
    <div class="chips">${chipsHtml}<span class="chips__nota">${qtdRespondidas}/${roteadas.length} células responderam</span></div>
    ${infos ? `<div class="info-grid">${infos}</div>` : ""}
    <div class="bloco">
      <span class="rotulo">Resposta da ${celula.abrev}</span>
      ${respostas.length ? respostas.map((r) => respostaHtml(r, true)).join("") : '<p class="vazio">Nenhuma resposta enviada.</p>'}
      <div class="resposta-form">
        <span class="rotulo">Nova resposta</span>
        <textarea class="resposta-texto" rows="3" data-chave="${chave}"
          placeholder="Para quem:&#10;Assunto:&#10;Outras informações:"></textarea>
        <div class="resposta-form__acoes">
          <label class="btn btn--ghost btn--small">
            Escolher arquivo (DOCX/PDF)
            <input type="file" class="anexo-picker__input resposta-file" accept="${ANEXO_ACCEPT}">
          </label>
          <span class="anexo-picker__nome resposta-arquivo"></span>
          <button type="button" class="btn-remover resposta-limpar" hidden>remover</button>
          <button type="button" class="btn btn--primary btn--small resposta-enviar">Enviar resposta</button>
        </div>
        <span class="form-msg resposta-msg"></span>
      </div>
    </div>
    ${rodapeHtml(it, true)}
  `;

  card.querySelectorAll(".resposta-remover").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const resposta = (it.respostas || []).find((r) => idResposta(r) === btn.dataset.respostaId);
      if (!resposta) return;
      const abrev = CELULA_MAP[resposta.celula] ? CELULA_MAP[resposta.celula].abrev : resposta.celula;
      const qual = resposta.nome ? `a resposta "${resposta.nome}"` : "esta resposta";
      const ok = window.confirm(`Remover ${qual} da ${abrev}? O PMS volta a ficar pendente para esta célula se não houver outra resposta dela.`);
      if (!ok) return;
      btn.disabled = true;
      try {
        // arrayRemove exige o objeto idêntico ao gravado.
        await updateDoc(doc(db, "injecoes", it.id), { respostas: arrayRemove(resposta) });
        if (resposta.arquivoId) {
          await deleteAnexo(resposta).catch((err) => console.error("Erro ao apagar arquivo da resposta:", err));
        }
      } catch (err) {
        console.error("Erro ao remover resposta:", err);
        btn.disabled = false;
        window.alert("Erro ao remover. Tente novamente.");
      }
    });
  });

  const rascunho = rascunhosResposta.get(chave) || { texto: "", arquivo: null };
  const respostaTexto = card.querySelector(".resposta-texto");
  const respostaFile = card.querySelector(".resposta-file");
  const arquivoNome = card.querySelector(".resposta-arquivo");
  const btnLimpar = card.querySelector(".resposta-limpar");
  const btnEnviar = card.querySelector(".resposta-enviar");
  const respostaMsg = card.querySelector(".resposta-msg");

  function mostrarArquivo() {
    arquivoNome.textContent = rascunho.arquivo
      ? `${rascunho.arquivo.name} (${formatBytes(rascunho.arquivo.size)})`
      : "Nenhum arquivo selecionado";
    btnLimpar.hidden = !rascunho.arquivo;
  }

  respostaTexto.value = rascunho.texto;
  mostrarArquivo();
  if (respostasEnviando.has(chave)) {
    respostaTexto.disabled = btnEnviar.disabled = btnLimpar.disabled = respostaFile.disabled = true;
    respostaMsg.textContent = "Enviando resposta...";
  }

  respostaTexto.addEventListener("input", () => {
    rascunho.texto = respostaTexto.value;
    rascunhosResposta.set(chave, rascunho);
  });

  respostaFile.addEventListener("change", () => {
    const file = respostaFile.files[0];
    respostaFile.value = "";
    if (!file) return;
    limparMsg(respostaMsg);
    const erro = validateAnexo(file);
    if (erro) return mostrarErro(respostaMsg, erro);
    rascunho.arquivo = file;
    rascunhosResposta.set(chave, rascunho);
    mostrarArquivo();
  });

  btnLimpar.addEventListener("click", () => {
    rascunho.arquivo = null;
    rascunhosResposta.set(chave, rascunho);
    mostrarArquivo();
  });

  btnEnviar.addEventListener("click", async () => {
    limparMsg(respostaMsg);
    const texto = respostaTexto.value.trim();
    const arquivo = rascunho.arquivo;
    if (!texto && !arquivo) {
      mostrarErro(respostaMsg, "Escreva a resposta ou escolha um arquivo.");
      respostaTexto.focus();
      return;
    }

    respostasEnviando.add(chave);
    respostaTexto.disabled = btnEnviar.disabled = btnLimpar.disabled = respostaFile.disabled = true;
    respostaMsg.textContent = "Enviando resposta...";
    let aviso = null;
    try {
      const anexo = arquivo ? await uploadAnexo(arquivo) : {};
      rascunhosResposta.delete(chave);
      await updateDoc(doc(db, "injecoes", it.id), {
        respostas: arrayUnion({ id: crypto.randomUUID(), ...anexo, celula: celulaId, texto, enviadoEm: Date.now() })
      });
      aviso = { texto: "Resposta enviada.", erro: false };
    } catch (err) {
      console.error("Erro ao enviar resposta:", err);
      rascunhosResposta.set(chave, { texto, arquivo });
      aviso = { texto: "Erro ao enviar. Tente novamente.", erro: true };
    } finally {
      respostasEnviando.delete(chave);
    }

    // O cartão pode ter sido recriado pelo tempo real durante o envio:
    // re-renderiza e mostra o aviso no cartão atual.
    renderDirex();
    const atual = document.querySelector(`.resposta-texto[data-chave="${chave}"]`);
    const msgAtual = atual && atual.closest(".resposta-form").querySelector(".resposta-msg");
    if (msgAtual) {
      if (aviso.erro) mostrarErro(msgAtual, aviso.texto);
      else msgAtual.textContent = aviso.texto;
    }
  });

  card.querySelector(".btn-editar").addEventListener("click", () => iniciarEdicao(it));

  const btnExcluir = card.querySelector(".btn-excluir");
  btnExcluir.addEventListener("click", () => excluirPms(it, btnExcluir));

  return card;
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str == null ? "" : String(str);
  return div.innerHTML;
}

// ---------------------------------------------------------------------
// Controle Mensagens — só consulta nas seções; editar/excluir no
// "Listar Mensagens".
// ---------------------------------------------------------------------

function mensagensDaCelula(celulaId) {
  return mensagensCache.filter((m) => (m.destinos || []).includes(celulaId));
}

const controleMensagens = criarControle({
  prefixo: "msgs",
  titulo: "Controle Mensagens",
  tituloLista: "Controle Mensagens · Mensagens lançadas",
  vazioTile: "Nenhuma mensagem",
  vazioSecao: "Nenhuma mensagem para esta seção.",
  vazioLista: "Nenhuma mensagem registrada.",
  itensDaCelula: mensagensDaCelula,
  cardSecao: (m, celulaId) => buildMensagemCard(m, celulaId, false),
  itensLista: () => mensagensCache,
  linhaLista: (m) => buildMensagemCard(m, null, true),
  tile: (celulaId) => {
    const ms = mensagensDaCelula(celulaId);
    return {
      contagem: plural(ms.length, "mensagem", "mensagens"),
      alerta: false,
      linhas: ms.map((m) => `
        <span class="secao-tile__item">
          <span class="secao-tile__gdh mono">${escapeHtml(m.gdh || "—")}</span>
          <span class="secao-tile__resumo">${escapeHtml(m.assunto || "")}</span>
        </span>
      `)
    };
  }
});

function renderMensagens() {
  controleMensagens.render();
}

function destinosChipsHtml(m, celulaAtual) {
  const destinos = m.destinos || [];
  if (CELULAS.every((c) => destinos.includes(c.id))) {
    return '<span class="chip">Todas as seções</span>';
  }
  return CELULAS
    .filter((c) => destinos.includes(c.id))
    .map((c) => `<span class="chip${c.id === celulaAtual ? " chip--atual" : ""}">${c.abrev}</span>`)
    .join("");
}

// Cartão de mensagem. Na seção é só consulta; no "Listar Mensagens"
// (comAcoes) ganha Editar/Excluir.
function buildMensagemCard(m, celulaAtual, comAcoes) {
  const card = document.createElement("article");
  card.className = "injecao-card mensagem-card";

  card.innerHTML = `
    ${cabecalhoHtml('<span class="pill pill--mensagem">Mensagem</span>', m.gdh, m.origem)}
    <div class="injecao-card__resumo">${escapeHtml(m.assunto || "")}</div>
    ${m.texto ? `<div class="caixa-texto${comAcoes ? " lancamento__texto" : ""}">${escapeHtml(m.texto)}</div>` : ""}
    <div class="info-grid">
      ${m.anexo ? infoHtml("Arquivo da mensagem", anexoLinkHtml(m.anexo)) : ""}
      ${infoHtml("Destinatários", `<div class="chips">${destinosChipsHtml(m, celulaAtual)}</div>`)}
    </div>
    ${rodapeHtml(m, comAcoes)}
  `;

  if (comAcoes) {
    card.querySelector(".btn-editar").addEventListener("click", () => iniciarEdicaoMensagem(m));
    const btnExcluir = card.querySelector(".btn-excluir");
    btnExcluir.addEventListener("click", () => excluirMensagem(m, btnExcluir));
  }
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

  // Situação por célula: quem já respondeu (texto e/ou arquivo da resposta)
  // e quem ainda está pendente.
  const roteadas = celulasRoteadas(it);
  const qtdRespondidas = roteadas.filter((c) => celulaRespondeu(it, c.id)).length;
  const celulasHtml = roteadas.map((c) => {
    const respostas = respostasDaCelula(it, c.id);
    const respondeu = respostas.length > 0;
    const prazo = it.celulas[c.id].texto;
    return `
      <div class="kanban-celula kanban-celula--${respondeu ? "ok" : "pendente"}">
        <div class="kanban-celula__head">
          <span class="kanban-celula__status">${respondeu ? "✓" : "⏳"}</span>
          <span class="kanban-celula__abrev mono">${c.abrev}</span>
          <span class="kanban-celula__nome">${escapeHtml(c.nome)}</span>
          <span class="kanban-celula__estado">${respondeu ? "Respondeu" : "Pendente"}</span>
        </div>
        ${prazo ? `<div class="kanban-celula__prazo"><span class="rotulo">Prazo</span>${escapeHtml(prazo)}</div>` : ""}
        ${respostas.map((r) => respostaHtml(r, false)).join("")}
      </div>
    `;
  }).join("");

  card.innerHTML = `
    ${cabecalhoHtml(urgenciaPill(it.urgencia), it.gdh, it.origem)}
    <div class="injecao-card__resumo">${escapeHtml(it.resumo || "")}</div>
    ${it.anexo ? `<div class="info-grid">${infoHtml("Arquivo do PMS", anexoLinkHtml(it.anexo))}</div>` : ""}
    <div class="bloco">
      <span class="rotulo">Situação por célula · ${qtdRespondidas}/${roteadas.length} responderam</span>
      <div class="kanban-celulas">${celulasHtml}</div>
    </div>
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

  const mensagensQuery = query(collection(db, "mensagens"), orderBy("createdAt", "desc"));
  onSnapshot(
    mensagensQuery,
    (snapshot) => {
      mensagensCache = snapshot.docs.map((d) => ({ id: d.id, ...d.data() }));
      renderMensagens();
    },
    (err) => {
      console.error("Erro no listener de mensagens:", err);
      setConn("offline");
    }
  );
}
