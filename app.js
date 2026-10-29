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
  deleteField,
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
  { id: "d1", nome: "E-1 Pessoal", abrev: "E1" },
  { id: "d2", nome: "E-2 Inteligência", abrev: "E2" },
  { id: "d3", nome: "E-3 Operações", abrev: "E3" },
  { id: "d4", nome: "E-4 Logística", abrev: "E4" },
  { id: "d5", nome: "E-5 Planejamento", abrev: "E5" },
  { id: "d6", nome: "E-6 C2/Ciber", abrev: "E6" },
  { id: "d7d8", nome: "E-7/E-8 ComSoc/OpInfo", abrev: "E7/E8" },
  { id: "d9", nome: "E-9 Assuntos Civis", abrev: "E9" },
  { id: "d10", nome: "E-10 Finanças", abrev: "E10" },
  { id: "targeting", nome: "Targeting", abrev: "TGT" }
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
// PMS, mensagens e respostas guardam só as referências, em
// anexos: [{ nome, arquivoId, tamanho }] (até ANEXOS_MAX).
// ---------------------------------------------------------------------

const ANEXO_MAX_BYTES = 10 * 1024 * 1024; // 10 MB por arquivo
const ANEXO_PARTE_BYTES = 900 * 1024;
const ANEXOS_MAX = 10; // por PMS, mensagem ou resposta
const ANEXO_TIPOS = {
  pdf: "application/pdf",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
};
const ANEXO_ACCEPT = ".pdf,.docx," + Object.values(ANEXO_TIPOS).join(",");

// Anexos de um PMS ou mensagem. Os antigos tinham um só, em "anexo".
function anexosDe(item) {
  return item.anexos || (item.anexo ? [item.anexo] : []);
}

// Anexos de uma resposta. As antigas guardavam o arquivo nos próprios campos.
function anexosDaResposta(r) {
  if (r.anexos) return r.anexos;
  return r.arquivoId ? [{ nome: r.nome, arquivoId: r.arquivoId, tamanho: r.tamanho }] : [];
}

// Retorna mensagem de erro, ou null se o arquivo é aceitável.
function validateAnexo(file) {
  const ext = file.name.split(".").pop().toLowerCase();
  if (!ANEXO_TIPOS[ext]) return "Apenas arquivos DOCX ou PDF são aceitos.";
  if (file.size > ANEXO_MAX_BYTES) return "Arquivo maior que 10 MB.";
  return null;
}

// Separa os arquivos escolhidos que podem ser anexados, respeitando o
// limite de ANEXOS_MAX; `erro` explica os recusados (ou fica vazio).
function filtrarArquivos(files, jaAnexados) {
  const aceitos = [];
  const erros = [];
  let excedeu = false;
  files.forEach((f) => {
    const erro = validateAnexo(f);
    if (erro) erros.push(`${f.name}: ${erro}`);
    else if (jaAnexados + aceitos.length >= ANEXOS_MAX) excedeu = true;
    else aceitos.push(f);
  });
  if (excedeu) erros.push(`Limite de ${ANEXOS_MAX} arquivos atingido.`);
  return { aceitos, erro: erros.join(" ") };
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

// Envia os arquivos um a um, mostrando o progresso. Se algum falhar, os
// já enviados são apagados para não ficarem órfãos.
async function enviarArquivos(files, msgEl) {
  const anexos = [];
  try {
    for (let i = 0; i < files.length; i++) {
      msgEl.textContent = files.length > 1 ? `Enviando arquivo ${i + 1} de ${files.length}...` : "Enviando arquivo...";
      anexos.push(await uploadAnexo(files[i]));
    }
  } catch (err) {
    await Promise.allSettled(anexos.map(deleteAnexo));
    throw err;
  }
  return anexos;
}

async function carregarAnexo(arquivoId) {
  const arquivoRef = doc(db, "arquivos", arquivoId);
  const meta = await getDoc(arquivoRef);
  if (!meta.exists()) throw new Error("Arquivo não encontrado.");
  const { nome, tipo, partes } = meta.data();

  const snaps = await Promise.all(
    Array.from({ length: partes }, (_, i) => getDoc(doc(arquivoRef, "partes", String(i))))
  );
  const blob = new Blob(snaps.map((p) => p.data().dados.toUint8Array()), { type: tipo });
  return { nome, tipo, blob };
}

function salvarBlob(nome, blob) {
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

function extensao(nome) {
  return (nome || "").split(".").pop().toUpperCase();
}

// Arquivo anexado: [PDF] nome  12 KB          Visualizar · Baixar
function anexoHtml(anexo) {
  const id = escapeHtml(anexo.arquivoId);
  const nome = escapeHtml(anexo.nome);
  return `
    <div class="arquivo">
      <span class="arquivo__tipo">${escapeHtml(extensao(anexo.nome))}</span>
      <span class="arquivo__nome">${nome}</span>
      <span class="anexo-link__tam">${formatBytes(anexo.tamanho)}</span>
      <span class="arquivo__acoes">
        <button type="button" class="anexo-acao" data-visualizar="${id}" data-nome="${nome}">Visualizar</button>
        <button type="button" class="anexo-acao" data-baixar="${id}">Baixar</button>
      </span>
    </div>
  `;
}

function anexosHtml(anexos) {
  return `<div class="arquivos">${anexos.map(anexoHtml).join("")}</div>`;
}

// Informação "Arquivo(s) do PMS/da mensagem" no padrão dos cartões.
function infoAnexosHtml(rotuloSingular, rotuloPlural, anexos) {
  if (anexos.length === 0) return "";
  return infoHtml(anexos.length === 1 ? rotuloSingular : rotuloPlural, anexosHtml(anexos));
}

// Delegação: qualquer "Visualizar" ou "Baixar" de anexo na página.
document.addEventListener("click", async (e) => {
  const ver = e.target.closest("[data-visualizar]");
  if (ver) return visualizarAnexo(ver.dataset.visualizar, ver.dataset.nome);

  const baixar = e.target.closest("[data-baixar]");
  if (!baixar || baixar.disabled) return;
  const textoOriginal = baixar.textContent;
  baixar.disabled = true;
  baixar.textContent = "Baixando...";
  try {
    const { nome, blob } = await carregarAnexo(baixar.dataset.baixar);
    salvarBlob(nome, blob);
  } catch (err) {
    console.error("Erro ao baixar anexo:", err);
    window.alert("Erro ao baixar o arquivo. Tente novamente.");
  } finally {
    baixar.disabled = false;
    baixar.textContent = textoOriginal;
  }
});

// ---------------------------------------------------------------------
// Visualizador de anexos — janela sobre a tela, sem baixar o arquivo.
// PDF: visualizador do navegador no computador; nos celulares (que não
// mostram PDF dentro da página) as páginas são desenhadas pelo PDF.js.
// DOCX: convertido para a tela pelo docx-preview (formatação aproximada).
// As bibliotecas só são carregadas na primeira visualização.
// ---------------------------------------------------------------------

const PDFJS_URL = "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.min.mjs";
const PDFJS_WORKER_URL = "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.worker.min.mjs";
const JSZIP_URL = "https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js";
const DOCX_PREVIEW_URL = "https://cdn.jsdelivr.net/npm/docx-preview@0.3.5/dist/docx-preview.min.js";
const PDF_LARGURA_MAX = 1600; // px do desenho de cada página (memória no celular)

const visualizador = document.getElementById("visualizador");
const visTipo = document.getElementById("visTipo");
const visNome = document.getElementById("visNome");
const visBaixar = document.getElementById("visBaixar");
const visCorpo = document.getElementById("visCorpo");

let visAtual = null; // { token, nome, blob, url } do documento aberto

function visAberto(token) {
  return visAtual && visAtual.token === token;
}

function usarPdfNativo() {
  return navigator.pdfViewerEnabled && !window.matchMedia("(pointer: coarse)").matches;
}

// Carregamento único de cada biblioteca; em caso de falha, tenta de novo na próxima vez.
function carregarUmaVez(carregar) {
  let promessa = null;
  return () => {
    promessa ||= carregar().catch((err) => {
      promessa = null;
      throw err;
    });
    return promessa;
  };
}

function carregarScript(src) {
  return new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = src;
    s.onload = resolve;
    s.onerror = () => reject(new Error(`Falha ao carregar ${src}`));
    document.head.appendChild(s);
  });
}

const carregarPdfJs = carregarUmaVez(async () => {
  const pdfjs = await import(PDFJS_URL);
  pdfjs.GlobalWorkerOptions.workerSrc = PDFJS_WORKER_URL;
  return pdfjs;
});

const carregarDocxPreview = carregarUmaVez(async () => {
  if (!window.JSZip) await carregarScript(JSZIP_URL);
  if (!window.docx) await carregarScript(DOCX_PREVIEW_URL);
  return window.docx;
});

function avisoVisualizador(texto, erro = false) {
  visCorpo.innerHTML = `<p class="visualizador__aviso${erro ? " error" : ""}">${texto}</p>`;
}

async function visualizarAnexo(arquivoId, nome) {
  const token = {};
  visAtual = { token, nome, blob: null, url: null };
  visTipo.textContent = extensao(nome);
  visNome.textContent = nome;
  visNome.title = nome;
  visBaixar.disabled = true;
  avisoVisualizador("Carregando documento...");
  if (!visualizador.open) visualizador.showModal();

  try {
    const arquivo = await carregarAnexo(arquivoId);
    if (!visAberto(token)) return;
    Object.assign(visAtual, { nome: arquivo.nome, blob: arquivo.blob });
    visBaixar.disabled = false;
    if (arquivo.tipo === ANEXO_TIPOS.pdf) await mostrarPdf(arquivo.blob, token);
    else await mostrarDocx(arquivo.blob, token);
  } catch (err) {
    console.error("Erro ao visualizar anexo:", err);
    if (visAberto(token)) avisoVisualizador('Não foi possível abrir o documento aqui. Use "Baixar".', true);
  }
}

async function mostrarPdf(blob, token) {
  if (usarPdfNativo()) {
    visAtual.url = URL.createObjectURL(blob);
    visCorpo.innerHTML = '<iframe class="visualizador__pdf" title="Documento PDF"></iframe>';
    visCorpo.firstElementChild.src = visAtual.url;
    return;
  }

  const pdfjs = await carregarPdfJs();
  const pdf = await pdfjs.getDocument({ data: new Uint8Array(await blob.arrayBuffer()) }).promise;
  try {
    if (!visAberto(token)) return;
    visCorpo.innerHTML = '<div class="visualizador__paginas"></div>';
    const paginas = visCorpo.firstElementChild;
    const largura = Math.min(paginas.clientWidth * (window.devicePixelRatio || 1), PDF_LARGURA_MAX);
    for (let n = 1; n <= pdf.numPages && visAberto(token); n++) {
      const pagina = await pdf.getPage(n);
      const viewport = pagina.getViewport({ scale: largura / pagina.getViewport({ scale: 1 }).width });
      const canvas = document.createElement("canvas");
      canvas.className = "visualizador__pagina";
      canvas.width = Math.floor(viewport.width);
      canvas.height = Math.floor(viewport.height);
      paginas.appendChild(canvas);
      await pagina.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;
    }
  } finally {
    pdf.destroy();
  }
}

async function mostrarDocx(blob, token) {
  const docx = await carregarDocxPreview();
  if (!visAberto(token)) return;
  visCorpo.innerHTML = '<div class="visualizador__docx"></div>';
  // Tela estreita: o texto se ajusta à largura em vez da folha A4 inteira.
  await docx.renderAsync(blob, visCorpo.firstElementChild, null, {
    className: "docx",
    inWrapper: true,
    ignoreWidth: visCorpo.clientWidth < 820,
    ignoreHeight: true,
    breakPages: true
  });
}

visBaixar.addEventListener("click", () => {
  if (visAtual && visAtual.blob) salvarBlob(visAtual.nome, visAtual.blob);
});

document.getElementById("visFechar").addEventListener("click", () => visualizador.close());

// Clique fora do documento (no fundo escurecido) também fecha.
visualizador.addEventListener("click", (e) => {
  if (e.target === visualizador) visualizador.close();
});

visualizador.addEventListener("close", () => {
  if (visAtual && visAtual.url) URL.revokeObjectURL(visAtual.url);
  visAtual = null;
  visCorpo.innerHTML = "";
});

// ---------------------------------------------------------------------
// Seleção de vários arquivos (formulários e respostas): botão de
// escolher, contagem e lista com "remover" em cada arquivo.
// ---------------------------------------------------------------------

// itens: [{ nome, tamanho, remover }]
function renderListaArquivos(lista, infoEl, itens) {
  lista.innerHTML = "";
  lista.hidden = itens.length === 0;
  infoEl.textContent = itens.length === 0
    ? "Nenhum arquivo selecionado"
    : `${plural(itens.length, "arquivo", "arquivos")} · máximo ${ANEXOS_MAX}`;
  itens.forEach((item) => {
    const li = document.createElement("li");
    li.className = "arquivo";
    li.innerHTML = `
      <span class="arquivo__tipo">${escapeHtml(extensao(item.nome))}</span>
      <span class="arquivo__nome">${escapeHtml(item.nome)}</span>
      <span class="anexo-link__tam">${formatBytes(item.tamanho)}</span>
      <span class="arquivo__acoes"><button type="button" class="btn-remover">remover</button></span>
    `;
    li.querySelector(".btn-remover").addEventListener("click", item.remover);
    lista.appendChild(li);
  });
}

// Seletor dos formulários (PMS e Mensagem). Na edição, `carregar` mostra
// os arquivos atuais, que podem ser mantidos ou removidos um a um.
function criarAnexoPicker(prefixo, msgEl) {
  const input = document.getElementById(`${prefixo}Anexo`);
  const infoEl = document.getElementById(`${prefixo}AnexoInfo`);
  const lista = document.getElementById(`${prefixo}AnexoLista`);
  const botao = document.getElementById(`${prefixo}AnexoBtn`);
  const picker = { atuais: [], novos: [], removidos: [] };

  picker.total = () => picker.atuais.length + picker.novos.length;

  function render() {
    renderListaArquivos(lista, infoEl, [
      ...picker.atuais.map((a) => ({
        nome: a.nome,
        tamanho: a.tamanho,
        remover: () => {
          picker.atuais = picker.atuais.filter((x) => x !== a);
          picker.removidos.push(a);
          render();
        }
      })),
      ...picker.novos.map((f) => ({
        nome: f.name,
        tamanho: f.size,
        remover: () => {
          picker.novos = picker.novos.filter((x) => x !== f);
          render();
        }
      }))
    ]);
    const cheio = picker.total() >= ANEXOS_MAX;
    input.disabled = cheio;
    botao.classList.toggle("is-disabled", cheio);
  }

  picker.carregar = (anexos) => {
    picker.atuais = [...anexos];
    picker.novos = [];
    picker.removidos = [];
    render();
  };

  input.addEventListener("change", () => {
    limparMsg(msgEl);
    const { aceitos, erro } = filtrarArquivos([...input.files], picker.total());
    input.value = "";
    picker.novos.push(...aceitos);
    render();
    if (erro) mostrarErro(msgEl, erro);
  });

  return picker;
}

// Anexos finais ao salvar: os atuais mantidos + os novos enviados. Devolve
// também os que foram removidos, para serem apagados depois de salvar.
async function resolverAnexos(picker, msgEl) {
  const novos = await enviarArquivos(picker.novos, msgEl);
  return { anexos: [...picker.atuais, ...novos], descartados: [...picker.removidos] };
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
  pickerPms.carregar([]);
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

  pickerPms.carregar(anexosDe(it));
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
    const { anexos } = await resolverAnexos(pickerPms, novaInjecaoMsg);

    await addDoc(collection(db, "injecoes"), {
      gdh,
      origem,
      autor,
      urgencia,
      resumo,
      celulas,
      anexos,
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
    const { anexos, descartados } = await resolverAnexos(pickerPms, novaInjecaoMsg);

    // "anexo" (um só arquivo) é o formato antigo, substituído por "anexos".
    const alteracoes = { ...dados, anexos, anexo: deleteField(), editadoEm: serverTimestamp() };
    if (orfas.length > 0) alteracoes.respostas = arrayRemove(...orfas);
    await updateDoc(doc(db, "injecoes", edicao.id), alteracoes);

    // Melhor esforço: apaga os arquivos removidos e os das respostas descartadas.
    const lixo = [...descartados, ...orfas.flatMap(anexosDaResposta)];
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
    const arquivos = [...anexosDe(it), ...(it.respostas || []).flatMap(anexosDaResposta)];
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
  pickerMsg.carregar([]);
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

  pickerMsg.carregar(anexosDe(m));
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
  if (!texto && pickerMsg.total() === 0) {
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
    const { anexos, descartados } = await resolverAnexos(pickerMsg, mensagemMsg);
    if (edicaoMsg) {
      await updateDoc(doc(db, "mensagens", edicaoMsg.id), { ...dados, anexos, anexo: deleteField(), editadoEm: serverTimestamp() });
      // Melhor esforço: apaga os arquivos removidos.
      await Promise.allSettled(descartados.map(deleteAnexo));
      limparFormMensagem();
      mensagemMsg.textContent = "Mensagem atualizada.";
    } else {
      await addDoc(collection(db, "mensagens"), { ...dados, anexos, createdAt: serverTimestamp() });
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
    `Excluir a mensagem ${m.gdh || ""}${m.assunto ? ` (${m.assunto})` : ""}? Os arquivos anexados também serão apagados. Esta ação não pode ser desfeita.`
  );
  if (!ok) return;
  btn.disabled = true;
  try {
    await deleteDoc(doc(db, "mensagens", m.id));
    // Melhor esforço: falha ao apagar um arquivo não impede a exclusão da mensagem.
    await Promise.allSettled(anexosDe(m).map(deleteAnexo));
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

// Resposta de uma célula: texto (para quem, assunto...) e/ou arquivos.
function respostaHtml(r, comRemover) {
  const anexos = anexosDaResposta(r);
  return `
    <div class="resposta">
      ${r.texto ? `<div class="resposta__texto">${escapeHtml(r.texto)}</div>` : ""}
      ${anexos.length ? anexosHtml(anexos) : ""}
      <div class="resposta__rodape">
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

// Rascunhos das respostas (texto e arquivos escolhidos, ainda não
// enviados): sobrevivem às re-renderizações em tempo real.
const rascunhosResposta = new Map(); // "<pmsId>:<celulaId>" -> { texto, arquivos }
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
    ${anexosDe(it).length ? `<div class="info-grid">${infoAnexosHtml("Arquivo do PMS", "Arquivos do PMS", anexosDe(it))}</div>` : ""}
    ${rodapeHtml(it, true)}
  `;

  row.querySelector(".btn-editar").addEventListener("click", () => iniciarEdicao(it));
  const btnExcluir = row.querySelector(".btn-excluir");
  btnExcluir.addEventListener("click", () => excluirPms(it, btnExcluir));
  return row;
}

// Cartão de um PMS dentro da seção de uma célula: mostra o prazo, as
// respostas e o envio de resposta dessa célula (texto e/ou arquivos, que
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
    infoAnexosHtml("Arquivo do PMS", "Arquivos do PMS", anexosDe(it));

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
        <div class="anexo-picker">
          <label class="btn btn--ghost btn--small anexo-picker__btn">
            Escolher arquivos (DOCX/PDF)
            <input type="file" class="anexo-picker__input resposta-file" accept="${ANEXO_ACCEPT}" multiple>
          </label>
          <span class="anexo-picker__nome resposta-arquivo"></span>
        </div>
        <ul class="arquivos arquivos--edicao resposta-lista" hidden></ul>
        <div class="resposta-form__acoes">
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
      const anexos = anexosDaResposta(resposta);
      const qual = !resposta.texto && anexos.length === 1 ? `a resposta "${anexos[0].nome}"` : "esta resposta";
      const ok = window.confirm(`Remover ${qual} da ${abrev}? O PMS volta a ficar pendente para esta célula se não houver outra resposta dela.`);
      if (!ok) return;
      btn.disabled = true;
      try {
        // arrayRemove exige o objeto idêntico ao gravado.
        await updateDoc(doc(db, "injecoes", it.id), { respostas: arrayRemove(resposta) });
        // Melhor esforço: apaga os arquivos da resposta.
        await Promise.allSettled(anexos.map(deleteAnexo));
      } catch (err) {
        console.error("Erro ao remover resposta:", err);
        btn.disabled = false;
        window.alert("Erro ao remover. Tente novamente.");
      }
    });
  });

  const rascunho = rascunhosResposta.get(chave) || { texto: "", arquivos: [] };
  const respostaForm = card.querySelector(".resposta-form");
  const respostaTexto = card.querySelector(".resposta-texto");
  const respostaFile = card.querySelector(".resposta-file");
  const respostaLista = card.querySelector(".resposta-lista");
  const arquivosInfo = card.querySelector(".resposta-arquivo");
  const respostaMsg = card.querySelector(".resposta-msg");

  function mostrarArquivos() {
    renderListaArquivos(respostaLista, arquivosInfo, rascunho.arquivos.map((f) => ({
      nome: f.name,
      tamanho: f.size,
      remover: () => {
        rascunho.arquivos = rascunho.arquivos.filter((x) => x !== f);
        rascunhosResposta.set(chave, rascunho);
        mostrarArquivos();
      }
    })));
    const cheio = rascunho.arquivos.length >= ANEXOS_MAX;
    respostaFile.disabled = cheio;
    respostaFile.closest("label").classList.toggle("is-disabled", cheio);
  }

  // Durante o envio, nada do formulário pode ser alterado.
  function travarForm() {
    respostaForm.querySelectorAll("textarea, input, button").forEach((el) => (el.disabled = true));
    respostaFile.closest("label").classList.add("is-disabled");
  }

  respostaTexto.value = rascunho.texto;
  mostrarArquivos();
  if (respostasEnviando.has(chave)) {
    travarForm();
    respostaMsg.textContent = "Enviando resposta...";
  }

  respostaTexto.addEventListener("input", () => {
    rascunho.texto = respostaTexto.value;
    rascunhosResposta.set(chave, rascunho);
  });

  respostaFile.addEventListener("change", () => {
    limparMsg(respostaMsg);
    const { aceitos, erro } = filtrarArquivos([...respostaFile.files], rascunho.arquivos.length);
    respostaFile.value = "";
    rascunho.arquivos = [...rascunho.arquivos, ...aceitos];
    rascunhosResposta.set(chave, rascunho);
    mostrarArquivos();
    if (erro) mostrarErro(respostaMsg, erro);
  });

  card.querySelector(".resposta-enviar").addEventListener("click", async () => {
    limparMsg(respostaMsg);
    const texto = respostaTexto.value.trim();
    const arquivos = rascunho.arquivos;
    if (!texto && arquivos.length === 0) {
      mostrarErro(respostaMsg, "Escreva a resposta ou escolha um arquivo.");
      respostaTexto.focus();
      return;
    }

    respostasEnviando.add(chave);
    travarForm();
    respostaMsg.textContent = "Enviando resposta...";
    let aviso = null;
    try {
      const anexos = await enviarArquivos(arquivos, respostaMsg);
      rascunhosResposta.delete(chave);
      await updateDoc(doc(db, "injecoes", it.id), {
        respostas: arrayUnion({ id: crypto.randomUUID(), celula: celulaId, texto, anexos, enviadoEm: Date.now() })
      });
      aviso = { texto: "Resposta enviada.", erro: false };
    } catch (err) {
      console.error("Erro ao enviar resposta:", err);
      rascunhosResposta.set(chave, { texto, arquivos });
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
      ${infoAnexosHtml("Arquivo da mensagem", "Arquivos da mensagem", anexosDe(m))}
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
    ${anexosDe(it).length ? `<div class="info-grid">${infoAnexosHtml("Arquivo do PMS", "Arquivos do PMS", anexosDe(it))}</div>` : ""}
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
