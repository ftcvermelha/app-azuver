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
// Anexos (PDF e formatos do Office/LibreOffice) — guardados no próprio Firestore (plano gratuito),
// fatiados em partes de até 900 KB, pois cada documento tem limite de 1 MiB:
//   arquivos/<id>               { nome, tipo, tamanho, partes }
//   arquivos/<id>/partes/<n>    { dados: Bytes }
// PMS, mensagens e respostas guardam só as referências, em
// anexos: [{ nome, arquivoId, tamanho }] (até ANEXOS_MAX).
// ---------------------------------------------------------------------

const ANEXO_MAX_BYTES = 10 * 1024 * 1024; // 10 MB por arquivo
const ANEXO_PARTE_BYTES = 900 * 1024;
const ANEXOS_MAX = 10; // por PMS, mensagem ou resposta
// Extensão aceita → [tipo MIME, leitor usado no "Visualizar"].
const OOXML = "application/vnd.openxmlformats-officedocument";
const ODF = "application/vnd.oasis.opendocument";
const ANEXO_FORMATOS = {
  pdf: ["application/pdf", "pdf"],
  doc: ["application/msword", "doc"],
  dot: ["application/msword", "doc"],
  docx: [`${OOXML}.wordprocessingml.document`, "docx"],
  docm: ["application/vnd.ms-word.document.macroEnabled.12", "docx"],
  dotx: [`${OOXML}.wordprocessingml.template`, "docx"],
  dotm: ["application/vnd.ms-word.template.macroEnabled.12", "docx"],
  rtf: ["application/rtf", "rtf"],
  odt: [`${ODF}.text`, "odt"],
  ott: [`${ODF}.text-template`, "odt"],
  xls: ["application/vnd.ms-excel", "planilha"],
  xlt: ["application/vnd.ms-excel", "planilha"],
  xlsx: [`${OOXML}.spreadsheetml.sheet`, "planilha"],
  xlsm: ["application/vnd.ms-excel.sheet.macroEnabled.12", "planilha"],
  xlsb: ["application/vnd.ms-excel.sheet.binary.macroEnabled.12", "planilha"],
  xltx: [`${OOXML}.spreadsheetml.template`, "planilha"],
  xltm: ["application/vnd.ms-excel.template.macroEnabled.12", "planilha"],
  ods: [`${ODF}.spreadsheet`, "planilha"],
  ots: [`${ODF}.spreadsheet-template`, "planilha"],
  csv: ["text/csv", "csv"],
  ppt: ["application/vnd.ms-powerpoint", "ppt"],
  pps: ["application/vnd.ms-powerpoint", "ppt"],
  pot: ["application/vnd.ms-powerpoint", "ppt"],
  pptx: [`${OOXML}.presentationml.presentation`, "pptx"],
  pptm: ["application/vnd.ms-powerpoint.presentation.macroEnabled.12", "pptx"],
  ppsx: [`${OOXML}.presentationml.slideshow`, "pptx"],
  ppsm: ["application/vnd.ms-powerpoint.slideshow.macroEnabled.12", "pptx"],
  potx: [`${OOXML}.presentationml.template`, "pptx"],
  potm: ["application/vnd.ms-powerpoint.template.macroEnabled.12", "pptx"],
  odp: [`${ODF}.presentation`, "odp"],
  otp: [`${ODF}.presentation-template`, "odp"]
};
const ANEXO_ACCEPT = Object.keys(ANEXO_FORMATOS).map((ext) => `.${ext}`).join(",");

function formatoDe(nome) {
  return ANEXO_FORMATOS[(nome || "").split(".").pop().toLowerCase()] || null;
}

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
  if (!formatoDe(file.name)) return "Formato não aceito (use PDF ou arquivos do Office/LibreOffice).";
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
    tipo: formatoDe(file.name)[0],
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

// Arquivos já baixados ficam na memória da página (até ANEXO_CACHE_BYTES),
// para que reabrir ou baixar de novo seja instantâneo.
const ANEXO_CACHE_BYTES = 80 * 1024 * 1024;
const anexosCache = new Map(); // arquivoId → Promise<{ nome, tipo, blob }>

function carregarAnexo(arquivoId) {
  let promessa = anexosCache.get(arquivoId);
  if (promessa) {
    anexosCache.delete(arquivoId); // reinserido abaixo: fica como o mais recente
  } else {
    promessa = buscarAnexo(arquivoId);
    promessa.then(limitarCacheAnexos, () => anexosCache.delete(arquivoId));
  }
  anexosCache.set(arquivoId, promessa);
  return promessa;
}

async function limitarCacheAnexos() {
  const tamanhos = await Promise.all(
    [...anexosCache].map(([id, p]) => p.then((a) => [id, a.blob.size], () => [id, 0]))
  );
  let total = tamanhos.reduce((soma, [, t]) => soma + t, 0);
  for (const [id, t] of tamanhos) {
    if (total <= ANEXO_CACHE_BYTES) break;
    anexosCache.delete(id); // os mais antigos saem primeiro
    total -= t;
  }
}

async function buscarAnexo(arquivoId) {
  const arquivoRef = doc(db, "arquivos", arquivoId);
  // A primeira parte vem junto com os metadados (quase todo arquivo tem uma só).
  const [meta, primeira] = await Promise.all([
    getDoc(arquivoRef),
    getDoc(doc(arquivoRef, "partes", "0"))
  ]);
  if (!meta.exists()) throw new Error("Arquivo não encontrado.");
  const { nome, tipo, partes } = meta.data();

  const resto = await Promise.all(
    Array.from({ length: partes - 1 }, (_, i) => getDoc(doc(arquivoRef, "partes", String(i + 1))))
  );
  const blob = new Blob([primeira, ...resto].map((p) => p.data().dados.toUint8Array()), { type: tipo });
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
  anexosCache.delete(anexo.arquivoId);
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
// Cada formato tem seu leitor, todos guardados em vendor/ (nada vem de
// sites externos) e carregados só quando necessários:
//   PDF: visualizador do navegador no computador; nos celulares (que não
//        mostram PDF dentro da página) as páginas são desenhadas pelo PDF.js.
//   DOCX/DOCM/DOTX: docx-preview (formatação aproximada).
//   Planilhas (XLS, XLSX, ODS, CSV...): SheetJS, uma aba por planilha.
//   PPTX/PPSX: pptx-preview (aparência aproximada).
//   RTF: rtf.js.
//   ODT/ODP: convertidos aqui mesmo (texto, títulos, listas, tabelas e imagens).
//   DOC/PPT (Office 97–2003): só o texto, lido do arquivo binário.
// Passar o mouse em "Visualizar" já começa a buscar o arquivo e o leitor.
// ---------------------------------------------------------------------

const VENDOR = new URL("./vendor/", import.meta.url).href;
const PDF_LARGURA_MAX = 1600; // px do desenho de cada página (memória no celular)
const PLANILHA_LINHAS_MAX = 2000; // por planilha; o restante fica para o "Baixar"
const PLANILHA_COLUNAS_MAX = 100;
const AVISO_SIMPLIFICADO = 'Visualização simplificada. Para ver a formatação original, use "Baixar".';
const AVISO_SO_TEXTO = 'Formato antigo do Office: aqui aparece só o texto. Para ver a formatação original, use "Baixar".';

const visualizador = document.getElementById("visualizador");
const visTipo = document.getElementById("visTipo");
const visNome = document.getElementById("visNome");
const visBaixar = document.getElementById("visBaixar");
const visCorpo = document.getElementById("visCorpo");

let visAtual = null; // { token, nome, blob } do documento aberto

function visAberto(token) {
  return visAtual && visAtual.token === token;
}

// Endereços temporários de arquivos na memória; liberados ao fechar a janela
// ou ao abrir outro documento.
let visUrls = [];

function urlTemporaria(blob) {
  const url = URL.createObjectURL(blob);
  visUrls.push(url);
  return url;
}

function liberarUrls() {
  visUrls.forEach((url) => URL.revokeObjectURL(url));
  visUrls = [];
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

async function carregarGlobal(nome, arquivo) {
  if (!window[nome]) await carregarScript(VENDOR + arquivo);
  return window[nome];
}

const carregarJsZip = carregarUmaVez(() => carregarGlobal("JSZip", "jszip.min.js"));
const carregarSheetJs = carregarUmaVez(() => carregarGlobal("XLSX", "xlsx.full.min.js"));

// Leitor (ver ANEXO_FORMATOS) → biblioteca que ele usa.
const LEITORES = {
  pdf: carregarUmaVez(async () => {
    const pdfjs = await import(VENDOR + "pdf.min.mjs");
    pdfjs.GlobalWorkerOptions.workerSrc = VENDOR + "pdf.worker.min.mjs";
    return pdfjs;
  }),
  docx: carregarUmaVez(async () => {
    await carregarJsZip();
    return carregarGlobal("docx", "docx-preview.min.js");
  }),
  pptx: carregarUmaVez(() => carregarGlobal("pptxPreview", "pptx-preview.umd.js")),
  rtf: carregarUmaVez(async () => {
    // SheetJS e rtf.js trazem tabelas de caracteres com os mesmos nomes globais
    // (cptable, DO_NOT_EXPORT_CODEPAGE). O SheetJS guarda as suas ao carregar;
    // carregado antes, as globais dele são escondidas para o rtf.js usar as próprias.
    await carregarSheetJs();
    if (!window.RTFJS) window.cptable = window.DO_NOT_EXPORT_CODEPAGE = undefined;
    for (const nome of ["WMFJS", "EMFJS", "RTFJS"]) await carregarGlobal(nome, `${nome}.bundle.min.js`);
    [window.WMFJS, window.EMFJS, window.RTFJS].forEach((lib) => lib.loggingEnabled(false));
    return window.RTFJS;
  }),
  planilha: carregarSheetJs,
  csv: carregarSheetJs,
  doc: carregarSheetJs, // usa o leitor de arquivos binários (CFB) do SheetJS
  ppt: carregarSheetJs,
  odt: carregarJsZip,
  odp: carregarJsZip
};

// Registros antigos só tinham PDF e DOCX; o nome do arquivo decide o resto.
function leitorDe(nome, tipo) {
  const formato = formatoDe(nome);
  if (formato) return formato[1];
  return tipo === "application/pdf" ? "pdf" : "docx";
}

function prepararLeitor(leitor) {
  if (leitor === "pdf" && usarPdfNativo()) return Promise.resolve(null);
  return LEITORES[leitor]();
}

// Links dentro dos documentos abrem em outra aba; endereços que não sejam
// de páginas ou e-mail (ex.: "javascript:") são desativados.
function protegerLinks(raiz) {
  raiz.querySelectorAll("a[href]").forEach((a) => {
    if (a.getAttribute("href").startsWith("#")) return; // índice interno do documento
    let protocolo = "";
    try { protocolo = new URL(a.getAttribute("href"), location.href).protocol; } catch {}
    if (["http:", "https:", "mailto:"].includes(protocolo)) {
      a.target = "_blank";
      a.rel = "noopener noreferrer";
    } else {
      a.removeAttribute("href");
    }
  });
}

function avisoVisualizador(texto, erro = false) {
  visCorpo.innerHTML = `<p class="visualizador__aviso${erro ? " error" : ""}">${texto}</p>`;
}

async function visualizarAnexo(arquivoId, nome) {
  const token = {};
  liberarUrls();
  visAtual = { token, nome, blob: null };
  visTipo.textContent = extensao(nome);
  visNome.textContent = nome;
  visNome.title = nome;
  visBaixar.disabled = true;
  avisoVisualizador("Carregando documento...");
  if (!visualizador.open) visualizador.showModal();

  try {
    // Arquivo e leitor chegam ao mesmo tempo.
    const leitor = leitorDe(nome);
    const [arquivo, lib] = await Promise.all([carregarAnexo(arquivoId), prepararLeitor(leitor)]);
    if (!visAberto(token)) return;
    Object.assign(visAtual, { nome: arquivo.nome, blob: arquivo.blob });
    visBaixar.disabled = false;
    await MOSTRAR[leitor](arquivo.blob, token, lib);
    if (visAberto(token)) protegerLinks(visCorpo);
  } catch (err) {
    console.error("Erro ao visualizar anexo:", err);
    if (visAberto(token)) avisoVisualizador('Não foi possível abrir o documento aqui. Use "Baixar".', true);
  }
}

// Começa a buscar arquivo e leitor antes do clique (mouse por cima ou toque).
document.addEventListener("pointerover", (e) => {
  const ver = e.target.closest && e.target.closest("[data-visualizar]");
  if (!ver || ver.dataset.preparado) return;
  ver.dataset.preparado = "1";
  carregarAnexo(ver.dataset.visualizar).catch(() => {});
  prepararLeitor(leitorDe(ver.dataset.nome)).catch(() => {});
});

// Leitores dos formatos mais comuns já ficam prontos logo após a abertura da página.
setTimeout(() => {
  const precarregar = () => [LEITORES.docx, LEITORES.planilha, ...(usarPdfNativo() ? [] : [LEITORES.pdf])]
    .forEach((carregar) => carregar().catch(() => {}));
  if (window.requestIdleCallback) requestIdleCallback(precarregar, { timeout: 5000 });
  else precarregar();
}, 2000);

// Área com nota no alto (quando houver) e uma "folha" branca para o conteúdo.
function criarFolha(nota) {
  visCorpo.innerHTML = `
    <div class="visualizador__docs">
      ${nota ? `<p class="visualizador__nota">${nota}</p>` : ""}
      <article class="visualizador__folha"></article>
    </div>
  `;
  return visCorpo.querySelector(".visualizador__folha");
}

// ---- PDF -------------------------------------------------------------

async function mostrarPdf(blob, token, pdfjs) {
  if (!pdfjs) {
    visCorpo.innerHTML = '<iframe class="visualizador__pdf" title="Documento PDF"></iframe>';
    visCorpo.firstElementChild.src = urlTemporaria(blob);
    return;
  }

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

// ---- Word (DOCX, DOCM, DOTX) e RTF -------------------------------------

async function mostrarDocx(blob, token, docx) {
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

async function mostrarRtf(blob, token, RTFJS) {
  const documento = new RTFJS.Document(await blob.arrayBuffer(), {});
  const elementos = await documento.render();
  if (!visAberto(token)) return;
  criarFolha().append(...elementos);
}

// ---- Planilhas (XLS, XLSX, XLSM, XLSB, ODS, CSV...) --------------------

async function mostrarPlanilha(blob, token, XLSX) {
  const pasta = XLSX.read(new Uint8Array(await blob.arrayBuffer()), {
    type: "array",
    cellDates: true,
    sheetRows: PLANILHA_LINHAS_MAX
  });
  mostrarPasta(pasta, token, XLSX);
}

// CSV costuma vir em UTF-8 ou no padrão do Excel brasileiro (Windows-1252).
async function mostrarCsv(blob, token, XLSX) {
  const bytes = new Uint8Array(await blob.arrayBuffer());
  let texto;
  try {
    texto = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    texto = new TextDecoder("windows-1252").decode(bytes);
  }
  const linhas = lerCsv(texto.replace(/^\ufeff/, ""));
  const pasta = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(pasta, XLSX.utils.aoa_to_sheet(linhas.slice(0, PLANILHA_LINHAS_MAX)), "CSV");
  if (linhas.length > PLANILHA_LINHAS_MAX) {
    pasta.Sheets.CSV["!fullref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: linhas.length - 1, c: 0 } });
  }
  mostrarPasta(pasta, token, XLSX);
}

// Separador (vírgula, ponto e vírgula, tabulação ou barra) é o que mais
// aparece na primeira linha, fora de aspas.
function lerCsv(texto) {
  const primeira = texto.slice(0, texto.search(/\r?\n|$/)).replace(/"[^"]*"/g, "");
  const sep = [",", ";", "\t", "|"].reduce((melhor, c) =>
    primeira.split(c).length > primeira.split(melhor).length ? c : melhor, ",");

  const linhas = [];
  let linha = [];
  let campo = "";
  let aspas = false;
  for (let i = 0; i < texto.length; i++) {
    const c = texto[i];
    if (aspas) {
      if (c === '"' && texto[i + 1] === '"') { campo += '"'; i++; }
      else if (c === '"') aspas = false;
      else campo += c;
    } else if (c === '"') aspas = true;
    else if (c === sep) { linha.push(campo); campo = ""; }
    else if (c === "\n" || c === "\r") {
      if (c === "\r" && texto[i + 1] === "\n") i++;
      linha.push(campo);
      linhas.push(linha);
      linha = [];
      campo = "";
    } else campo += c;
  }
  if (campo || linha.length) linhas.push([...linha, campo]);
  return linhas;
}

function mostrarPasta(pasta, token, XLSX) {
  if (!visAberto(token)) return;
  const props = (pasta.Workbook && pasta.Workbook.Sheets) || [];
  const visiveis = pasta.SheetNames.filter((_, i) => !(props[i] && props[i].Hidden));
  const nomes = visiveis.length ? visiveis : pasta.SheetNames;

  visCorpo.innerHTML = `
    <div class="planilha">
      <div class="planilha__abas" role="tablist"${nomes.length > 1 ? "" : " hidden"}></div>
      <p class="visualizador__nota" hidden></p>
      <div class="planilha__tabela"></div>
    </div>
  `;
  const abas = visCorpo.querySelector(".planilha__abas");
  const nota = visCorpo.querySelector(".visualizador__nota");
  const area = visCorpo.querySelector(".planilha__tabela");

  function abrir(indice) {
    abas.querySelectorAll(".planilha__aba").forEach((b, i) => b.setAttribute("aria-selected", String(i === indice)));
    const ws = pasta.Sheets[nomes[indice]];
    area.scrollTo(0, 0);
    if (!ws || !ws["!ref"]) {
      nota.hidden = true;
      area.innerHTML = '<p class="visualizador__aviso">Planilha vazia.</p>';
      return;
    }
    // Sem as linhas e colunas vazias do fim; planilhas enormes mostram só o
    // começo, para a tela não travar.
    const faixa = XLSX.utils.decode_range(ws["!ref"]);
    const total = XLSX.utils.decode_range(ws["!fullref"] || ws["!ref"]);
    const fim = ultimaCelulaComDados(ws, XLSX);
    // Só avisa do corte se os dados chegam até a última linha lida.
    const cortaLinhas = total.e.r > faixa.e.r && fim.r >= faixa.e.r;
    faixa.e.r = Math.min(faixa.e.r, fim.r);
    faixa.e.c = Math.min(faixa.e.c, fim.c);
    if (faixa.e.r < faixa.s.r || faixa.e.c < faixa.s.c) {
      nota.hidden = true;
      area.innerHTML = '<p class="visualizador__aviso">Planilha vazia.</p>';
      return;
    }
    const cortaColunas = faixa.e.c - faixa.s.c >= PLANILHA_COLUNAS_MAX;
    if (cortaColunas) faixa.e.c = faixa.s.c + PLANILHA_COLUNAS_MAX - 1;
    nota.hidden = !(cortaColunas || cortaLinhas);
    nota.textContent = `Mostrando até ${PLANILHA_LINHAS_MAX} linhas e ${PLANILHA_COLUNAS_MAX} colunas. Para ver tudo, use "Baixar".`;
    const original = ws["!ref"];
    ws["!ref"] = XLSX.utils.encode_range(faixa);
    area.innerHTML = XLSX.utils.sheet_to_html(ws, { header: "", footer: "" });
    ws["!ref"] = original;
    protegerLinks(area);
  }

  nomes.forEach((nome, i) => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "planilha__aba";
    b.setAttribute("role", "tab");
    b.textContent = nome;
    b.addEventListener("click", () => abrir(i));
    abas.appendChild(b);
  });
  abrir(0);
}

// Última linha e coluna com conteúdo (células mescladas contam inteiras).
function ultimaCelulaComDados(ws, XLSX) {
  const fim = { r: -1, c: -1 };
  for (const ref in ws) {
    if (ref[0] === "!" || ws[ref].v === undefined || ws[ref].v === "") continue;
    const { r, c } = XLSX.utils.decode_cell(ref);
    fim.r = Math.max(fim.r, r);
    fim.c = Math.max(fim.c, c);
  }
  (ws["!merges"] || []).forEach((m) => {
    const inicio = ws[XLSX.utils.encode_cell(m.s)];
    if (!inicio || inicio.v === undefined || inicio.v === "") return;
    fim.r = Math.max(fim.r, m.e.r);
    fim.c = Math.max(fim.c, m.e.c);
  });
  return fim;
}

// ---- Apresentações PPTX/PPSX -------------------------------------------

async function mostrarPptx(blob, token, pptxPreview) {
  const dados = await blob.arrayBuffer();
  if (!visAberto(token)) return;
  visCorpo.innerHTML = '<div class="visualizador__slides"></div>';
  const area = visCorpo.firstElementChild;
  const largura = Math.min(area.clientWidth - 32, 1100);
  const previewer = pptxPreview.init(area, { width: largura, height: Math.round(largura * 9 / 16), mode: "list" });
  await previewer.preview(dados);
}

// ---- LibreOffice (ODT, ODP) -------------------------------------------
// Conversão própria do XML do arquivo: guarda texto, títulos, negrito,
// itálico, sublinhado, alinhamento, listas, tabelas e imagens.

async function lerOdf(blob, JSZip) {
  const zip = await JSZip.loadAsync(blob);
  const xml = async (caminho) => {
    const arq = zip.file(caminho);
    return arq ? new DOMParser().parseFromString(await arq.async("string"), "application/xml") : null;
  };
  const [conteudo, estilos] = await Promise.all([xml("content.xml"), xml("styles.xml")]);
  if (!conteudo) throw new Error("Arquivo ODF sem content.xml");
  return { zip, conteudo, estilo: estilosOdf([estilos, conteudo]) };
}

// Nome do estilo → CSS (negrito, itálico, sublinhado, alinhamento), seguindo os estilos-pai.
function estilosOdf(docs) {
  const estilos = new Map();
  docs.filter(Boolean).forEach((d) => {
    for (const s of d.getElementsByTagName("style:style")) {
      const texto = s.getElementsByTagName("style:text-properties")[0];
      const paragrafo = s.getElementsByTagName("style:paragraph-properties")[0];
      const props = {};
      if (texto) {
        const peso = texto.getAttribute("fo:font-weight");
        if (peso) props["font-weight"] = peso === "bold" || Number(peso) >= 600 ? "bold" : "normal";
        const estiloFonte = texto.getAttribute("fo:font-style");
        if (estiloFonte) props["font-style"] = estiloFonte === "italic" ? "italic" : "normal";
        const sublinhado = texto.getAttribute("style:text-underline-style");
        if (sublinhado) props["text-decoration"] = sublinhado === "none" ? "none" : "underline";
      }
      const alinhamento = paragrafo && paragrafo.getAttribute("fo:text-align");
      if (alinhamento) {
        props["text-align"] = { start: "left", end: "right" }[alinhamento] || alinhamento;
      }
      estilos.set(s.getAttribute("style:name"), { pai: s.getAttribute("style:parent-style-name"), props });
    }
  });

  const cache = new Map();
  return (nome) => {
    if (!nome) return "";
    if (!cache.has(nome)) {
      const props = {};
      for (let n = nome, passos = 0; n && estilos.has(n) && passos < 10; passos++) {
        const s = estilos.get(n);
        Object.entries(s.props).forEach(([k, v]) => { if (!(k in props)) props[k] = v; });
        n = s.pai;
      }
      cache.set(nome, Object.entries(props).map(([k, v]) => `${k}:${v}`).join(";"));
    }
    return cache.get(nome);
  };
}

const ODF_IGNORADOS = new Set([
  "office:annotation", "office:annotation-end", "office:forms", "text:note", "text:tracked-changes",
  "text:sequence-decls", "text:variable-decls", "text:user-field-decls", "text:table-of-content-source",
  "text:alphabetical-index-source", "text:index-title-template", "text:bookmark-start", "text:bookmark-end",
  "table:table-column", "table:table-columns", "table:table-header-columns", "table:covered-table-cell",
  "svg:title", "svg:desc", "presentation:notes", "draw:object", "draw:object-ole", "office:event-listeners"
]);
const ODF_IMAGENS = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", svg: "image/svg+xml", webp: "image/webp", bmp: "image/bmp" };

function converterOdf(origem, destino, odf) {
  for (const filho of origem.childNodes) {
    const no = converterNoOdf(filho, odf);
    if (no) destino.append(no);
  }
  return destino;
}

function converterNoOdf(no, odf) {
  if (no.nodeType === Node.TEXT_NODE) return document.createTextNode(no.nodeValue);
  if (no.nodeType !== Node.ELEMENT_NODE || ODF_IGNORADOS.has(no.nodeName)) return null;

  const criar = (tag) => {
    const el = document.createElement(tag);
    const css = odf.estilo(no.getAttribute("text:style-name") || no.getAttribute("table:style-name"));
    if (css) el.style.cssText = css;
    return converterOdf(no, el, odf);
  };

  switch (no.nodeName) {
    case "text:h": {
      const nivel = Math.min(Math.max(Number(no.getAttribute("text:outline-level")) || 1, 1), 5);
      return criar(`h${nivel + 1}`);
    }
    case "text:p": {
      const p = criar("p");
      if (!p.textContent.trim() && !p.querySelector("img")) p.append(document.createElement("br"));
      return p;
    }
    case "text:span":
    case "text:a":
      return criar("span");
    case "text:s":
      return document.createTextNode(" ".repeat(Math.min(Number(no.getAttribute("text:c")) || 1, 100)));
    case "text:tab":
      return document.createTextNode("\t");
    case "text:line-break":
      return document.createElement("br");
    case "text:list":
      return criar("ul");
    case "text:list-item":
    case "text:list-header":
      return criar("li");
    case "table:table":
      return criar("table");
    case "table:table-row":
      return criar("tr");
    case "table:table-cell": {
      const td = criar("td");
      const colunas = Number(no.getAttribute("table:number-columns-spanned")) || 1;
      const linhas = Number(no.getAttribute("table:number-rows-spanned")) || 1;
      if (colunas > 1) td.colSpan = colunas;
      if (linhas > 1) td.rowSpan = linhas;
      const repetir = Math.min(Number(no.getAttribute("table:number-columns-repeated")) || 1, 20);
      if (repetir === 1) return td;
      const grupo = document.createDocumentFragment();
      for (let i = 0; i < repetir; i++) grupo.append(td.cloneNode(true));
      return grupo;
    }
    case "draw:frame": {
      // Imagem com versões alternativas (ex.: SVG + PNG): usa só a primeira.
      const imagem = [...no.children].find((c) => c.nodeName === "draw:image");
      return imagem ? converterNoOdf(imagem, odf) : converterOdf(no, document.createDocumentFragment(), odf);
    }
    case "draw:image": {
      const caminho = (no.getAttribute("xlink:href") || "").replace(/^\.\//, "");
      const tipo = ODF_IMAGENS[caminho.split(".").pop().toLowerCase()];
      const arq = tipo && !odf.semImagens && odf.zip.file(caminho);
      if (!arq) return converterOdf(no, document.createDocumentFragment(), odf);
      const img = document.createElement("img");
      img.alt = "";
      odf.imagens.push(arq.async("uint8array").then((bytes) => { img.src = urlTemporaria(new Blob([bytes], { type: tipo })); }));
      return img;
    }
    default:
      return converterOdf(no, document.createDocumentFragment(), odf);
  }
}

async function mostrarOdt(blob, token, JSZip) {
  const odf = { ...(await lerOdf(blob, JSZip)), imagens: [] };
  if (!visAberto(token)) return;
  const corpo = odf.conteudo.getElementsByTagName("office:text")[0];
  if (!corpo) throw new Error("ODT sem office:text");
  const folha = converterOdf(corpo, document.createDocumentFragment(), odf);
  criarFolha(AVISO_SIMPLIFICADO).append(folha);
  await Promise.allSettled(odf.imagens);
}

async function mostrarOdp(blob, token, JSZip) {
  const odf = { ...(await lerOdf(blob, JSZip)), imagens: [] };
  if (!visAberto(token)) return;
  const slides = [...odf.conteudo.getElementsByTagName("draw:page")].map((pagina) =>
    converterOdf(pagina, document.createElement("div"), odf)
  );
  mostrarSlidesSimples(slides, AVISO_SIMPLIFICADO);
  await Promise.allSettled(odf.imagens);
}

// Slides em sequência, cada um num quadro "Slide N".
function mostrarSlidesSimples(conteudos, nota) {
  visCorpo.innerHTML = `
    <div class="visualizador__docs">
      <p class="visualizador__nota">${nota}</p>
      <div class="slides-simples"></div>
    </div>
  `;
  const lista = visCorpo.querySelector(".slides-simples");
  if (conteudos.length === 0) lista.innerHTML = '<p class="visualizador__aviso">Nenhum texto encontrado.</p>';
  conteudos.forEach((conteudo, i) => {
    const slide = document.createElement("section");
    slide.className = "slide-simples";
    slide.innerHTML = `<span class="slide-simples__num">Slide ${i + 1}</span>`;
    slide.append(conteudo);
    lista.appendChild(slide);
  });
}

// ---- Office 97–2003 (DOC, PPT): só o texto -------------------------------
// Esses arquivos são "pastas" binárias (CFB); o SheetJS abre a pasta e o
// texto é lido conforme a especificação pública da Microsoft ([MS-DOC], [MS-PPT]).

function fluxoCfb(XLSX, bytes, nome) {
  const pasta = XLSX.CFB.read(bytes, { type: "array" });
  const entrada = XLSX.CFB.find(pasta, nome);
  if (!entrada || !entrada.content) throw new Error(`Arquivo sem o fluxo ${nome}`);
  return Uint8Array.from(entrada.content);
}

function visao(bytes) {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

function textoDoc(XLSX, bytes) {
  const wd = fluxoCfb(XLSX, bytes, "WordDocument");
  const fib = visao(wd);
  if (fib.getUint16(0, true) !== 0xa5ec) throw new Error("DOC inválido");
  const opcoes = fib.getUint16(0x0a, true);
  if (opcoes & 0x0100) throw new Error("DOC protegido por senha");
  const tabela = fluxoCfb(XLSX, bytes, opcoes & 0x0200 ? "1Table" : "0Table");
  const t = visao(tabela);
  const caracteres = fib.getInt32(76, true); // ccpText: só o corpo do documento
  const fcClx = fib.getUint32(418, true);
  const lcbClx = fib.getUint32(422, true);

  // Tabela de pedaços: o texto fica em trechos espalhados pelo arquivo.
  let i = fcClx;
  while (i < fcClx + lcbClx && tabela[i] === 0x01) i += 3 + t.getInt16(i + 1, true);
  if (tabela[i] !== 0x02) throw new Error("DOC sem tabela de texto");
  const lcb = t.getUint32(i + 1, true);
  const plc = i + 5;
  const pedacos = (lcb - 4) / 12;
  const cp1252 = new TextDecoder("windows-1252");
  const utf16 = new TextDecoder("utf-16le");
  let texto = "";
  for (let k = 0; k < pedacos && texto.length < caracteres; k++) {
    const qtd = t.getInt32(plc + 4 * (k + 1), true) - t.getInt32(plc + 4 * k, true);
    const fc = t.getUint32(plc + 4 * (pedacos + 1) + 8 * k + 2, true);
    if (fc & 0x40000000) {
      const ini = (fc & 0x3fffffff) / 2;
      texto += cp1252.decode(wd.subarray(ini, ini + qtd));
    } else {
      texto += utf16.decode(wd.subarray(fc, fc + 2 * qtd));
    }
  }
  return limparTextoDoc(texto.slice(0, caracteres));
}

// Campos (índices, numeração...): fica só o resultado; caracteres de controle viram quebras.
function limparTextoDoc(texto) {
  let saida = "";
  const campos = []; // true = já no resultado do campo
  for (const c of texto) {
    if (c === "\x13") campos.push(false);
    else if (c === "\x14") { if (campos.length) campos[campos.length - 1] = true; }
    else if (c === "\x15") campos.pop();
    else if (!campos.length || campos[campos.length - 1]) saida += c;
  }
  return saida
    .replace(/\x07\x07/g, "\n") // fim da linha de tabela
    .replace(/\x07/g, "\t")     // fim de célula
    .replace(/[\r\x0b\x0c]/g, "\n")
    .replace(/\x1e/g, "-")
    .replace(/[\x00-\x08\x0e-\x1f]/g, "");
}

function slidesPpt(XLSX, bytes) {
  const fluxo = fluxoCfb(XLSX, bytes, "PowerPoint Document");
  const v = visao(fluxo);
  const cp1252 = new TextDecoder("windows-1252");
  const utf16 = new TextDecoder("utf-16le");

  // Percorre os registros; `modo` "lista" lê a lista de textos dos slides,
  // "slides" (reserva) lê as caixas de texto dentro de cada slide.
  function ler(modo) {
    const slides = [];
    let atual = null;
    let titulo = false;
    (function percorrer(ini, fim, naLista) {
      for (let i = ini; i + 8 <= fim;) {
        const verInst = v.getUint16(i, true);
        const tipo = v.getUint16(i + 2, true);
        const corpo = i + 8;
        const prox = corpo + v.getUint32(i + 4, true);
        if (prox > fim) break;
        if ((verInst & 0xf) === 0xf) {
          if (tipo === 0x0ff0) {
            if (modo === "lista" && verInst >> 4 === 0) percorrer(corpo, prox, true);
          } else if (tipo === 0x03ee) {
            if (modo === "slides") {
              atual = { linhas: [] };
              slides.push(atual);
              percorrer(corpo, prox, false);
              atual = null;
            }
          } else if (tipo !== 0x03f0 && tipo !== 0x03f8) {
            percorrer(corpo, prox, naLista); // anotações (0x03F0) e mestres (0x03F8) ficam de fora
          }
        } else if (tipo === 0x03f3 && naLista) {
          atual = { linhas: [] };
          slides.push(atual);
        } else if (tipo === 0x0f9f) {
          const tipoTexto = v.getUint32(corpo, true);
          titulo = tipoTexto === 0 || tipoTexto === 6;
        } else if ((tipo === 0x0fa0 || tipo === 0x0fa8) && atual) {
          const dados = fluxo.subarray(corpo, prox);
          const texto = (tipo === 0x0fa0 ? utf16 : cp1252).decode(dados).replace(/[\r\x0b]/g, "\n").trim();
          if (texto) atual.linhas.push({ titulo, texto });
        }
        i = prox;
      }
    })(0, fluxo.length, false);
    return slides;
  }

  const comTexto = (lista) => lista.some((s) => s.linhas.length);
  const lista = ler("lista");
  return comTexto(lista) ? lista : ler("slides").filter((s) => s.linhas.length);
}

async function mostrarDoc(blob, token, XLSX) {
  const texto = textoDoc(XLSX, new Uint8Array(await blob.arrayBuffer()));
  if (!visAberto(token)) return;
  const folha = criarFolha(AVISO_SO_TEXTO);
  let tabela = null;
  texto.split("\n").forEach((linha) => {
    // Linhas com células (separadas por tabulação) viram linhas de tabela.
    if (linha.includes("\t")) {
      tabela ||= folha.appendChild(document.createElement("table"));
      const tr = tabela.insertRow();
      linha.split("\t").forEach((celula) => { tr.insertCell().textContent = celula; });
      return;
    }
    tabela = null;
    const p = folha.appendChild(document.createElement("p"));
    if (linha.trim()) p.textContent = linha;
    else p.append(document.createElement("br"));
  });
}

async function mostrarPpt(blob, token, XLSX) {
  const slides = slidesPpt(XLSX, new Uint8Array(await blob.arrayBuffer()));
  if (!visAberto(token)) return;
  mostrarSlidesSimples(slides.map((slide) => {
    const div = document.createElement("div");
    slide.linhas.forEach(({ titulo, texto }) => {
      texto.split("\n").forEach((linha) => {
        const el = div.appendChild(document.createElement(titulo ? "h3" : "p"));
        el.textContent = linha;
      });
    });
    return div;
  }), AVISO_SO_TEXTO);
}

const MOSTRAR = {
  pdf: mostrarPdf,
  docx: mostrarDocx,
  rtf: mostrarRtf,
  planilha: mostrarPlanilha,
  csv: mostrarCsv,
  pptx: mostrarPptx,
  odt: mostrarOdt,
  odp: mostrarOdp,
  doc: mostrarDoc,
  ppt: mostrarPpt
};

visBaixar.addEventListener("click", () => {
  if (visAtual && visAtual.blob) salvarBlob(visAtual.nome, visAtual.blob);
});

document.getElementById("visFechar").addEventListener("click", () => visualizador.close());

// Clique fora do documento (no fundo escurecido) também fecha.
visualizador.addEventListener("click", (e) => {
  if (e.target === visualizador) visualizador.close();
});

// O evento chega depois do fechamento: se outro documento já foi aberto, ele fica.
visualizador.addEventListener("close", () => {
  if (visualizador.open) return;
  liberarUrls();
  visAtual = null;
  visCorpo.innerHTML = "";
});

// ---------------------------------------------------------------------
// Compilação das respostas de um PMS pronto num único DOCX editável:
// cabeçalho do PMS e, por célula, o prazo, o texto das respostas e o
// texto extraído de cada arquivo anexado (imagens ficam de fora).
//
// O conteúdo extraído é uma lista de blocos, montada depois em WordprocessingML:
//   { tipo: "p", runs: [{ texto, negrito, italico }] }   parágrafo
//   { tipo: "h", runs }                                   subtítulo dentro do arquivo
//   { tipo: "tabela", linhas: [["célula", ...], ...] }
//   { tipo: "aviso", texto }                              observação em itálico
// ---------------------------------------------------------------------

const paragrafo = (texto, extra = {}) => ({ tipo: "p", runs: [{ texto, ...extra }] });
const subtitulo = (texto) => ({ tipo: "h", runs: [{ texto }] });
const aviso = (texto) => ({ tipo: "aviso", texto });

function blocosDeLinhas(texto) {
  return texto.split("\n").map((linha) => paragrafo(linha));
}

// Linhas com tabulação (células de tabela no DOC) viram linhas de tabela.
function blocosDeTextoComTabelas(texto) {
  const blocos = [];
  let tabela = null;
  texto.split("\n").forEach((linha) => {
    if (linha.includes("\t")) {
      if (!tabela) blocos.push((tabela = { tipo: "tabela", linhas: [] }));
      tabela.linhas.push(linha.split("\t"));
    } else {
      tabela = null;
      blocos.push(paragrafo(linha));
    }
  });
  return blocos;
}

// ---- Word (DOCX, DOCM, DOTX): lido do XML do próprio arquivo ------------

function ligado(rPr, tag) {
  const el = rPr && [...rPr.children].find((c) => c.nodeName === tag);
  return !!el && !["0", "false", "off"].includes(el.getAttribute("w:val"));
}

function paragrafoWord(p) {
  const estilo = (p.getElementsByTagName("w:pStyle")[0] || { getAttribute: () => "" }).getAttribute("w:val") || "";
  const runs = [...p.getElementsByTagName("w:r")].map((r) => {
    const rPr = [...r.children].find((c) => c.nodeName === "w:rPr");
    let texto = "";
    for (const c of r.children) {
      if (c.nodeName === "w:t") texto += c.textContent;
      else if (c.nodeName === "w:tab") texto += "\t";
      else if (c.nodeName === "w:br" || c.nodeName === "w:cr") texto += "\n";
    }
    return { texto, negrito: ligado(rPr, "w:b"), italico: ligado(rPr, "w:i") };
  }).filter((r) => r.texto);
  const titulo = /^(heading|t[ií]tulo|title)/i.test(estilo);
  return { tipo: titulo ? "h" : "p", runs };
}

function textoParagrafosWord(el) {
  return [...el.getElementsByTagName("w:p")].map((p) => paragrafoWord(p).runs.map((r) => r.texto).join("")).join("\n");
}

function blocosDoCorpoWord(no, blocos) {
  for (const el of no.children) {
    if (el.nodeName === "w:p") blocos.push(paragrafoWord(el));
    else if (el.nodeName === "w:tbl") {
      blocos.push({
        tipo: "tabela",
        linhas: [...el.children].filter((c) => c.nodeName === "w:tr").map((tr) =>
          [...tr.children].filter((c) => c.nodeName === "w:tc").map(textoParagrafosWord)
        )
      });
    } else if (["w:sdt", "w:sdtContent", "w:customXml", "w:smartTag"].includes(el.nodeName)) {
      blocosDoCorpoWord(el, blocos);
    }
  }
  return blocos;
}

async function extrairDocx(blob, JSZip) {
  const zip = await JSZip.loadAsync(blob);
  const arq = zip.file("word/document.xml");
  if (!arq) throw new Error("DOCX sem word/document.xml");
  const xml = new DOMParser().parseFromString(await arq.async("string"), "application/xml");
  const corpo = xml.getElementsByTagName("w:body")[0];
  return corpo ? blocosDoCorpoWord(corpo, []) : [];
}

// ---- PDF: texto de cada página (PDF digitalizado não tem texto) ---------

async function extrairPdf(blob, pdfjs) {
  const pdf = await pdfjs.getDocument({ data: new Uint8Array(await blob.arrayBuffer()) }).promise;
  try {
    const blocos = [];
    let caracteres = 0;
    for (let n = 1; n <= pdf.numPages; n++) {
      const conteudo = await (await pdf.getPage(n)).getTextContent();
      let texto = "";
      conteudo.items.forEach((item) => {
        if (!("str" in item)) return;
        texto += item.str + (item.hasEOL ? "\n" : "");
      });
      caracteres += texto.trim().length;
      if (pdf.numPages > 1) blocos.push(subtitulo(`Página ${n}`));
      // Linhas em branco repetidas viram uma só.
      blocos.push(...blocosDeLinhas(texto.replace(/\n{3,}/g, "\n\n").trim()));
    }
    if (caracteres === 0) {
      return [aviso("PDF sem texto selecionável (provavelmente digitalizado). Consulte o arquivo no app.")];
    }
    return blocos;
  } finally {
    pdf.destroy();
  }
}

// ---- Planilhas e CSV: uma tabela por planilha ---------------------------

function tabelaDaPlanilha(ws, XLSX) {
  if (!ws || !ws["!ref"]) return null;
  const faixa = XLSX.utils.decode_range(ws["!ref"]);
  const fim = ultimaCelulaComDados(ws, XLSX);
  faixa.e.r = Math.min(faixa.e.r, fim.r);
  faixa.e.c = Math.min(faixa.e.c, fim.c, faixa.s.c + PLANILHA_COLUNAS_MAX - 1);
  if (faixa.e.r < faixa.s.r || faixa.e.c < faixa.s.c) return null;
  const linhas = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "", blankrows: true, range: faixa });
  return { tipo: "tabela", linhas: linhas.map((l) => l.map((v) => String(v))) };
}

async function extrairPlanilha(blob, XLSX) {
  const pasta = XLSX.read(new Uint8Array(await blob.arrayBuffer()), {
    type: "array",
    cellDates: true,
    sheetRows: PLANILHA_LINHAS_MAX
  });
  const props = (pasta.Workbook && pasta.Workbook.Sheets) || [];
  const blocos = [];
  pasta.SheetNames.forEach((nome, i) => {
    if (props[i] && props[i].Hidden) return;
    const tabela = tabelaDaPlanilha(pasta.Sheets[nome], XLSX);
    if (!tabela) return;
    if (pasta.SheetNames.length > 1) blocos.push(subtitulo(`Planilha: ${nome}`));
    blocos.push(tabela);
  });
  return blocos.length ? blocos : [aviso("Planilha vazia.")];
}

async function extrairCsv(blob) {
  const bytes = new Uint8Array(await blob.arrayBuffer());
  let texto;
  try {
    texto = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    texto = new TextDecoder("windows-1252").decode(bytes);
  }
  const linhas = lerCsv(texto.replace(/^﻿/, "")).slice(0, PLANILHA_LINHAS_MAX);
  return linhas.length ? [{ tipo: "tabela", linhas }] : [aviso("Arquivo vazio.")];
}

// ---- Apresentações PPTX/PPSX: texto de cada slide, na ordem -------------

function blocosDoSlide(no, blocos) {
  for (const el of no.children) {
    if (el.nodeName === "a:tbl") {
      blocos.push({
        tipo: "tabela",
        linhas: [...el.getElementsByTagName("a:tr")].map((tr) =>
          [...tr.getElementsByTagName("a:tc")].map((tc) =>
            [...tc.getElementsByTagName("a:p")].map((p) =>
              [...p.getElementsByTagName("a:t")].map((t) => t.textContent).join("")
            ).join("\n")
          )
        )
      });
    } else if (el.nodeName === "a:p") {
      const runs = [];
      for (const c of el.children) {
        if (c.nodeName === "a:r" || c.nodeName === "a:fld") {
          const rPr = [...c.children].find((x) => x.nodeName === "a:rPr");
          const t = c.getElementsByTagName("a:t")[0];
          if (t && t.textContent) {
            runs.push({
              texto: t.textContent,
              negrito: !!rPr && rPr.getAttribute("b") === "1",
              italico: !!rPr && rPr.getAttribute("i") === "1"
            });
          }
        } else if (c.nodeName === "a:br") runs.push({ texto: "\n" });
      }
      if (runs.some((r) => r.texto.trim())) blocos.push({ tipo: "p", runs });
    } else {
      blocosDoSlide(el, blocos);
    }
  }
  return blocos;
}

async function extrairPptx(blob, JSZip) {
  const zip = await JSZip.loadAsync(blob);
  const xml = async (caminho) => {
    const arq = zip.file(caminho);
    return arq ? new DOMParser().parseFromString(await arq.async("string"), "application/xml") : null;
  };
  const [apresentacao, rels] = await Promise.all([xml("ppt/presentation.xml"), xml("ppt/_rels/presentation.xml.rels")]);
  if (!apresentacao || !rels) throw new Error("PPTX sem presentation.xml");
  const alvos = Object.fromEntries(
    [...rels.getElementsByTagName("Relationship")].map((r) => [r.getAttribute("Id"), r.getAttribute("Target")])
  );
  const caminhos = [...apresentacao.getElementsByTagName("p:sldId")]
    .map((s) => alvos[s.getAttribute("r:id")])
    .filter(Boolean)
    .map((alvo) => (alvo.startsWith("/") ? alvo.slice(1) : `ppt/${alvo}`));
  const slides = await Promise.all(caminhos.map(xml));

  const blocos = [];
  slides.forEach((slide, i) => {
    blocos.push(subtitulo(`Slide ${i + 1}`));
    const conteudo = slide && slide.getElementsByTagName("p:cSld")[0];
    const doSlide = conteudo ? blocosDoSlide(conteudo, []) : [];
    blocos.push(...(doSlide.length ? doSlide : [aviso("Slide sem texto.")]));
  });
  return blocos;
}

// ---- ODT, ODP e RTF: convertidos para a tela e lidos de volta como blocos --

const TAGS_BLOCO = new Set(["P", "H1", "H2", "H3", "H4", "H5", "H6", "LI", "DIV", "SECTION", "UL", "OL", "TABLE"]);

function runsDoHtml(el, raiz, runs = []) {
  for (const no of el.childNodes) {
    if (no.nodeType === Node.TEXT_NODE) {
      if (!no.nodeValue) continue;
      let negrito = false;
      let italico = false;
      for (let a = no.parentElement; a && a !== raiz.parentElement; a = a.parentElement) {
        negrito ||= ["B", "STRONG"].includes(a.tagName) || a.style.fontWeight === "bold" || Number(a.style.fontWeight) >= 600;
        italico ||= ["I", "EM"].includes(a.tagName) || a.style.fontStyle === "italic";
      }
      runs.push({ texto: no.nodeValue, negrito, italico });
    } else if (no.nodeName === "BR") {
      runs.push({ texto: "\n" });
    } else if (no.nodeType === Node.ELEMENT_NODE) {
      runsDoHtml(no, raiz, runs);
    }
  }
  return runs;
}

function blocosDoHtml(el, blocos = []) {
  if (el.tagName === "TABLE") {
    blocos.push({
      tipo: "tabela",
      linhas: [...el.rows].map((tr) => [...tr.cells].map((td) => td.innerText || td.textContent.trim()))
    });
    return blocos;
  }
  const temBlocos = [...el.children].some((c) => TAGS_BLOCO.has(c.tagName));
  if (!temBlocos) {
    const runs = runsDoHtml(el, el);
    if (runs.length) blocos.push({ tipo: /^H\d$/.test(el.tagName) ? "h" : "p", runs });
    return blocos;
  }
  for (const no of el.childNodes) {
    if (no.nodeType === Node.ELEMENT_NODE && TAGS_BLOCO.has(no.tagName)) blocosDoHtml(no, blocos);
    else if (no.nodeType === Node.ELEMENT_NODE || (no.nodeValue || "").trim()) {
      const span = document.createElement("span");
      span.append(no.cloneNode(true));
      blocos.push({ tipo: "p", runs: runsDoHtml(span, span) });
    }
  }
  return blocos;
}

async function extrairOdt(blob, JSZip) {
  const odf = { ...(await lerOdf(blob, JSZip)), imagens: [], semImagens: true };
  const corpo = odf.conteudo.getElementsByTagName("office:text")[0];
  if (!corpo) throw new Error("ODT sem office:text");
  return blocosDoHtml(converterOdf(corpo, document.createElement("div"), odf));
}

async function extrairOdp(blob, JSZip) {
  const odf = { ...(await lerOdf(blob, JSZip)), imagens: [], semImagens: true };
  const blocos = [];
  [...odf.conteudo.getElementsByTagName("draw:page")].forEach((pagina, i) => {
    blocos.push(subtitulo(`Slide ${i + 1}`));
    const doSlide = blocosDoHtml(converterOdf(pagina, document.createElement("div"), odf));
    blocos.push(...(doSlide.length ? doSlide : [aviso("Slide sem texto.")]));
  });
  return blocos;
}

async function extrairRtf(blob, RTFJS) {
  const elementos = await new RTFJS.Document(await blob.arrayBuffer(), {}).render();
  const raiz = document.createElement("div");
  raiz.append(...elementos);
  return blocosDoHtml(raiz);
}

// ---- Office 97–2003 ------------------------------------------------------

async function extrairDoc(blob, XLSX) {
  return blocosDeTextoComTabelas(textoDoc(XLSX, new Uint8Array(await blob.arrayBuffer())));
}

async function extrairPpt(blob, XLSX) {
  const blocos = [];
  slidesPpt(XLSX, new Uint8Array(await blob.arrayBuffer())).forEach((slide, i) => {
    blocos.push(subtitulo(`Slide ${i + 1}`));
    slide.linhas.forEach(({ titulo, texto }) => {
      texto.split("\n").forEach((linha) => blocos.push(paragrafo(linha, { negrito: titulo })));
    });
  });
  return blocos.length ? blocos : [aviso("Apresentação sem texto.")];
}

const EXTRAIR = {
  pdf: extrairPdf,
  docx: extrairDocx,
  rtf: extrairRtf,
  planilha: extrairPlanilha,
  csv: extrairCsv,
  pptx: extrairPptx,
  odt: extrairOdt,
  odp: extrairOdp,
  doc: extrairDoc,
  ppt: extrairPpt
};

// Leitor usado na extração; o PDF precisa do PDF.js mesmo no computador.
function leitorDaExtracao(leitor) {
  if (leitor === "pdf") return LEITORES.pdf();
  if (leitor === "docx" || leitor === "pptx") return carregarJsZip();
  return LEITORES[leitor]();
}

async function extrairAnexo(anexo) {
  try {
    const leitor = leitorDe(anexo.nome);
    const [arquivo, lib] = await Promise.all([carregarAnexo(anexo.arquivoId), leitorDaExtracao(leitor)]);
    const blocos = await EXTRAIR[leitor](arquivo.blob, lib);
    return blocos.length ? blocos : [aviso("Arquivo sem texto.")];
  } catch (err) {
    console.error("Erro ao extrair texto do anexo:", anexo.nome, err);
    return [aviso("Não foi possível extrair o texto deste arquivo. Consulte-o no app.")];
  }
}

// ---- Montagem do DOCX ----------------------------------------------------

// Caracteres de controle não são aceitos no XML do Word.
function xmlTexto(texto) {
  return String(texto)
    .replace(/[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]/g, "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function runXml({ texto, negrito, italico }) {
  const props = (negrito ? "<w:b/>" : "") + (italico ? "<w:i/>" : "");
  const partes = String(texto).split(/(\t|\n)/).map((parte) => {
    if (parte === "\t") return "<w:tab/>";
    if (parte === "\n") return "<w:br/>";
    return parte ? `<w:t xml:space="preserve">${xmlTexto(parte)}</w:t>` : "";
  }).join("");
  return `<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ""}${partes}</w:r>`;
}

function paragrafoXml(runs, estilo) {
  const pPr = estilo ? `<w:pPr><w:pStyle w:val="${estilo}"/></w:pPr>` : "";
  return `<w:p>${pPr}${runs.map(runXml).join("")}</w:p>`;
}

const LARGURA_TEXTO = 9638; // twips: A4 com margens de 2 cm

function tabelaXml(linhas) {
  const colunas = Math.max(1, ...linhas.map((l) => l.length));
  const largura = Math.floor(LARGURA_TEXTO / colunas);
  const celula = (texto) =>
    `<w:tc><w:tcPr><w:tcW w:w="${largura}" w:type="dxa"/></w:tcPr>${
      String(texto).split("\n").map((linha) => paragrafoXml([{ texto: linha }], "Tabela")).join("")
    }</w:tc>`;
  return `
    <w:tbl>
      <w:tblPr><w:tblStyle w:val="GradeTabela"/><w:tblW w:w="${largura * colunas}" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>
      <w:tblGrid>${`<w:gridCol w:w="${largura}"/>`.repeat(colunas)}</w:tblGrid>
      ${linhas.map((l) => `<w:tr>${[...l, ...Array(colunas - l.length).fill("")].map(celula).join("")}</w:tr>`).join("")}
    </w:tbl>
    <w:p/>`;
}

function blocoXml(bloco) {
  if (bloco.tipo === "tabela") return bloco.linhas.length ? tabelaXml(bloco.linhas) : "";
  if (bloco.tipo === "aviso") return paragrafoXml([{ texto: bloco.texto, italico: true }], "Aviso");
  if (bloco.tipo === "h") return paragrafoXml(bloco.runs, "Titulo4");
  return paragrafoXml(bloco.runs);
}

const ESTILOS_DOCX = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="pt-BR"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="80" w:line="264" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
  <w:style w:type="paragraph" w:styleId="Titulo"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:spacing w:after="200"/></w:pPr><w:rPr><w:b/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Titulo1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:pBdr><w:bottom w:val="single" w:sz="8" w:space="4" w:color="7A1F1F"/></w:pBdr><w:spacing w:before="480" w:after="160"/><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:b/><w:color w:val="7A1F1F"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Titulo2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Titulo3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:shd w:val="clear" w:color="auto" w:fill="EEF2F4"/><w:spacing w:before="200" w:after="80"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:sz w:val="23"/><w:szCs w:val="23"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Titulo4"><w:name w:val="heading 4"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:spacing w:before="160" w:after="60"/><w:outlineLvl w:val="3"/></w:pPr><w:rPr><w:b/><w:color w:val="44546A"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Aviso"><w:name w:val="Aviso"/><w:basedOn w:val="Normal"/><w:rPr><w:i/><w:color w:val="6B7A86"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Tabela"><w:name w:val="Texto de tabela"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>
  <w:style w:type="table" w:default="1" w:styleId="TabelaNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
  <w:style w:type="table" w:styleId="GradeTabela"><w:name w:val="Table Grid"/><w:basedOn w:val="TabelaNormal"/>
    <w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="A6A6A6"/><w:left w:val="single" w:sz="4" w:color="A6A6A6"/><w:bottom w:val="single" w:sz="4" w:color="A6A6A6"/><w:right w:val="single" w:sz="4" w:color="A6A6A6"/><w:insideH w:val="single" w:sz="4" w:color="A6A6A6"/><w:insideV w:val="single" w:sz="4" w:color="A6A6A6"/></w:tblBorders></w:tblPr></w:style>
</w:styles>`;

async function gerarDocx(corpoXml, JSZip) {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);
  zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`);
  zip.file("word/styles.xml", ESTILOS_DOCX);
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${corpoXml}
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr>
  </w:body>
</w:document>`);
  return zip.generateAsync({ type: "blob", mimeType: ANEXO_FORMATOS.docx[0], compression: "DEFLATE" });
}

// ---- Documento do PMS ----------------------------------------------------

function linhaRotulo(rotulo, valor) {
  return paragrafoXml([{ texto: `${rotulo}: `, negrito: true }, { texto: valor || "—" }]);
}

async function compilarRespostasDocx(it) {
  const JSZip = await carregarJsZip();
  const partes = [
    paragrafoXml([{ texto: "Compilação de respostas — PMS" }], "Titulo"),
    linhaRotulo("GDH", it.gdh),
    linhaRotulo("Origem", it.origem),
    linhaRotulo("Autor", it.autor),
    linhaRotulo("Urgência", URGENCIA_LABEL[it.urgencia] || it.urgencia),
    linhaRotulo("Resumo", it.resumo),
    linhaRotulo("Células", celulasRoteadas(it).map((c) => c.abrev).join(", ")),
    paragrafoXml([{ texto: `Gerado pelo Painel FTC em ${formatTimestamp(Date.now())}.`, italico: true }], "Aviso")
  ];

  for (const c of celulasRoteadas(it)) {
    partes.push(paragrafoXml([{ texto: c.nome }], "Titulo1"));
    const prazo = it.celulas[c.id].texto;
    if (prazo) partes.push(linhaRotulo("Prazo", prazo));
    const respostas = respostasDaCelula(it, c.id);
    for (const [i, r] of respostas.entries()) {
      const quando = r.enviadoEm ? ` (enviada ${formatTimestamp(r.enviadoEm)})` : "";
      partes.push(paragrafoXml([{ texto: `${respostas.length > 1 ? `Resposta ${i + 1}` : "Resposta"}${quando}` }], "Titulo2"));
      if (r.texto) partes.push(...r.texto.split("\n").map((linha) => paragrafoXml([{ texto: linha }])));
      for (const anexo of anexosDaResposta(r)) {
        partes.push(paragrafoXml([{ texto: `Arquivo: ${anexo.nome}` }], "Titulo3"));
        partes.push(...(await extrairAnexo(anexo)).map(blocoXml));
      }
    }
  }
  return gerarDocx(partes.join(""), JSZip);
}

function nomeArquivoCompilado(it) {
  const gdh = (it.gdh || "sem GDH").replace(/[\\/:*?"<>|]+/g, " ").replace(/\s+/g, " ").trim().slice(0, 40);
  return `Respostas PMS ${gdh}.docx`;
}

async function baixarCompilado(it, botao) {
  const texto = botao.textContent;
  botao.disabled = true;
  botao.textContent = "Compilando...";
  try {
    salvarBlob(nomeArquivoCompilado(it), await compilarRespostasDocx(it));
  } catch (err) {
    console.error("Erro ao compilar respostas:", err);
    window.alert("Erro ao compilar as respostas. Tente novamente.");
  } finally {
    botao.disabled = false;
    botao.textContent = texto;
  }
}

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
  input.accept = ANEXO_ACCEPT;
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
            <span class="secao-tile__gdh mono" title="${escapeHtml(it.gdh || "")}">${escapeHtml(it.gdh || "—")}</span>
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
            Escolher arquivos (até 10 MB cada)
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
          <span class="secao-tile__gdh mono" title="${escapeHtml(m.gdh || "")}">${escapeHtml(m.gdh || "—")}</span>
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

// PMS cujo "ver respostas" está aberto: sobrevive às re-renderizações.
const respostasAbertas = new Set();

// Quadro de uma célula que já respondeu, com prazo e respostas.
function kanbanCelulaHtml(it, c) {
  const prazo = it.celulas[c.id].texto;
  return `
    <div class="kanban-celula kanban-celula--ok">
      <div class="kanban-celula__head">
        <span class="kanban-celula__status">✓</span>
        <span class="kanban-celula__abrev mono">${c.abrev}</span>
        <span class="kanban-celula__nome">${escapeHtml(c.nome)}</span>
        <span class="kanban-celula__estado">Respondeu</span>
      </div>
      ${prazo ? `<div class="kanban-celula__prazo"><span class="rotulo">Prazo</span>${escapeHtml(prazo)}</div>` : ""}
      ${respostasDaCelula(it, c.id).map((r) => respostaHtml(r, false)).join("")}
    </div>
  `;
}

// Cartão do quadro: primeiro, em destaque, as células que ainda faltam
// responder (com prazo); depois quem já respondeu, com as respostas
// recolhidas atrás de "ver respostas".
function buildKanbanCard(it) {
  const card = document.createElement("div");
  card.className = "kanban-card";
  card.dataset.urgencia = it.urgencia;

  const roteadas = celulasRoteadas(it);
  const faltam = roteadas.filter((c) => !celulaRespondeu(it, c.id));
  const responderam = roteadas.filter((c) => celulaRespondeu(it, c.id));
  const aberto = respostasAbertas.has(it.id);

  const faltamHtml = faltam.length === 0 ? "" : `
    <div class="faltam">
      <span class="rotulo faltam__titulo">⏳ Faltam responder · ${faltam.length} de ${roteadas.length}</span>
      <div class="faltam__lista">
        ${faltam.map((c) => {
          const prazo = it.celulas[c.id].texto;
          return `
            <span class="faltam__item" title="${escapeHtml(c.nome)}">
              <span class="faltam__abrev mono">${c.abrev}</span>
              <span class="faltam__prazo">${prazo ? `prazo ${escapeHtml(prazo)}` : "sem prazo"}</span>
            </span>
          `;
        }).join("")}
      </div>
    </div>
  `;

  const responderamHtml = responderam.length === 0 ? "" : `
    <div class="responderam">
      <div class="responderam__head">
        <span class="rotulo">✓ Responderam · ${responderam.length} de ${roteadas.length}</span>
        <span class="chips">${responderam.map((c) => `<span class="chip chip--ok">${c.abrev}</span>`).join("")}</span>
        <button type="button" class="anexo-acao responderam__toggle" aria-expanded="${aberto}">
          ${aberto ? "ocultar respostas ▴" : "ver respostas ▾"}
        </button>
      </div>
      ${statusPms(it) === "pronta" ? `
        <button type="button" class="btn btn--ghost btn--small responderam__compilar" title="Baixa um DOCX editável com o texto das respostas e dos arquivos, por célula">
          ⤓ Compilar respostas (DOCX)
        </button>
      ` : ""}
      ${aberto ? `<div class="kanban-celulas">${responderam.map((c) => kanbanCelulaHtml(it, c)).join("")}</div>` : ""}
    </div>
  `;

  card.innerHTML = `
    ${cabecalhoHtml(urgenciaPill(it.urgencia), it.gdh, it.origem)}
    <div class="injecao-card__resumo">${escapeHtml(it.resumo || "")}</div>
    ${anexosDe(it).length ? `<div class="info-grid">${infoAnexosHtml("Arquivo do PMS", "Arquivos do PMS", anexosDe(it))}</div>` : ""}
    <div class="bloco">${faltamHtml}${responderamHtml}</div>
  `;

  const compilar = card.querySelector(".responderam__compilar");
  if (compilar) compilar.addEventListener("click", () => baixarCompilado(it, compilar));

  const toggle = card.querySelector(".responderam__toggle");
  if (toggle) {
    toggle.addEventListener("click", () => {
      if (aberto) respostasAbertas.delete(it.id);
      else respostasAbertas.add(it.id);
      card.replaceWith(buildKanbanCard(it));
    });
  }

  return card;
}

// ---------------------------------------------------------------------
// PMS Versão Final — um cartão por arquivo (o nome identifica o PMS):
//   versoesFinais/<id>   { anexo: { nome, arquivoId, tamanho }, link,
//                          createdAt, atualizadoEm }
// "Carregar nova versão" troca o arquivo do cartão e apaga o anterior.
// O link opcional leva ao documento no Google Drive/OneDrive, onde a
// edição a várias mãos acontece.
// ---------------------------------------------------------------------

const finalLista = document.getElementById("finalLista");
const finalMsg = document.getElementById("finalMsg");
const finalArquivo = document.getElementById("finalArquivo");
const finalArquivoBtn = document.getElementById("finalArquivoBtn");
finalArquivo.accept = ANEXO_ACCEPT;

let versoesFinaisCache = [];
// Sobrevivem às re-renderizações do tempo real: link sendo editado
// (id → texto digitado) e aviso de cada cartão (id → { texto, erro }).
const linksEditando = new Map();
const avisosFinal = new Map();

// Só endereços http(s); devolve null se o texto não for um link válido.
function normalizarLink(texto) {
  try {
    const url = new URL(texto.trim());
    return ["http:", "https:"].includes(url.protocol) ? url.href : null;
  } catch {
    return null;
  }
}

function plataformaDoLink(link) {
  const host = new URL(link).hostname;
  if (/(^|\.)google\.com$/.test(host)) return "Google Drive";
  if (/(^|\.)(sharepoint\.com|onedrive\.live\.com|1drv\.ms|live\.com|office\.com)$/.test(host)) return "OneDrive";
  return host;
}

finalArquivo.addEventListener("change", async () => {
  const file = finalArquivo.files[0];
  finalArquivo.value = "";
  if (!file) return;
  limparMsg(finalMsg);
  const erro = validateAnexo(file);
  if (erro) return mostrarErro(finalMsg, `${file.name}: ${erro}`);

  finalArquivo.disabled = true;
  finalArquivoBtn.classList.add("is-disabled");
  let anexo = null;
  try {
    [anexo] = await enviarArquivos([file], finalMsg);
    await addDoc(collection(db, "versoesFinais"), {
      anexo,
      link: "",
      createdAt: serverTimestamp(),
      atualizadoEm: serverTimestamp()
    });
    finalMsg.textContent = "Versão final carregada.";
  } catch (err) {
    console.error("Erro ao carregar versão final:", err);
    if (anexo) await deleteAnexo(anexo).catch(() => {});
    mostrarErro(finalMsg, "Erro ao carregar. Tente novamente.");
  } finally {
    finalArquivo.disabled = false;
    finalArquivoBtn.classList.remove("is-disabled");
  }
});

async function substituirVersaoFinal(v, file) {
  const erro = validateAnexo(file);
  if (erro) {
    avisosFinal.set(v.id, { texto: `${file.name}: ${erro}`, erro: true });
    return renderVersoesFinais();
  }
  avisosFinal.set(v.id, { texto: "Enviando nova versão...", erro: false, enviando: true });
  renderVersoesFinais();
  let novo = null;
  try {
    novo = await uploadAnexo(file);
    await updateDoc(doc(db, "versoesFinais", v.id), { anexo: novo, atualizadoEm: serverTimestamp() });
    avisosFinal.set(v.id, { texto: "Nova versão carregada; a anterior foi apagada.", erro: false });
    setTimeout(() => {
      if (avisosFinal.get(v.id) && !avisosFinal.get(v.id).erro) {
        avisosFinal.delete(v.id);
        renderVersoesFinais();
      }
    }, 6000);
    // Melhor esforço: o cartão já aponta para o arquivo novo.
    await deleteAnexo(v.anexo).catch((err) => console.error("Erro ao apagar versão anterior:", err));
  } catch (err) {
    console.error("Erro ao carregar nova versão:", err);
    if (novo) await deleteAnexo(novo).catch(() => {});
    avisosFinal.set(v.id, { texto: "Erro ao carregar a nova versão. Tente novamente.", erro: true });
  }
  renderVersoesFinais();
}

async function salvarLinkFinal(v, texto) {
  const link = texto.trim() ? normalizarLink(texto) : "";
  if (link === null) {
    avisosFinal.set(v.id, { texto: "Link inválido: cole o endereço completo, começando com https://", erro: true });
    return renderVersoesFinais();
  }
  try {
    await updateDoc(doc(db, "versoesFinais", v.id), { link });
    linksEditando.delete(v.id);
    avisosFinal.delete(v.id);
  } catch (err) {
    console.error("Erro ao salvar link:", err);
    avisosFinal.set(v.id, { texto: "Erro ao salvar o link. Tente novamente.", erro: true });
  }
  renderVersoesFinais();
}

async function excluirVersaoFinal(v, btn) {
  const ok = window.confirm(`Excluir a versão final "${v.anexo.nome}"? O arquivo será apagado. Esta ação não pode ser desfeita.`);
  if (!ok) return;
  btn.disabled = true;
  try {
    await deleteDoc(doc(db, "versoesFinais", v.id));
    await deleteAnexo(v.anexo).catch((err) => console.error("Erro ao apagar arquivo:", err));
  } catch (err) {
    console.error("Erro ao excluir versão final:", err);
    btn.disabled = false;
    window.alert("Erro ao excluir. Tente novamente.");
  }
}

function linkFinalHtml(v) {
  if (linksEditando.has(v.id)) {
    return `
      <div class="final-link final-link--editando">
        <input type="url" class="final-link__input" placeholder="https://docs.google.com/... ou https://onedrive.live.com/..."
          value="${escapeHtml(linksEditando.get(v.id))}" aria-label="Link do documento no Google Drive ou OneDrive">
        <span class="final-link__acoes">
          <button type="button" class="btn btn--primary btn--small final-link__salvar">Salvar link</button>
          <button type="button" class="btn btn--ghost btn--small final-link__cancelar">Cancelar</button>
        </span>
      </div>
    `;
  }
  if (!v.link) {
    return `<button type="button" class="anexo-acao final-link__adicionar">+ Adicionar link para editar on-line (Google Drive / OneDrive)</button>`;
  }
  return `
    <div class="final-link">
      <a class="btn btn--primary btn--small" href="${escapeHtml(v.link)}" target="_blank" rel="noopener noreferrer">Editar on-line ↗</a>
      <span class="final-link__onde">no ${escapeHtml(plataformaDoLink(v.link))}</span>
      <span class="final-link__gerenciar">
        <button type="button" class="anexo-acao final-link__alterar">alterar link</button>
        <button type="button" class="anexo-acao final-link__remover">remover link</button>
      </span>
    </div>
  `;
}

function buildVersaoFinalCard(v) {
  const card = document.createElement("article");
  card.className = "injecao-card final-card";
  const aviso = avisosFinal.get(v.id);
  const atualizado = v.atualizadoEm && v.createdAt && v.atualizadoEm.toMillis && v.createdAt.toMillis &&
    v.atualizadoEm.toMillis() - v.createdAt.toMillis() > 1000;

  card.innerHTML = `
    ${anexosHtml([v.anexo])}
    ${linkFinalHtml(v)}
    ${aviso ? `<span class="form-msg${aviso.erro ? " error" : ""}">${escapeHtml(aviso.texto)}</span>` : ""}
    <div class="injecao-card__footer">
      <span class="injecao-card__autor">
        carregado ${formatTimestamp(v.createdAt)}${atualizado ? ` · nova versão ${formatTimestamp(v.atualizadoEm)}` : ""}
      </span>
      <span class="injecao-card__acoes">
        <label class="btn btn--ghost btn--small anexo-picker__btn${aviso && aviso.enviando ? " is-disabled" : ""}">
          Carregar nova versão
          <input type="file" class="anexo-picker__input final-nova" accept="${ANEXO_ACCEPT}"${aviso && aviso.enviando ? " disabled" : ""}>
        </label>
        <button type="button" class="btn btn--danger-ghost btn--small btn-excluir">Excluir</button>
      </span>
    </div>
  `;

  const on = (seletor, evento, fn) => {
    const el = card.querySelector(seletor);
    if (el) el.addEventListener(evento, fn);
  };
  on(".final-nova", "change", (e) => {
    const file = e.target.files[0];
    if (file) substituirVersaoFinal(v, file);
  });
  on(".final-link__adicionar", "click", () => {
    linksEditando.set(v.id, "");
    renderVersoesFinais();
    finalLista.querySelector(`[data-id="${v.id}"] .final-link__input`).focus();
  });
  on(".final-link__alterar", "click", () => {
    linksEditando.set(v.id, v.link);
    renderVersoesFinais();
    finalLista.querySelector(`[data-id="${v.id}"] .final-link__input`).focus();
  });
  on(".final-link__remover", "click", () => salvarLinkFinal(v, ""));
  on(".final-link__input", "input", (e) => linksEditando.set(v.id, e.target.value));
  on(".final-link__input", "keydown", (e) => {
    if (e.key === "Enter") salvarLinkFinal(v, e.target.value);
    if (e.key === "Escape") card.querySelector(".final-link__cancelar").click();
  });
  on(".final-link__salvar", "click", () => salvarLinkFinal(v, card.querySelector(".final-link__input").value));
  on(".final-link__cancelar", "click", () => {
    linksEditando.delete(v.id);
    avisosFinal.delete(v.id);
    renderVersoesFinais();
  });
  const btnExcluir = card.querySelector(".btn-excluir");
  btnExcluir.addEventListener("click", () => excluirVersaoFinal(v, btnExcluir));

  card.dataset.id = v.id;
  return card;
}

function renderVersoesFinais() {
  // O campo de link em edição mantém o foco e o cursor após o tempo real.
  const ativo = document.activeElement;
  const focoId = ativo && ativo.classList.contains("final-link__input") ? ativo.closest(".final-card").dataset.id : null;
  const cursor = focoId ? ativo.selectionStart : null;

  finalLista.innerHTML = "";
  if (versoesFinaisCache.length === 0) {
    finalLista.innerHTML = '<p class="empty-state">Nenhuma versão final carregada.</p>';
    return;
  }
  versoesFinaisCache.forEach((v) => finalLista.appendChild(buildVersaoFinalCard(v)));

  const campo = focoId && finalLista.querySelector(`[data-id="${focoId}"] .final-link__input`);
  if (campo) {
    campo.focus();
    campo.setSelectionRange(cursor, cursor);
  }
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

  const versoesFinaisQuery = query(collection(db, "versoesFinais"), orderBy("atualizadoEm", "desc"));
  onSnapshot(
    versoesFinaisQuery,
    (snapshot) => {
      versoesFinaisCache = snapshot.docs.map((d) => ({ id: d.id, ...d.data({ serverTimestamps: "estimate" }) }));
      renderVersoesFinais();
    },
    (err) => {
      console.error("Erro no listener de versões finais:", err);
      setConn("offline");
    }
  );
}
