# Painel FTC · AZUVER

Ferramenta de apoio para um Estado-Maior Conjunto (FTC, país **VERMELHO**)
durante o exercício de simulação **AZUVER**, na ECEME (Escola de Comando e
Estado-Maior do Exército).

Aplicativo web **estático** — HTML + CSS + JS puro, sem build step, sem
npm/webpack/Node. Funciona hospedado diretamente no GitHub Pages. Os dados
são compartilhados em tempo real entre todos os usuários via Firebase
Firestore — inclusive os anexos (DOCX/PDF), que são gravados no próprio
Firestore em partes. Tudo roda no plano gratuito (Spark), sem cartão.

## Estrutura

```
index.html          layout e abas do painel
style.css            tema visual (sala de operações, escuro/claro)
app.js                lógica da aplicação (Firebase, formulários, listas)
firebase-config.js   credenciais do projeto Firebase (placeholders)
firestore.rules      regras de segurança do Firestore
```

## Passo a passo para colocar no ar

### 1. Criar/usar um projeto Firebase e preencher `firebase-config.js`

1. Acesse https://console.firebase.google.com e crie um projeto (ou use um
   existente).
2. Em **Configurações do projeto → Seus apps**, adicione um app da **Web**
   (ícone `</>`). Não é necessário Firebase Hosting.
3. Copie o objeto `firebaseConfig` gerado e cole os valores no arquivo
   `firebase-config.js` deste repositório, substituindo cada
   `"COLE_AQUI"` pelo valor correspondente:

   ```js
   export const firebaseConfig = {
     apiKey: "...",
     authDomain: "...",
     projectId: "...",
     storageBucket: "...",
     messagingSenderId: "...",
     appId: "..."
   };
   ```

### 2. Ativar o Firestore

1. No console do Firebase, vá em **Build → Firestore Database → Criar
   banco de dados**.
2. Escolha o modo de produção (as regras corretas serão coladas no passo
   seguinte) e a região mais próxima.

### 3. Colar as regras de segurança

1. Na aba **Regras** do Firestore, apague o conteúdo padrão e cole o
   conteúdo do arquivo `firestore.rules` deste repositório:

   ```
   rules_version = '2';
   service cloud.firestore {
     match /databases/{database}/documents {
       match /{document=**} {
         allow read, write: if request.auth != null;
       }
     }
   }
   ```

2. Clique em **Publicar**.

Essas regras permitem leitura e escrita para qualquer usuário
**autenticado** (inclusive anonimamente) — suficiente para uso interno do
exercício, sem exigir login visível.

### 4. Ativar autenticação anônima

1. Vá em **Build → Authentication → Sign-in method**.
2. Ative o provedor **Anonymous** (Anônimo).

O app chama `signInAnonymously()` automaticamente ao carregar a página —
não existe tela de login. O `uid` anônimo é usado apenas internamente como
identificador técnico e nunca é exibido na interface.

### 5. Publicar o repositório no GitHub

```bash
git remote add origin <URL_DO_SEU_REPOSITORIO_GITHUB>
git push -u origin main
```

### 6. Ativar o GitHub Pages

1. No repositório no GitHub, vá em **Settings → Pages**.
2. Em **Source**, selecione a branch `main` e a pasta `/ (root)`.
3. Salve. Após alguns instantes, o GitHub exibirá o link público, no
   formato:

   ```
   https://<seu-usuario>.github.io/<nome-do-repositorio>/
   ```

Esse é o link final para compartilhar com o Estado-Maior. Qualquer pessoa
que abrir o link é autenticada anonimamente de forma automática e já pode
usar o painel — não é necessário nenhum cadastro ou login manual.

## Modelo de dados (Firestore)

### Coleção `injecoes`

```
{
  gdh: string,
  origem: string,
  autor: string,
  urgencia: "critico" | "importante" | "informativo",
  resumo: string,
  celulas: {
    d1:   { sel: bool, texto: string },
    d2:   { sel: bool, texto: string },
    d3:   { sel: bool, texto: string },
    d4:   { sel: bool, texto: string },
    d5:   { sel: bool, texto: string },
    d6:   { sel: bool, texto: string },
    d7d8: { sel: bool, texto: string },
    d9:   { sel: bool, texto: string },
    d10:  { sel: bool, texto: string }
  },
  anexo: { nome, arquivoId, tamanho } | null,  // arquivo do PMS
  respostas: [                                   // respostas das células
    { nome, arquivoId, tamanho, celula, enviadoEm: number }, ...
  ],
  status: "pendente" | "analise" | "decisao" | "pronta",
  createdAt: Timestamp
}
```

A coleção mantém o nome técnico `injecoes`, mas cada documento é um **PMS**.
### Coleção `arquivos` (anexos DOCX/PDF, até 10 MB)

```
arquivos/<id>             { nome, tipo, tamanho, partes, createdAt }
arquivos/<id>/partes/<n>  { dados: Bytes }   // fatias de até 900 KB
```

Cada documento do Firestore aceita no máximo 1 MiB, por isso o arquivo é
fatiado. O PMS guarda só `{ nome, arquivoId, tamanho }`. A cota gratuita
do Firestore é de 1 GiB armazenado e 50 mil leituras/dia — um download de
um arquivo de 10 MB consome cerca de 13 leituras.

### Coleção `simulas`

```
{
  reuniao: string,
  gdh: string,
  decisao: string,
  autor: string,
  impactos: [ { celula: string, texto: string }, ... ],
  createdAt: Timestamp
}
```

### Células fixas

| id     | nome                        | abreviação |
|--------|------------------------------|------------|
| d1     | D-1 Pessoal                  | D1         |
| d2     | D-2 Inteligência              | D2         |
| d3     | D-3 Operações                | D3         |
| d4     | D-4 Logística                 | D4         |
| d5     | D-5 Planejamento              | D5         |
| d6     | D-6 C2/Ciber                   | D6         |
| d7d8   | D-7/D-8 ComSoc/OpInfo         | D7/D8      |
| d9     | D-9 Assuntos Civis             | D9         |
| d10    | D-10 Finanças                  | D10        |

## Telas

1. **Novo PMS** — registra um PMS (com arquivo DOCX/PDF opcional) e roteia
   para as células selecionadas, cada uma com seu próprio texto de
   orientação.
2. **Controle PMS** — lista em tempo real, filtrável por status e pelo
   seletor global "Ver como". Cada PMS pode receber arquivos de resposta
   das células (DOCX/PDF) e pode ser excluído (com os anexos).
3. **Súmulas de Reunião** — registra decisões de reunião com impactos por
   célula; feed cronológico reverso.
4. **Painel do Comando** — KPIs e quadro Kanban por status.

O seletor global **"Ver como"** (fixo no topo) filtra o que aparece nas
abas de Injeções e Súmulas: ao escolher uma célula específica, só aparecem
itens endereçados a ela e apenas o texto escrito especificamente para
aquela célula. As opções "Todas as células" e "ComTO / ChEM" mostram tudo,
sem filtragem. A escolha fica salva no `localStorage` do navegador.

## Observações

- Não há nenhum passo de build: basta servir os arquivos estáticos (é
  exatamente o que o GitHub Pages faz).
- Não há tela de login visível: a autenticação anônima acontece de forma
  transparente ao carregar a página.
- Para uso real de estado-maior, revise as regras de segurança do
  Firestore caso deseje restringir a leitura/escrita por outros critérios
  além de "usuário autenticado".
