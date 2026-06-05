const express = require("express");
const pptxgen = require("pptxgenjs");
const cors = require("cors");

const app = express();
app.use(cors());
app.use(express.json({ limit: "10mb" }));

const PORT = process.env.PORT || 3000;

const FONTE = "Arial";
const COR_TEXTO = "000000";
const MAX_LINHAS = 6;
const FONT_MIN = 36;

function calcFonte(linhas) {
  const n = linhas.length;
  const maxChars = Math.max(...linhas.map(l => l.length));
  let fs = n <= 2 ? 54 : n <= 3 ? 48 : n <= 4 ? 44 : n <= 5 ? 40 : 38;
  if (maxChars > 30) fs = Math.min(fs, 44);
  if (maxChars > 38) fs = Math.min(fs, 40);
  if (maxChars > 48) fs = Math.min(fs, 38);
  if (maxChars > 56) fs = Math.min(fs, 36);
  return Math.max(fs, FONT_MIN);
}

function addSlideTexto(pres, texto, negrito, corFundo) {
  const todasLinhas = texto.split("\n").filter(l => l.trim() !== "" || texto.split("\n").length <= 2);
  for (let i = 0; i < todasLinhas.length; i += MAX_LINHAS) {
    const bloco = todasLinhas.slice(i, i + MAX_LINHAS);
    const slide = pres.addSlide();
    slide.background = { color: corFundo || "FFFFFF" };
    if (corFundo === "000000") return slide;
    const fs = calcFonte(bloco);
    const arr = bloco.map((l, idx) => ({
      text: l,
      options: { breakLine: idx < bloco.length - 1, bold: negrito }
    }));
    slide.addText(arr, {
      x: 0.3, y: 0.2, w: 9.4, h: 7.1,
      fontSize: fs, color: COR_TEXTO,
      fontFace: FONTE, align: "center",
      valign: "middle", lineSpacingMultiple: 1.2
    });
  }
}

function addSlidePreto(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "000000" };
  return slide;
}

function addSlideTitulo(pres, titulo, autor) {
  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };
  slide.addText(titulo, {
    x: 0.3, y: 1.5, w: 9.4, h: 2.5,
    fontSize: 40, bold: true, color: COR_TEXTO,
    fontFace: FONTE, align: "center", valign: "middle"
  });
  if (autor) {
    slide.addText(autor, {
      x: 0.3, y: 4.2, w: 9.4, h: 1.0,
      fontSize: 24, italic: true, color: "666666",
      fontFace: FONTE, align: "center"
    });
  }
}

function addSlideFixo(pres, titulo, linhas) {
  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };
  // Título no topo
  slide.addText(titulo, {
    x: 0.3, y: 0.2, w: 9.4, h: 0.7,
    fontSize: 22, bold: true, color: COR_TEXTO,
    fontFace: FONTE, align: "left", valign: "top"
  });
  // Linha separadora
  slide.addShape("rect", {
    x: 0.3, y: 0.95, w: 9.4, h: 0.03,
    fill: { color: "CCCCCC" }, line: { color: "CCCCCC" }
  });
  // Conteúdo centralizado no meio
  const n = linhas.length;
  const fs = n <= 1 ? 48 : n <= 2 ? 44 : n <= 3 ? 40 : n <= 4 ? 36 : 32;
  const arr = linhas.map((l, i) => ({
    text: l,
    options: { breakLine: i < linhas.length - 1, bold: false }
  }));
  slide.addText(arr, {
    x: 0.3, y: 1.1, w: 9.4, h: 6.3,
    fontSize: fs, color: COR_TEXTO,
    fontFace: FONTE, align: "center",
    valign: "middle", lineSpacingMultiple: 1.3
  });
}

function addSlideCapa(pres, domingo, subtitulo) {
  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };
  slide.addText(domingo, {
    x: 1.0, y: 1.8, w: 8, h: 0.5,
    fontSize: 18, color: "888888",
    fontFace: FONTE, align: "center"
  });
  slide.addText(subtitulo, {
    x: 1.0, y: 2.4, w: 8, h: 2.5,
    fontSize: 52, bold: true, color: COR_TEXTO,
    fontFace: FONTE, align: "center", valign: "middle"
  });
}

function renderMusica(pres, musica) {
  musica.secoes.forEach(function(secao) {
    const isRefrao = secao.tipo === "Refrão" || secao.tipo === "Única" || secao.tipo === "Aclamação";
    addSlideTexto(pres, secao.texto, isRefrao);
  });
}

// =============================================
// BANCO DE MÚSICAS
// =============================================

const banco = {
  cantos_entrada: [
    {
      titulo: "O Espírito do Senhor",
      autor: "Reginaldo Veloso",
      secoes: [
        { tipo: "Refrão", texto: "O Espírito do Senhor\no universo todo encheu,\ntudo abarca em seu saber.\nTudo enlaça em seu amor.\nAleluia, aleluia! Aleluia, aleluia! (Bis)" },
        { tipo: "Estrofe", texto: "1. Levanta-se Deus, cadê os inimigos?\nNa sua presença perecem os iníquos!\nSão como fumaça que desaparece,\nsão cera no fogo, que logo derrete!" },
        { tipo: "Refrão", texto: "O Espírito do Senhor\no universo todo encheu,\ntudo abarca em seu saber.\nTudo enlaça em seu amor.\nAleluia, aleluia! Aleluia, aleluia! (Bis)" },
        { tipo: "Estrofe", texto: "2. Os justos se alegram diante de Deus,\ncantai ao Senhor, vibrai, filhos seus!\nAbri o caminho ao grão cavaleiro,\ndançai diante dele, Senhor justiceiro." },
        { tipo: "Refrão", texto: "O Espírito do Senhor\no universo todo encheu,\ntudo abarca em seu saber.\nTudo enlaça em seu amor.\nAleluia, aleluia! Aleluia, aleluia! (Bis)" },
        { tipo: "Estrofe", texto: "3. Dos órfãos é Pai, das viúvas juiz,\nem sua morada só Ele é quem diz:\nquem estava sozinho, família encontrou,\nquem estava oprimido, tua mão libertou!" },
        { tipo: "Refrão", texto: "O Espírito do Senhor\no universo todo encheu,\ntudo abarca em seu saber.\nTudo enlaça em seu amor.\nAleluia, aleluia! Aleluia, aleluia! (Bis)" },
        { tipo: "Estrofe", texto: "4. À frente do povo saíste, ó Deus,\nos céus gotejaram, a terra tremeu;\nna sua presença se abala o Sinai,\né Deus que avança, que avança e vai!" },
        { tipo: "Refrão", texto: "O Espírito do Senhor\no universo todo encheu,\ntudo abarca em seu saber.\nTudo enlaça em seu amor.\nAleluia, aleluia! Aleluia, aleluia! (Bis)" },
        { tipo: "Estrofe", texto: "5. Uma chuva abundante do céu derramaste\ne a tua herança exausta saciaste;\nfizeste em tua paz viver teu rebanho\ne os necessitados tiveram seu ganho." },
        { tipo: "Refrão", texto: "O Espírito do Senhor\no universo todo encheu,\ntudo abarca em seu saber.\nTudo enlaça em seu amor.\nAleluia, aleluia! Aleluia, aleluia! (Bis)" },
        { tipo: "Estrofe", texto: "6. Falou sua palavra, saem os portadores,\ndebandam os reis e fartam-se os pobres!\nImenso é o poder de nosso Senhor,\nsubindo às alturas, cativos levou." },
        { tipo: "Refrão", texto: "O Espírito do Senhor\no universo todo encheu,\ntudo abarca em seu saber.\nTudo enlaça em seu amor.\nAleluia, aleluia! Aleluia, aleluia! (Bis)" },
        { tipo: "Estrofe", texto: "7. Bendito tu sejas, Senhor, todo dia,\nTu és quem nos salva, quem nos alivia;\nÉs tu nosso Deus, o libertador!\nQuem livra da morte, só mesmo o Senhor!" },
        { tipo: "Refrão", texto: "O Espírito do Senhor\no universo todo encheu,\ntudo abarca em seu saber.\nTudo enlaça em seu amor.\nAleluia, aleluia! Aleluia, aleluia! (Bis)" },
      ]
    },
    {
      titulo: "Estaremos Aqui Reunidos",
      autor: "Pe. Zezinho",
      secoes: [
        { tipo: "Refrão", texto: "Estaremos aqui reunidos,\ncomo estavam em Jerusalém,\npois só quando vivemos unidos,\né que o Espírito Santo nos vem." },
        { tipo: "Estrofe", texto: "1. Ninguém para esse vento passando,\nninguém vê, e ele sopra onde quer.\nForça igual têm o Espírito quando\nfaz a Igreja de Cristo crescer." },
        { tipo: "Refrão", texto: "Estaremos aqui reunidos,\ncomo estavam em Jerusalém,\npois só quando vivemos unidos,\né que o Espírito Santo nos vem." },
        { tipo: "Estrofe", texto: "2. Feita de homens a Igreja é divina,\npois o Espírito Santo a conduz,\ncomo um fogo que aquece e ilumina,\nque é pureza, que é vida, que é luz." },
        { tipo: "Refrão", texto: "Estaremos aqui reunidos,\ncomo estavam em Jerusalém,\npois só quando vivemos unidos,\né que o Espírito Santo nos vem." },
        { tipo: "Estrofe", texto: "3. Sua imagem são línguas ardentes,\npois o Amor é comunicação.\nE é preciso que todas as gentes\nsaibam quanto felizes serão." },
        { tipo: "Refrão", texto: "Estaremos aqui reunidos,\ncomo estavam em Jerusalém,\npois só quando vivemos unidos,\né que o Espírito Santo nos vem." },
        { tipo: "Estrofe", texto: "4. Quando o Espírito espalma suas graças,\nfaz dos povos um só coração.\nCresce a Igreja onde todas as raças\num só Deus, um só Pai louvarão." },
        { tipo: "Refrão", texto: "Estaremos aqui reunidos,\ncomo estavam em Jerusalém,\npois só quando vivemos unidos,\né que o Espírito Santo nos vem." },
      ]
    },
    {
      titulo: "Vinde, Espírito de Deus",
      autor: "Frei Fabretti",
      secoes: [
        { tipo: "Refrão", texto: "E cantaremos aleluia!\nE a nossa terra renovada ficará,\nse o vosso Espírito, Senhor, nos enviais. (Bis)" },
        { tipo: "Estrofe", texto: "1. Vinde Espírito de Deus,\ne enchei os corações dos fiéis com vossos dons!\nAcendei neles o amor\ncom um fogo abrasador,\nvos pedimos, ó Senhor!" },
        { tipo: "Refrão", texto: "E cantaremos aleluia!\nE a nossa terra renovada ficará,\nse o vosso Espírito, Senhor, nos enviais. (Bis)" },
        { tipo: "Estrofe", texto: "2. Vós que unistes tantas gentes,\ntantas línguas diferentes\nnuma fé na unidade;\npra buscar sempre a verdade\ne servir o vosso Reino\ncom a mesma caridade." },
        { tipo: "Refrão", texto: "E cantaremos aleluia!\nE a nossa terra renovada ficará,\nse o vosso Espírito, Senhor, nos enviais. (Bis)" },
      ]
    },
    {
      titulo: "O Espírito do Senhor Encheu o Universo",
      autor: "Frei Wanderson",
      secoes: [
        { tipo: "Refrão", texto: "O Espírito do Senhor encheu o universo;\nele mantém unidas todas as coisas\ne conhece todas as línguas. Aleluia." },
        { tipo: "Estrofe", texto: "1. Eis que Deus se põe de pé\ne os inimigos se dispersam!\nFogem longe de sua face\nos que odeiam o Senhor." },
        { tipo: "Refrão", texto: "O Espírito do Senhor encheu o universo;\nele mantém unidas todas as coisas\ne conhece todas as línguas. Aleluia." },
        { tipo: "Estrofe", texto: "2. Mas os justos se alegram\nna presença do Senhor,\nrejubilam satisfeitos\ne exultam de alegria." },
        { tipo: "Refrão", texto: "O Espírito do Senhor encheu o universo;\nele mantém unidas todas as coisas\ne conhece todas as línguas. Aleluia." },
        { tipo: "Estrofe", texto: "3. Derramastes lá do alto\numa chuva generosa,\ne vossa terra, vossa herança,\njá cansada renovastes." },
        { tipo: "Refrão", texto: "O Espírito do Senhor encheu o universo;\nele mantém unidas todas as coisas\ne conhece todas as línguas. Aleluia." },
      ]
    },
    {
      titulo: "Enviai o Vosso Espírito, Senhor",
      autor: "Eliana Ribeiro",
      secoes: [
        { tipo: "Refrão", texto: "A Deus clamaremos:\nEnviai Teu Santo Espírito, Senhor,\ne renovai a face da Terra. (Bis)" },
        { tipo: "Estrofe", texto: "1. Vem, Espírito de amor,\nreacender a chama que se apagou,\nreinflamar os corações com línguas de fogo." },
        { tipo: "Refrão", texto: "A Deus clamaremos:\nEnviai Teu Santo Espírito, Senhor,\ne renovai a face da Terra. (Bis)" },
        { tipo: "Estrofe", texto: "2. Vem, Espírito de amor,\nrecriar o que o mundo destruiu,\nreavivar em nós o primeiro amor,\ncomo em Pentecostes, vem viver em nós,\nnos unindo à criação." },
        { tipo: "Refrão", texto: "A Deus clamaremos:\nEnviai Teu Santo Espírito, Senhor,\ne renovai a face da Terra. (Bis)" },
      ]
    },
  ],

  aspersao: [
    {
      titulo: "Banhados em Cristo",
      autor: "",
      secoes: [
        { tipo: "Refrão", texto: "Banhados em Cristo,\nsomos uma nova criatura.\nAs coisas antigas já se passaram,\nsomos nascidos de novo.\nAleluia, aleluia, aleluia!" },
        { tipo: "Estrofe", texto: "1. Lavados na fonte viva\ndo lado aberto de Cristo,\ntranspomos vitoriosos\nas portas do paraíso!" },
        { tipo: "Refrão", texto: "Banhados em Cristo,\nsomos uma nova criatura.\nAs coisas antigas já se passaram,\nsomos nascidos de novo.\nAleluia, aleluia, aleluia!" },
      ]
    },
    {
      titulo: "Derramarei Sobre Vós",
      autor: "",
      secoes: [
        { tipo: "Refrão", texto: "Derramarei sobre vós uma água pura,\nsereis purificados de todas as faltas,\neu vos darei um coração novo, diz o Senhor." },
        { tipo: "Estrofe", texto: "1. Tende piedade de mim, ó Deus,\nsegundo a vossa grande misericórdia,\npor vossa bondade imensa,\napagai minha iniquidade." },
        { tipo: "Refrão", texto: "Derramarei sobre vós uma água pura,\nsereis purificados de todas as faltas,\neu vos darei um coração novo, diz o Senhor." },
        { tipo: "Estrofe", texto: "2. Lavai-me todo inteiro de minha culpa\ne do meu pecado purificai-me,\nsim, reconheço minha maldade,\nà minha frente está sempre o meu pecado." },
        { tipo: "Refrão", texto: "Derramarei sobre vós uma água pura,\nsereis purificados de todas as faltas,\neu vos darei um coração novo, diz o Senhor." },
        { tipo: "Estrofe", texto: "3. Dos meus pecados desviai a vossa face\ne todas as minhas culpas apagai,\ncriai em mim um coração que seja puro, meu Deus,\nponde em mim um espírito resoluto." },
        { tipo: "Refrão", texto: "Derramarei sobre vós uma água pura,\nsereis purificados de todas as faltas,\neu vos darei um coração novo, diz o Senhor." },
      ]
    },
  ],

  vela: [
    {
      titulo: "Ó Luz do Senhor",
      autor: "",
      secoes: [
        { tipo: "Única", texto: "Ó luz do Senhor\nque vem sobre a terra,\ninunda meu ser,\npermanece em nós." },
      ]
    },
    {
      titulo: "A Luz Virá",
      autor: "",
      secoes: [
        { tipo: "Única", texto: "A luz virá, a luz virá,\ne resplandecerá um novo dia. (Bis)" },
      ]
    },
  ],

  penitencial: [
    {
      titulo: "Senhor, que Subindo ao Céu",
      autor: "",
      secoes: [
        { tipo: "Estrofe", texto: "Senhor, que subindo ao céu\nnos presenteastes com o dom do Espírito,\ntende piedade de nós.\nKyrie eleison! Kyrie eleison! (Bis)" },
        { tipo: "Estrofe", texto: "Cristo, que dais vida a todas as coisas\ncom o poder da vossa palavra,\ntende piedade de nós.\nChriste eleison! Christe eleison! (Bis)" },
        { tipo: "Estrofe", texto: "Senhor, Rei do universo e Senhor dos séculos,\ntende piedade de nós.\nKyrie eleison! Kyrie eleison! (Bis)" },
      ]
    },
    {
      titulo: "Sacerdote da Nova Aliança",
      autor: "",
      secoes: [
        { tipo: "Estrofe", texto: "Senhor, que sois o eterno sacerdote\nda nova Aliança,\ntende piedade de nós.\nSenhor, tende piedade de nós. (Bis)" },
        { tipo: "Estrofe", texto: "Cristo, que nos edificais como pedras vivas\nno templo santo de Deus,\ntende piedade de nós.\nCristo, tende piedade de nós. (Bis)" },
        { tipo: "Estrofe", texto: "Senhor, que nos tornais concidadãos dos santos\nno reino dos céus,\ntende piedade de nós.\nSenhor, tende piedade de nós. (Bis)" },
      ]
    },
    {
      titulo: "Senhor, Senhor, Tende Piedade de Nós",
      autor: "",
      secoes: [
        { tipo: "Estrofe", texto: "Senhor, nossa paz,\ntende piedade de nós.\nSenhor, Senhor, tende piedade de nós. (Bis)" },
        { tipo: "Estrofe", texto: "Cristo, nossa Páscoa,\ntende piedade de nós.\nCristo, Cristo, tende piedade de nós. (Bis)" },
        { tipo: "Estrofe", texto: "Senhor, nossa vida,\ntende piedade de nós.\nSenhor, Senhor, tende piedade de nós. (Bis)" },
      ]
    },
  ],

  gloria: [
    {
      titulo: "Glória (Vinde Cristãos)",
      autor: "",
      secoes: [
        { tipo: "Refrão", texto: "Glória a Deus nas alturas! (Bis)" },
        { tipo: "Estrofe", texto: "1. Glória a Deus nos altos céus,\npaz na terra aos seus amados.\nA vós louvam, Rei celeste,\nos que foram libertados." },
        { tipo: "Refrão", texto: "Glória a Deus nas alturas! (Bis)" },
        { tipo: "Estrofe", texto: "2. Deus e Pai nós vos louvamos,\nadoramos, bendizemos.\nDamos glória ao vosso nome,\nvossos dons agradecemos." },
        { tipo: "Refrão", texto: "Glória a Deus nas alturas! (Bis)" },
        { tipo: "Estrofe", texto: "3. Senhor nosso Jesus Cristo,\nunigênito do Pai.\nVós de Deus Cordeiro Santo,\nnossas culpas perdoai." },
        { tipo: "Refrão", texto: "Glória a Deus nas alturas! (Bis)" },
        { tipo: "Estrofe", texto: "4. Vós que estais junto do Pai,\ncomo nosso intercessor.\nAcolhei nossos pedidos,\natendei nosso clamor." },
        { tipo: "Refrão", texto: "Glória a Deus nas alturas! (Bis)" },
        { tipo: "Estrofe", texto: "5. Vós somente sois o Santo,\no Altíssimo, o Senhor,\ncom o Espírito Divino,\nde Deus Pai no esplendor." },
        { tipo: "Refrão", texto: "Glória a Deus nas alturas! (Bis)" },
      ]
    },
    {
      titulo: "Glória, Glória, Glória a Deus",
      autor: "",
      secoes: [
        { tipo: "Refrão", texto: "Glória, glória, glória a Deus\nnas alturas e na terra paz aos homens!" },
        { tipo: "Estrofe", texto: "1. Senhor Deus, Rei dos céus,\nDeus Pai todo-poderoso.\nNós vos louvamos, nós vos bendizemos,\nnós vos adoramos, nós vos glorificamos.\nNós vos damos graças por vossa imensa glória." },
        { tipo: "Refrão", texto: "Glória, glória, glória a Deus\nnas alturas e na terra paz aos homens!" },
        { tipo: "Estrofe", texto: "2. Senhor Jesus Cristo, Filho Unigênito,\nSenhor Deus, Cordeiro de Deus, Filho de Deus Pai.\nVós que tirais o pecado do mundo,\ntende piedade de nós.\nVós que tirais o pecado do mundo,\nacolhei a nossa súplica.\nVós que estais à direita do Pai,\ntende piedade de nós." },
        { tipo: "Refrão", texto: "Glória, glória, glória a Deus\nnas alturas e na terra paz aos homens!" },
        { tipo: "Estrofe", texto: "3. Só vós sois Santo, só vós o Senhor,\nsó vós o Altíssimo, Jesus Cristo!\nCom o Espírito Santo,\nna glória de Deus Pai. Amém." },
        { tipo: "Refrão", texto: "Glória, glória, glória a Deus\nnas alturas e na terra paz aos homens!" },
      ]
    },
    {
      titulo: "Glória a Deus nas Alturas",
      autor: "Casemiro Nogueira",
      secoes: [
        { tipo: "Refrão", texto: "Glória a Deus nas alturas,\nglória a Deus nas alturas,\ne paz na terra aos homens por Ele amados." },
        { tipo: "Estrofe", texto: "1. Senhor Deus, Rei dos céus, Deus Pai todo-poderoso.\nNós vos louvamos, nós vos bendizemos,\nnós vos adoramos, nós vos glorificamos,\nnós vos damos graças por vossa imensa glória." },
        { tipo: "Refrão", texto: "Glória a Deus nas alturas,\nglória a Deus nas alturas,\ne paz na terra aos homens por Ele amados." },
        { tipo: "Estrofe", texto: "2. Senhor Jesus Cristo, Filho Unigênito,\nSenhor Deus, Cordeiro de Deus, Filho de Deus Pai.\nVós que tirais o pecado do mundo,\ntende piedade de nós.\nVós que tirais o pecado do mundo,\nacolhei a nossa súplica.\nVós que estais à direita do Pai,\ntende piedade de nós." },
        { tipo: "Refrão", texto: "Glória a Deus nas alturas,\nglória a Deus nas alturas,\ne paz na terra aos homens por Ele amados." },
        { tipo: "Estrofe", texto: "3. Só vós sois Santo.\nSó vós o Senhor,\nsó vós o Altíssimo, Jesus Cristo!\nCom o Espírito Santo,\nna glória de Deus Pai. Amém!" },
        { tipo: "Refrão", texto: "Glória a Deus nas alturas,\nglória a Deus nas alturas,\ne paz na terra aos homens por Ele amados." },
      ]
    },
    {
      titulo: "Glória a Deus nos Altos Céus",
      autor: "",
      secoes: [
        { tipo: "Refrão", texto: "Glória a Deus lá nos céus\ne paz aos seus, amém. (Bis)" },
        { tipo: "Estrofe", texto: "1. Glória a Deus nos altos céus,\npaz na terra a seus amados.\nA vós louvam, Rei celeste,\nos que foram libertados." },
        { tipo: "Refrão", texto: "Glória a Deus lá nos céus\ne paz aos seus, amém. (Bis)" },
        { tipo: "Estrofe", texto: "2. Deus e Pai nós vos louvamos,\nadoramos, bendizemos.\nDamos glória ao vosso nome,\nvossos dons agradecemos." },
        { tipo: "Refrão", texto: "Glória a Deus lá nos céus\ne paz aos seus, amém. (Bis)" },
        { tipo: "Estrofe", texto: "3. Senhor nosso Jesus Cristo,\nunigênito do Pai.\nVós de Deus Cordeiro Santo,\nnossas culpas perdoai." },
        { tipo: "Refrão", texto: "Glória a Deus lá nos céus\ne paz aos seus, amém. (Bis)" },
        { tipo: "Estrofe", texto: "4. Vós que estais junto do Pai,\ncomo nosso intercessor.\nAcolhei nossos pedidos,\natendei nosso clamor." },
        { tipo: "Refrão", texto: "Glória a Deus lá nos céus\ne paz aos seus, amém. (Bis)" },
        { tipo: "Estrofe", texto: "5. Vós somente sois o Santo,\no Altíssimo, o Senhor,\ncom o Espírito Divino,\nde Deus Pai no esplendor!" },
        { tipo: "Refrão", texto: "Glória a Deus lá nos céus\ne paz aos seus, amém. (Bis)" },
      ]
    },
    {
      titulo: "Glória a Deus lá nos Céus, e Paz na Terra aos Seus!",
      autor: "",
      secoes: [
        { tipo: "Refrão", texto: "Glória a Deus lá nos céus,\ne paz na terra aos seus! (Bis)" },
        { tipo: "Estrofe", texto: "1. Glória a Deus nos altos céus,\npaz na terra a seus amados,\na vós louvam, Rei celeste,\nos que foram libertados!" },
        { tipo: "Refrão", texto: "Glória a Deus lá nos céus,\ne paz na terra aos seus! (Bis)" },
        { tipo: "Estrofe", texto: "2. Deus e Pai, nós vos louvamos,\nadoramos, bendizemos;\ndamos glória ao vosso nome,\nvossos dons agradecemos!" },
        { tipo: "Refrão", texto: "Glória a Deus lá nos céus,\ne paz na terra aos seus! (Bis)" },
        { tipo: "Estrofe", texto: "3. Senhor nosso, Jesus Cristo,\nunigênito do Pai,\nvós de Deus, Cordeiro Santo,\nnossas culpas perdoai!" },
        { tipo: "Refrão", texto: "Glória a Deus lá nos céus,\ne paz na terra aos seus! (Bis)" },
        { tipo: "Estrofe", texto: "4. Vós que estais junto do Pai,\ncomo nosso intercessor,\nacolhei nossos pedidos,\natendei nosso clamor!" },
        { tipo: "Refrão", texto: "Glória a Deus lá nos céus,\ne paz na terra aos seus! (Bis)" },
        { tipo: "Estrofe", texto: "5. Vós somente sois o Santo,\no Altíssimo, o Senhor,\ncom o Espírito Divino,\nde Deus Pai no esplendor!" },
        { tipo: "Refrão", texto: "Glória a Deus lá nos céus,\ne paz na terra aos seus! (Bis)" },
      ]
    },
    {
      titulo: "Glória, Glória a Deus nas Alturas",
      autor: "",
      secoes: [
        { tipo: "Refrão", texto: "Glória, glória, a Deus nas alturas. (Bis)" },
        { tipo: "Estrofe", texto: "1. Glória a Deus nas alturas\ne paz na terra aos homens por Ele amados.\nSenhor Deus, Rei dos céus, Deus Pai todo-poderoso.\nNós vos louvamos, nós vos bendizemos,\nnós vos adoramos, nós vos glorificamos.\nNós vos damos graças por vossa imensa glória." },
        { tipo: "Refrão", texto: "Glória, glória, a Deus nas alturas. (Bis)" },
        { tipo: "Estrofe", texto: "2. Senhor Jesus Cristo, Filho Unigênito,\nSenhor Deus, Cordeiro de Deus, Filho de Deus Pai.\nVós que tirais o pecado do mundo,\ntende piedade de nós.\nVós que tirais o pecado do mundo,\nacolhei a nossa súplica.\nVós que estais à direita do Pai,\ntende piedade de nós." },
        { tipo: "Refrão", texto: "Glória, glória, a Deus nas alturas. (Bis)" },
        { tipo: "Estrofe", texto: "3. Só vós sois Santo.\nSó vós o Senhor,\nsó vós o Altíssimo, Jesus Cristo.\nCom o Espírito Santo,\nna glória de Deus Pai, amém!" },
        { tipo: "Refrão", texto: "Glória, glória, a Deus nas alturas. (Bis)" },
      ]
    },
    {
      titulo: "Glória Shalom",
      autor: "",
      secoes: [
        { tipo: "Refrão", texto: "Glória, glória a Deus nas alturas,\nô, ô glória e a nós a Sua paz. (Bis)" },
        { tipo: "Estrofe", texto: "1. Senhor Deus, Rei dos céus, Deus Pai onipotente,\nvos louvamos, bendizemos, adoramos.\nNós vos glorificamos\ne nós vos damos graças por vossa imensa glória." },
        { tipo: "Refrão", texto: "Glória, glória a Deus nas alturas,\nô, ô glória e a nós a Sua paz. (Bis)" },
        { tipo: "Estrofe", texto: "2. Jesus Cristo, Senhor Deus, Filho único do Pai,\nCordeiro de Deus que tirais o pecado do mundo,\ntende piedade.\nVós que estais à direita do Pai, tende piedade.\nVós que tirais o pecado do mundo,\nacolhei a nossa súplica." },
        { tipo: "Refrão", texto: "Glória, glória a Deus nas alturas,\nô, ô glória e a nós a Sua paz. (Bis)" },
        { tipo: "Estrofe", texto: "3. Só vós sois o Santo, o Senhor,\no Altíssimo, só vós,\nJesus Cristo, com o Espírito e o Pai." },
        { tipo: "Refrão", texto: "Glória, glória a Deus nas alturas,\nô, ô glória e a nós a Sua paz. (Bis)" },
      ]
    },
    {
      titulo: "Ô Ô Glória",
      autor: "",
      secoes: [
        { tipo: "Refrão", texto: "Ô ô glória, ô glória!\nA Deus Pai, a Jesus Cristo\ne ao Espírito de amor. (Bis)" },
        { tipo: "Estrofe", texto: "1. Deus e Pai nós vos louvamos,\nadoramos, bendizemos.\nDamos glória a vosso nome,\nvossos dons agradecemos." },
        { tipo: "Refrão", texto: "Ô ô glória, ô glória!\nA Deus Pai, a Jesus Cristo\ne ao Espírito de amor. (Bis)" },
        { tipo: "Estrofe", texto: "2. Senhor nosso Jesus Cristo,\nunigênito do Pai,\nvós de Deus, Cordeiro Santo,\nnossas culpas perdoai." },
        { tipo: "Refrão", texto: "Ô ô glória, ô glória!\nA Deus Pai, a Jesus Cristo\ne ao Espírito de amor. (Bis)" },
        { tipo: "Estrofe", texto: "3. Vós que estais junto do Pai,\ncomo nosso intercessor,\nacolhei nossos pedidos,\natendei nosso clamor." },
        { tipo: "Refrão", texto: "Ô ô glória, ô glória!\nA Deus Pai, a Jesus Cristo\ne ao Espírito de amor. (Bis)" },
        { tipo: "Estrofe", texto: "4. Vós somente sois o Santo,\no Altíssimo, o Senhor,\ncom o Espírito Divino,\nde Deus Pai no esplendor." },
        { tipo: "Refrão", texto: "Ô ô glória, ô glória!\nA Deus Pai, a Jesus Cristo\ne ao Espírito de amor. (Bis)" },
      ]
    },
    {
      titulo: "Glória, Glória, Anjos do Céu",
      autor: "",
      secoes: [
        { tipo: "Refrão", texto: "Glória, glória, anjos do céu,\ncantam todos Seu amor.\nE na terra, homens de paz,\nDeus merece o louvor. (Bis)" },
        { tipo: "Estrofe", texto: "1. Deus e Pai, nós vos louvamos,\nadoramos, bendizemos.\nDamos glória ao vosso nome,\nvossos dons agradecemos." },
        { tipo: "Refrão", texto: "Glória, glória, anjos do céu,\ncantam todos Seu amor.\nE na terra, homens de paz,\nDeus merece o louvor. (Bis)" },
        { tipo: "Estrofe", texto: "2. Senhor nosso, Jesus Cristo,\nunigênito do Pai.\nVós de Deus Cordeiro Santo,\nnossas culpas perdoai." },
        { tipo: "Refrão", texto: "Glória, glória, anjos do céu,\ncantam todos Seu amor.\nE na terra, homens de paz,\nDeus merece o louvor. (Bis)" },
        { tipo: "Estrofe", texto: "3. Vós que estais junto do Pai,\ncomo nosso intercessor.\nAcolhei nossos pedidos,\natendei nosso clamor." },
        { tipo: "Refrão", texto: "Glória, glória, anjos do céu,\ncantam todos Seu amor.\nE na terra, homens de paz,\nDeus merece o louvor. (Bis)" },
        { tipo: "Estrofe", texto: "4. Vós somente sois o Santo,\no Altíssimo, o Senhor,\ncom o Espírito Divino,\nde Deus Pai no esplendor." },
        { tipo: "Refrão", texto: "Glória, glória, anjos do céu,\ncantam todos Seu amor.\nE na terra, homens de paz,\nDeus merece o louvor. (Bis)" },
      ]
    },
    {
      titulo: "Glória a Deus nas Alturas",
      autor: "Pe. Ney Brasil",
      secoes: [
        { tipo: "Refrão", texto: "Glória a Deus nas alturas,\ne paz na terra aos homens por Ele amados. (Bis)" },
        { tipo: "Estrofe", texto: "1. Senhor Deus, Rei dos céus, Deus Pai todo-poderoso:\nnós vos louvamos, nós vos bendizemos,\nnós vos adoramos, nós vos glorificamos." },
        { tipo: "Refrão", texto: "Glória a Deus nas alturas,\ne paz na terra aos homens por Ele amados. (Bis)" },
        { tipo: "Estrofe", texto: "2. Nós vos damos graças por vossa imensa glória.\nSenhor Jesus Cristo, Filho Unigênito,\nSenhor Deus, Cordeiro de Deus, Filho de Deus Pai." },
        { tipo: "Refrão", texto: "Glória a Deus nas alturas,\ne paz na terra aos homens por Ele amados. (Bis)" },
        { tipo: "Estrofe", texto: "3. Vós que tirais o pecado do mundo,\ntende piedade de nós.\nVós que tirais o pecado do mundo,\nacolhei a nossa súplica.\nVós que estais à direita do Pai,\ntende piedade de nós." },
        { tipo: "Refrão", texto: "Glória a Deus nas alturas,\ne paz na terra aos homens por Ele amados. (Bis)" },
        { tipo: "Estrofe", texto: "4. Só vós sois o Santo, só vós, o Senhor,\nsó vós, o Altíssimo, Jesus Cristo,\ncom o Espírito Santo, na glória de Deus Pai." },
        { tipo: "Refrão", texto: "Glória a Deus nas alturas,\ne paz na terra aos homens por Ele amados. (Bis)" },
      ]
    },
  ],

  aclamacao: [
    {
      titulo: "Aleluia (1)",
      autor: "",
      secoes: [
        { tipo: "Refrão", texto: "Aleluia! Aleluia! Aleluia! (Bis)" },
        { tipo: "Estrofe", texto: "Vinde Espírito Divino,\ne enchei com vossos dons\nos corações dos fiéis,\ne acendei neles o amor,\ncomo um fogo abrasador." },
        { tipo: "Refrão", texto: "Aleluia! Aleluia! Aleluia! (Bis)" },
      ]
    },
    {
      titulo: "Aleluia (2)",
      autor: "",
      secoes: [
        { tipo: "Refrão", texto: "Aleluia! Aleluia! Aleluia! Aleluia! (Bis)" },
        { tipo: "Estrofe", texto: "Vinde Espírito Divino,\ne enchei com vossos dons\nos corações dos fiéis,\ne acendei neles o amor,\ncomo um fogo abrasador." },
        { tipo: "Refrão", texto: "Aleluia! Aleluia! Aleluia! Aleluia! (Bis)" },
      ]
    },
  ],

  ofertorio: [
    {
      titulo: "Espírito Criador!",
      autor: "Pe. Zezinho",
      secoes: [
        { tipo: "Refrão", texto: "Espírito criador!\nBendito sempre sejais!\nPor tudo isso, e bem mais,\npelo imenso dom do amor." },
        { tipo: "Estrofe", texto: "1. Com o Pai fazeis fecundo\no solo imenso do mundo\npra nos dar trigo e flor.\nBendito sois noite e dia\npor tão grande doação,\nfonte sem fim de alegria,\nsão matérias pro nosso pão." },
        { tipo: "Refrão", texto: "Espírito criador!\nBendito sempre sejais!\nPor tudo isso, e bem mais,\npelo imenso dom do amor." },
        { tipo: "Estrofe", texto: "2. Foi dom de vossa bondade\nenchernos de habilidade\npro trabalho, Senhor.\nCom o Pai vós sois bendito\nporque dais à nossa mão,\ncom poder que é quase infinito,\ncontinuar a criação." },
        { tipo: "Refrão", texto: "Espírito criador!\nBendito sempre sejais!\nPor tudo isso, e bem mais,\npelo imenso dom do amor." },
      ]
    },
    {
      titulo: "Pão e Vinho, Pai, Poremos",
      autor: "Pe. Zezinho",
      secoes: [
        { tipo: "Refrão", texto: "Pão e vinho, Pai, poremos,\nnesta mesa, uma vez mais.\nÉ um pouco do que temos\npelo muito que nos dais." },
        { tipo: "Estrofe", texto: "1. Vós nos dais Jesus, o Cristo,\nmas o Cristo, o que nos faz?\nVem morrer crucificado,\npara vir ressuscitado,\ne nos dar a sua paz." },
        { tipo: "Refrão", texto: "Pão e vinho, Pai, poremos,\nnesta mesa, uma vez mais.\nÉ um pouco do que temos\npelo muito que nos dais." },
        { tipo: "Estrofe", texto: "2. Vós nos dais o vosso Filho\npara ser o nosso irmão.\nE pra termos de verdade,\nsó amor, fraternidade,\nEle nos deu o perdão." },
        { tipo: "Refrão", texto: "Pão e vinho, Pai, poremos,\nnesta mesa, uma vez mais.\nÉ um pouco do que temos\npelo muito que nos dais." },
      ]
    },
    {
      titulo: "As Sementes que Me Deste",
      autor: "",
      secoes: [
        { tipo: "Refrão", texto: "Dos meus dons que recebi\npelo Espírito do amor,\ntrago os frutos que colhi\ne em Tua mesa quero pôr. (Bis)" },
        { tipo: "Estrofe", texto: "1. As sementes que me deste\ne que não eram pra guardar,\npus no chão da minha vida,\nquis fazer frutificar." },
        { tipo: "Refrão", texto: "Dos meus dons que recebi\npelo Espírito do amor,\ntrago os frutos que colhi\ne em Tua mesa quero pôr. (Bis)" },
        { tipo: "Estrofe", texto: "2. Pelos campos deste mundo\nquero sempre semear,\nos talentos que me deste\npara eu mesmo cultivar." },
        { tipo: "Refrão", texto: "Dos meus dons que recebi\npelo Espírito do amor,\ntrago os frutos que colhi\ne em Tua mesa quero pôr. (Bis)" },
      ]
    },
    {
      titulo: "Suscitai, ó Senhor Deus",
      autor: "",
      secoes: [
        { tipo: "Refrão", texto: "Suscitai, ó Senhor Deus,\nsuscitai vosso poder,\nconfirmai este poder\nque por nós manifestastes!" },
        { tipo: "Estrofe", texto: "1. Contemplamos, ó Senhor,\nvosso cortejo que desfila,\né a entrada do meu Deus,\ndo meu Rei, no santuário." },
        { tipo: "Refrão", texto: "Suscitai, ó Senhor Deus,\nsuscitai vosso poder,\nconfirmai este poder\nque por nós manifestastes!" },
        { tipo: "Estrofe", texto: "2. Os cantores vão à frente,\nvão atrás os tocadores,\ne no meio vão as jovens\na tocar seus tamborins." },
        { tipo: "Refrão", texto: "Suscitai, ó Senhor Deus,\nsuscitai vosso poder,\nconfirmai este poder\nque por nós manifestastes!" },
      ]
    },
    {
      titulo: "As Nossas Ofertas de Vinho e de Pão",
      autor: "",
      secoes: [
        { tipo: "Refrão", texto: "As nossas ofertas de vinho e de pão\ncelebram a glória da ressurreição,\na glória da ressurreição." },
        { tipo: "Estrofe", texto: "1. O grão que morrera no seio do chão\nrenasce no trigo tornando-se pão.\nA uva amassada, pisada, moída,\nressurge no vinho, sustento da vida." },
        { tipo: "Refrão", texto: "As nossas ofertas de vinho e de pão\ncelebram a glória da ressurreição,\na glória da ressurreição." },
        { tipo: "Estrofe", texto: "2. O pão e o vinho são hoje memória\ndo novo cordeiro na sua vitória.\nSinais da aliança da terra e dos céus\nno corpo e no sangue do Filho de Deus." },
        { tipo: "Refrão", texto: "As nossas ofertas de vinho e de pão\ncelebram a glória da ressurreição,\na glória da ressurreição." },
      ]
    },
    {
      titulo: "Bendito Sejas, ó Rei da Glória",
      autor: "",
      secoes: [
        { tipo: "Refrão", texto: "Aqui trazemos as nossas ofertas,\nvê com bons olhos nossas humildes ofertas.\nTudo que temos seja pra Ti, ó Senhor." },
        { tipo: "Estrofe", texto: "1. Bendito sejas, ó Rei da glória,\nressuscitado, Senhor da Igreja.\nVidas se encontram no altar de Deus,\ngente se doa, dom que se imola." },
        { tipo: "Refrão", texto: "Aqui trazemos as nossas ofertas,\nvê com bons olhos nossas humildes ofertas.\nTudo que temos seja pra Ti, ó Senhor." },
        { tipo: "Estrofe", texto: "2. Maior motivo de oferenda,\npois o Senhor ressuscitou,\npara que todos tivessem vida." },
        { tipo: "Refrão", texto: "Aqui trazemos as nossas ofertas,\nvê com bons olhos nossas humildes ofertas.\nTudo que temos seja pra Ti, ó Senhor." },
      ]
    },
    {
      titulo: "Cristo é o Dom do Pai",
      autor: "",
      secoes: [
        { tipo: "Refrão", texto: "Cristo é o dom do Pai\nque se entregou por nós.\nAleluia, aleluia!\nBendito seja o nosso Deus!" },
        { tipo: "Estrofe", texto: "1. Dai graças a Deus, pois Ele é bom;\neterno por nós é seu amor." },
        { tipo: "Refrão", texto: "Cristo é o dom do Pai\nque se entregou por nós.\nAleluia, aleluia!\nBendito seja o nosso Deus!" },
        { tipo: "Estrofe", texto: "2. Coragem e força Ele nos dá,\nfazendo-se nosso Salvador." },
        { tipo: "Refrão", texto: "Cristo é o dom do Pai\nque se entregou por nós.\nAleluia, aleluia!\nBendito seja o nosso Deus!" },
      ]
    },
  ],

  santo: [
    {
      titulo: "Santo III",
      secoes: [
        { tipo: "Única", texto: "Santo, Santo, Santo.\nSenhor Deus do universo.\nO céu e a terra\nproclam a vossa glória.\nHosana nas alturas." },
        { tipo: "Única", texto: "Bendito o que vem\nem nome do Senhor.\nHosana, nas alturas." },
      ]
    },
    {
      titulo: "Santo - Hosana nas Alturas",
      secoes: [
        { tipo: "Única", texto: "Santo, Santo, Santo.\nSenhor Deus do universo.\nO céu e a terra proclamam\na vossa glória.\nHosana nas alturas. Hosana! (Bis)" },
        { tipo: "Única", texto: "Bendito Aquele que vem\nem nome do Senhor. (Bis)" },
      ]
    },
    {
      titulo: "Santo (I) - Pe. Ney Brasil",
      secoes: [
        { tipo: "Única", texto: "Santo, Santo, Santo sois Senhor\nSanto, Santo, Santo sois Senhor Nosso Deus.\nSenhor Deus do universo.\nO céu e a terra proclamam a vossa glória.\nHosana nas alturas." },
        { tipo: "Única", texto: "Santo, Santo, Santo Sois Senhor\nSanto, Santo, Santo Sois Senhor Nosso Deus.\nBendito o que vem\nem nome do Senhor.\nHosana nas alturas, hosana nas alturas." },
      ]
    },
    {
      titulo: "Santo - Pedro Ivan",
      secoes: [
        { tipo: "Única", texto: "Santo, Santo, Santo.\nSenhor Deus do universo.\nO céu e a terra\nproclam a vossa glória.\nHosana nas alturas." },
        { tipo: "Única", texto: "Bendito o que vem\nem nome do Senhor.\nHosana, nas alturas." },
      ]
    },
    {
      titulo: "O Senhor é Santo",
      secoes: [
        { tipo: "Única", texto: "O Senhor é Santo!\nO Senhor é Santo!\nO Senhor é Santo!\nSenhor Deus do universo.\nO céu e a terra\nproclam a vossa glória.\nHosana nas alturas." },
        { tipo: "Única", texto: "O Senhor é Santo!\nO Senhor é Santo!\nO Senhor é Santo!\nBendito o que vem em nome do Senhor.\nBendito o que vem em nome do Senhor.\nHosana! Hosana! Hosana! Hosana!" },
      ]
    },
    {
      titulo: "Santo - Padre Cleidimar Moreira",
      secoes: [
        { tipo: "Única", texto: "Santo, Santo, Santo é o Senhor.\nDeus do universo. (Bis)\nOs céus e a terra\nproclam a vossa glória.\nHosana nas alturas.\nHosana nas alturas. (Bis)" },
        { tipo: "Única", texto: "Santo, Santo, Santo é o Senhor.\nDeus do universo. (Bis)\nBendito o que vem\nem nome do Senhor.\nHosana, nas alturas.\nHosana nas alturas. (Bis)" },
      ]
    },
    {
      titulo: "Santo, Santo, Santo é o Senhor",
      secoes: [
        { tipo: "Única", texto: "Santo, Santo, Santo é o Senhor\nDeus do universo, do céu e a terra. (Bis)\nHosana! Hosana!\nHosana! nas alturas. (Bis)" },
        { tipo: "Única", texto: "Bendito o que vem\nem nome do Senhor.\nHosana, nas alturas. (Bis)" },
      ]
    },
  ],

  comunhao: [
    {
      titulo: "Ao Recebermos, Senhor",
      autor: "",
      secoes: [
        { tipo: "Refrão", texto: "Desamarrem as sandálias e descansem,\neste chão é terra santa, irmãos meus.\nVenham, orem, comam, cantem, venham todos,\ne renovem a esperança no Senhor." },
        { tipo: "Estrofe", texto: "1. Ao recebermos, Senhor, tua presença sagrada,\npra confirmar Teu amor, faz de nós Tua morada.\nSurge um sincero louvor, brota a semente plantada.\nFaz-nos seguir Teu caminho, sempre trilhar Tua estrada." },
        { tipo: "Refrão", texto: "Desamarrem as sandálias e descansem,\neste chão é terra santa, irmãos meus.\nVenham, orem, comam, cantem, venham todos,\ne renovem a esperança no Senhor." },
        { tipo: "Estrofe", texto: "2. O Filho de Deus com o Pai e o Espírito Santo,\nnesta trindade, um só ser que pede a nós sermos santos.\nDai-nos, Jesus, Teu poder de se doar sem medidas,\ndeixa que compreendamos que este é o sentido da vida." },
        { tipo: "Refrão", texto: "Desamarrem as sandálias e descansem,\neste chão é terra santa, irmãos meus.\nVenham, orem, comam, cantem, venham todos,\ne renovem a esperança no Senhor." },
      ]
    },
    {
      titulo: "Cantar a Beleza da Vida",
      autor: "",
      secoes: [
        { tipo: "Refrão", texto: "Vem dar-nos Teu Filho, Senhor,\nsustento no pão e no vinho,\ne a força do Espírito Santo,\nunindo Teu povo a caminho!" },
        { tipo: "Estrofe", texto: "1. Cantar a beleza da vida,\npresente do amor sem igual:\nmissão do Teu povo escolhido!\nSenhor, vem livrar-nos do mal!" },
        { tipo: "Refrão", texto: "Vem dar-nos Teu Filho, Senhor,\nsustento no pão e no vinho,\ne a força do Espírito Santo,\nunindo Teu povo a caminho!" },
        { tipo: "Estrofe", texto: "2. Falar do Teu Filho às nações,\nvivendo como Ele viveu:\nmissão do Teu povo escolhido!\nSenhor, vem cuidar do que é Teu!" },
        { tipo: "Refrão", texto: "Vem dar-nos Teu Filho, Senhor,\nsustento no pão e no vinho,\ne a força do Espírito Santo,\nunindo Teu povo a caminho!" },
      ]
    },
    {
      titulo: "Todos Ficaram Cheios do Espírito Santo",
      autor: "",
      secoes: [
        { tipo: "Refrão", texto: "Todos ficaram cheios do Espírito Santo\ne proclamavam as maravilhas de Deus.\nAleluia, aleluia!" },
        { tipo: "Estrofe", texto: "1. Eis que Deus se põe de pé,\ne os inimigos se dispersam!\nFogem longe de sua face\nos que odeiam o Senhor." },
        { tipo: "Refrão", texto: "Todos ficaram cheios do Espírito Santo\ne proclamavam as maravilhas de Deus.\nAleluia, aleluia!" },
        { tipo: "Estrofe", texto: "2. Mas os justos se alegram\nna presença do Senhor,\nrejubilam satisfeitos\ne exultam de alegria." },
        { tipo: "Refrão", texto: "Todos ficaram cheios do Espírito Santo\ne proclamavam as maravilhas de Deus.\nAleluia, aleluia!" },
      ]
    },
    {
      titulo: "Senhor, Vem Dar-nos Sabedoria",
      autor: "Pe. Zezinho",
      secoes: [
        { tipo: "Refrão", texto: "Dá-nos, Senhor, esses dons, essa luz,\ne nós veremos que Pão é Jesus! (Bis)" },
        { tipo: "Estrofe", texto: "1. Senhor, vem dar-nos sabedoria\nque faz ter tudo como Deus quis,\ne assim faremos da Eucaristia\no grande meio de ser feliz." },
        { tipo: "Refrão", texto: "Dá-nos, Senhor, esses dons, essa luz,\ne nós veremos que Pão é Jesus! (Bis)" },
        { tipo: "Estrofe", texto: "2. Dá-nos, Senhor, o entendimento\nque tudo ajuda a compreender,\npara nós vermos como é alimento\no pão e o vinho que Deus quer ser." },
        { tipo: "Refrão", texto: "Dá-nos, Senhor, esses dons, essa luz,\ne nós veremos que Pão é Jesus! (Bis)" },
      ]
    },
  ],

  final: [
    {
      titulo: "O Espírito de Deus Repousa Sobre Mim",
      autor: "",
      secoes: [
        { tipo: "Refrão", texto: "Glória, glória eterna,\nglória a Ti, Senhor. (Bis)" },
        { tipo: "Estrofe", texto: "1. O Espírito de Deus repousa sobre mim\ne assim caminhando eu vou.\nAlegria, paz e amor,\nfruto que vem de Ti, Senhor,\nem mim brotou." },
        { tipo: "Refrão", texto: "Glória, glória eterna,\nglória a Ti, Senhor. (Bis)" },
        { tipo: "Estrofe", texto: "2. Servi ao Senhor com toda alegria,\nvinde, exultaremos.\nSaber que o Senhor é Deus e Salvador,\ne só a Ele pertencemos." },
        { tipo: "Refrão", texto: "Glória, glória eterna,\nglória a Ti, Senhor. (Bis)" },
      ]
    },
    {
      titulo: "Vem, Vem, Vem, Espírito Santo",
      autor: "",
      secoes: [
        { tipo: "Refrão", texto: "Vem, vem, vem, Espírito Santo,\ntransforma a minha vida,\nquero renascer. (Bis)" },
        { tipo: "Estrofe", texto: "1. Quero abandonar-me em Seu amor,\nencharcar-me em Seus rios, Senhor,\nderrubar as barreiras em meu coração. (Bis)" },
        { tipo: "Refrão", texto: "Vem, vem, vem, Espírito Santo,\ntransforma a minha vida,\nquero renascer. (Bis)" },
      ]
    },
    {
      titulo: "Novo Sol Brilhou",
      autor: "",
      secoes: [
        { tipo: "Refrão", texto: "O Deus de amor jamais se descuidou,\nem seu vigor, Jesus ressuscitou. (Bis)" },
        { tipo: "Estrofe", texto: "1. Novo sol brilhou, a vida superou\nsofrimento, dor e morte, tudo enfim.\nNosso olhar se abriu, Deus mesmo se incumbiu\nde tomar-nos pela mão assim." },
        { tipo: "Refrão", texto: "O Deus de amor jamais se descuidou,\nem seu vigor, Jesus ressuscitou. (Bis)" },
        { tipo: "Estrofe", texto: "2. Estender a mão, abrir o coração,\nacolher, compartilhar e perdoar,\né fazer o céu cumprir o seu papel,\njá na terra tem que vigorar." },
        { tipo: "Refrão", texto: "O Deus de amor jamais se descuidou,\nem seu vigor, Jesus ressuscitou. (Bis)" },
      ]
    },
    {
      titulo: "Rainha do Céu",
      autor: "",
      secoes: [
        { tipo: "Única", texto: "Rainha do Céu, alegrai-vos, Aleluia!\nAleluia! Aleluia! Aleluia! Aleluia! (Bis)" },
      ]
    },
  ],
};

// =============================================
// SLIDES FIXOS DA LITURGIA DA PALAVRA
// =============================================

const slidesFixos = {
  primeiraLeitura: [
    "I LEITURA:",
    "Atos dos Apóstolos",
    "2,1-11",
  ],
  salmo: [
    "Salmo responsorial 103(104)",
    "R. Enviai o vosso Espírito, Senhor.",
    "Enviai o vosso Espírito, Senhor.",
    "E da terra toda a face renovai.",
    "E da terra toda a face renovai!",
  ],
  segundaLeitura: [
    "II LEITURA:",
    "I Coríntios 12,3b-7.12-13",
  ],
  aclamacao: [
    "R. Aleluia, aleluia, aleluia.",
    "V. Quem me ama realmente",
    "guardará a minha palavra.",
  ],
  evangelho: [
    "EVANGELHO",
    "João 20,19-23",
  ],
  preces: [
    "Oração dos Fiéis",
    "R. Mandai, Senhor,",
    "o vosso Espírito.",
  ],
};

// Oração Eucarística I — respostas do povo
const oracaoEucaristica = [
  "Abençoai nossa oferenda, ó Senhor!",
  "Lembrai-vos, ó Pai, de vossos filhos!",
  "Em comunhão com os vossos Santos\nvos louvamos!",
  "Enviai o vosso Espírito Santo!",
  // TELA PRETA aqui
  "Anunciamos, Senhor, a vossa morte\ne proclamamos a vossa ressurreição.\nVinde, Senhor Jesus!",
  "Todas as vezes que comemos deste pão\ne bebemos deste cálice,\nanunciamos, Senhor, a vossa morte,\nenquanto esperamos a vossa vinda!",
  "Aceitai, ó Senhor, a nossa oferta!",
  "O Espírito nos una num só corpo!",
  "Concedei-lhes, ó Senhor, a luz eterna!",
];

// =============================================
// FUNÇÃO PRINCIPAL DE GERAÇÃO
// =============================================

function gerarMissa(opcoes) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_4x3";
  pres.title = "LiturgiaPlay — Missa de Pentecostes";

  // 1. CAPA
  addSlideCapa(pres, "Domingo de Pentecostes — 01 de junho de 2025", "Solenidade de\nPentecostes");

  // 2. CANTO DE ENTRADA
  const entrada = banco.cantos_entrada.find(c => c.titulo === opcoes.entrada);
  if (entrada) renderMusica(pres, entrada);

  // 3. PRETO
  addSlidePreto(pres);

  // 4. ASPERSÃO (opcional)
  if (opcoes.aspersao) {
    const asp = banco.aspersao.find(c => c.titulo === opcoes.aspersao);
    if (asp) renderMusica(pres, asp);
    addSlidePreto(pres);
  }

  // 5. VELA (opcional)
  if (opcoes.vela) {
    const vela = banco.vela.find(c => c.titulo === opcoes.vela);
    if (vela) renderMusica(pres, vela);
    addSlidePreto(pres);
  }

  // 6. ATO PENITENCIAL
  if (opcoes.penitencial) {
    const pen = banco.penitencial.find(c => c.titulo === opcoes.penitencial);
    if (pen) renderMusica(pres, pen);
  }
  addSlidePreto(pres);

  // 7. GLÓRIA
  if (opcoes.gloria) {
    const gl = banco.gloria.find(c => c.titulo === opcoes.gloria);
    if (gl) renderMusica(pres, gl);
  }
  addSlidePreto(pres);

  // 8. LITURGIA DA PALAVRA
  addSlideFixo(pres, "I LEITURA:", ["Atos dos Apóstolos", "2,1-11"]);
  addSlideFixo(pres, "Salmo responsorial 103(104)", ["R. Enviai o vosso Espírito, Senhor.", "Enviai o vosso Espírito, Senhor.", "E da terra toda a face renovai.", "E da terra toda a face renovai!"]);
  addSlideFixo(pres, "II LEITURA:", ["I Coríntios 12,3b-7.12-13"]);
  addSlideFixo(pres, "Aclamação", ["R. Aleluia, aleluia, aleluia.", "V. Quem me ama realmente", "guardará a minha palavra."]);
  addSlideFixo(pres, "EVANGELHO", ["João 20,19-23"]);
  addSlidePreto(pres);

  // 9. ORAÇÃO DOS FIÉIS
  addSlideFixo(pres, "Oração dos Fiéis", ["R. Mandai, Senhor,", "o vosso Espírito."]);
  addSlidePreto(pres);

  // 10. OFERTÓRIO
  if (opcoes.ofertorio) {
    const of = banco.ofertorio.find(c => c.titulo === opcoes.ofertorio);
    if (of) renderMusica(pres, of);
  }
  addSlidePreto(pres);

  // 11. SANTO
  if (opcoes.santo) {
    const st = banco.santo.find(c => c.titulo === opcoes.santo);
    if (st) renderMusica(pres, st);
  }
  addSlidePreto(pres);

  // 12. ORAÇÃO EUCARÍSTICA I
  oracaoEucaristica.forEach((resposta, i) => {
    // Tela preta entre resposta 4 e 5
    if (i === 4) addSlidePreto(pres);
    addSlideTexto(pres, resposta, true);
  });
  addSlidePreto(pres);

  // 13. COMUNHÃO
  if (opcoes.comunhao) {
    const com = banco.comunhao.find(c => c.titulo === opcoes.comunhao);
    if (com) renderMusica(pres, com);
  }
  addSlidePreto(pres);

  // 14. CANTO FINAL
  if (opcoes.final) {
    const fin = banco.final.find(c => c.titulo === opcoes.final);
    if (fin) renderMusica(pres, fin);
  }

  return pres;
}
// =============================================
// BANCO DE MÚSICAS — SANTÍSSIMA TRINDADE
// =============================================

const bancoTrindade = {

  cantos_entrada: [
    {
      titulo: "Bendito Sejas Tu",
      autor: "Ir. Custódia Maria Cardoso",
      secoes: [
        { tipo: "Estrofe", texto: "1. Bendito sejas Tu, Senhor de nossos pais\nÉs pródigo de graças, ó Senhor" },
        { tipo: "Refrão", texto: "Glória ao Senhor\nCriador para sempre\nGlória ao Senhor\nCriador para sempre (Bis)" },
        { tipo: "Estrofe", texto: "2. Bendito sejas Tu, ó Verbo de Deus Pai\nA morte que sofreste nos deu vida" },
        { tipo: "Refrão", texto: "Glória ao Senhor\nCriador para sempre\nGlória ao Senhor\nCriador para sempre (Bis)" },
        { tipo: "Estrofe", texto: "3. Bendito sejas Tu, Espírito de Deus\nOperas na igreja a salvação" },
        { tipo: "Refrão", texto: "Glória ao Senhor\nCriador para sempre\nGlória ao Senhor\nCriador para sempre (Bis)" },
      ]
    },
    {
      titulo: "Vem Santíssima Trindade",
      autor: "José Acácio Santana",
      secoes: [
        { tipo: "Refrão", texto: "Vem Santíssima Trindade, Pai e Filho e Santo Amor\nRecriar a humanidade, renovar o seu valor" },
        { tipo: "Estrofe", texto: "1. Conheces a palavra e o pensamento\nPenetras coração e sentimento\nNão posso me afastar da tua frente\nAonde eu me esconder estás presente." },
        { tipo: "Refrão", texto: "Vem Santíssima Trindade, Pai e Filho e Santo Amor\nRecriar a humanidade, renovar o seu valor" },
        { tipo: "Estrofe", texto: "2. Envolves o meu ser por todo lado\nJamais vou me sentir abandonado\nO teu saber me encanta e me supera\nTu és a minha eterna primavera." },
        { tipo: "Refrão", texto: "Vem Santíssima Trindade, Pai e Filho e Santo Amor\nRecriar a humanidade, renovar o seu valor" },
        { tipo: "Estrofe", texto: "3. Se um dia me envolver à escuridão\nSerás a minha grande proteção\nContigo em tua santa companhia\nA noite é sempre clara como o dia." },
        { tipo: "Refrão", texto: "Vem Santíssima Trindade, Pai e Filho e Santo Amor\nRecriar a humanidade, renovar o seu valor" },
        { tipo: "Estrofe", texto: "4. As fibras do meu ser entrelaçaste\nNo seio maternal tu me formaste\nEu quero te louvar como ninguém\nAgora e pelos séculos, amém." },
        { tipo: "Refrão", texto: "Vem Santíssima Trindade, Pai e Filho e Santo Amor\nRecriar a humanidade, renovar o seu valor" },
      ]
    },
    {
      titulo: "Entremos com Grande Alegria",
      autor: "",
      secoes: [
        { tipo: "Estrofe", texto: "1. Entremos com grande alegria\nna casa do Senhor.\nEm sua fiel companhia\ncantemos seu louvor." },
        { tipo: "Refrão", texto: "Honra e glória à Santíssima Trindade.\nHonra e glória por toda a eternidade!\nHonra e glória à Trindade Santa!" },
        { tipo: "Estrofe", texto: "2. Aqui todos juntos oramos\ncom fé e gratidão.\nE a bênção de Deus invocamos\nde todo o coração." },
        { tipo: "Refrão", texto: "Honra e glória à Santíssima Trindade.\nHonra e glória por toda a eternidade!\nHonra e glória à Trindade Santa!" },
        { tipo: "Estrofe", texto: "3. Em nome do Pai sacrossanto\ndo Filho Salvador\nno amor do Espírito Santo\nvivemos sem temor." },
        { tipo: "Refrão", texto: "Honra e glória à Santíssima Trindade.\nHonra e glória por toda a eternidade!\nHonra e glória à Trindade Santa!" },
      ]
    },
    {
      titulo: "Trindade Bendita",
      autor: "Marcelo Oliveira",
      secoes: [
        { tipo: "Estrofe", texto: "1. Ó Trindade, num trono supremo\nque brilhais, num intenso fulgor.\nGlória a vós, que o profundo dos seres\npossuís e habitais pelo amor." },
        { tipo: "Refrão", texto: "Bendito seja Deus Pai,\nbendito o Filho Unigênito,\nbendito o Espírito Santo.\nMisericordioso foi Deus para conosco,\npara conosco." },
        { tipo: "Estrofe", texto: "2. Ó Deus Pai, Criador do universo,\nSois a força que a todos dá vida;\naos que dela fizestes consortes,\ndai a fé, que sustenta na lida." },
        { tipo: "Refrão", texto: "Bendito seja Deus Pai,\nbendito o Filho Unigênito,\nbendito o Espírito Santo.\nMisericordioso foi Deus para conosco,\npara conosco." },
        { tipo: "Estrofe", texto: "3. Esplendor e espelho da luz\nSois, ó Filho, que irmãos nos chamais;\ndai-nos ser ramos verdes e vivos\nda fecunda videira do Pai." },
        { tipo: "Refrão", texto: "Bendito seja Deus Pai,\nbendito o Filho Unigênito,\nbendito o Espírito Santo.\nMisericordioso foi Deus para conosco,\npara conosco." },
      ]
    },
    {
      titulo: "Bom É Poder Estar Aqui",
      autor: "",
      secoes: [
        { tipo: "Estrofe", texto: "1. Bom é poder estar aqui\nAmando, adorando e exaltando\na Santíssima Trindade que estará\naqui para fazer santo o lugar" },
        { tipo: "Refrão", texto: "Em nome do Pai, em nome do Filho\nem nome do Espírito Santo, amém\nLouvemos e adoremos ao nosso Deus vivo\nEm nome do Pai, em nome do Filho\nem nome do Espírito Santo, amém\nPodes entrar Senhor, neste lugar" },
      ]
    },
    {
      titulo: "Luz que Vem do Alto",
      autor: "",
      secoes: [
        { tipo: "Refrão", texto: "Luz que vem do alto.\nLuz que traz a vida.\nVem brilhar em nós, ó Luz Divina!" },
        { tipo: "Estrofe", texto: "1. Ó Pai Santo, teu amor criou o mundo,\nnós cantamos teu mistério Criador." },
        { tipo: "Refrão", texto: "Luz que vem do alto.\nLuz que traz a vida.\nVem brilhar em nós, ó Luz Divina!" },
        { tipo: "Estrofe", texto: "2. Filho Amado, és o Verbo que redime,\nnós cantamos teu mistério Redentor." },
        { tipo: "Refrão", texto: "Luz que vem do alto.\nLuz que traz a vida.\nVem brilhar em nós, ó Luz Divina!" },
        { tipo: "Estrofe", texto: "3. Ó Divino, Defensor da humanidade,\nnós cantamos teu mistério de amor." },
        { tipo: "Refrão", texto: "Luz que vem do alto.\nLuz que traz a vida.\nVem brilhar em nós, ó Luz Divina!" },
      ]
    },
  ],

  aspersao: [
    {
      titulo: "Banhados em Cristo",
      autor: "",
      secoes: [
        { tipo: "Refrão", texto: "Banhados em Cristo,\nsomos uma nova criatura.\nAs coisas antigas já se passaram,\nsomos nascidos de novo.\nAleluia, aleluia, aleluia!" },
        { tipo: "Estrofe", texto: "1. Lavados na fonte viva\ndo lado aberto de Cristo,\ntranspomos vitoriosos\nas portas do paraíso!" },
        { tipo: "Refrão", texto: "Banhados em Cristo,\nsomos uma nova criatura.\nAs coisas antigas já se passaram,\nsomos nascidos de novo.\nAleluia, aleluia, aleluia!" },
      ]
    },
    {
      titulo: "Derramarei Sobre Vós",
      autor: "",
      secoes: [
        { tipo: "Refrão", texto: "Derramarei sobre vós uma água pura,\nsereis purificados de todas as faltas,\neu vos darei um coração novo, diz o Senhor." },
        { tipo: "Estrofe", texto: "1. Tende piedade de mim, ó Deus,\nsegundo a vossa grande misericórdia,\npor vossa bondade imensa,\napagai minha iniquidade." },
        { tipo: "Refrão", texto: "Derramarei sobre vós uma água pura,\nsereis purificados de todas as faltas,\neu vos darei um coração novo, diz o Senhor." },
        { tipo: "Estrofe", texto: "2. Lavai-me todo inteiro de minha culpa\ne do meu pecado purificai-me,\nsim, reconheço minha maldade,\nà minha frente está sempre o meu pecado." },
        { tipo: "Refrão", texto: "Derramarei sobre vós uma água pura,\nsereis purificados de todas as faltas,\neu vos darei um coração novo, diz o Senhor." },
        { tipo: "Estrofe", texto: "3. Dos meus pecados desviai a vossa face\ne todas as minhas culpas apagai,\ncriai em mim um coração que seja puro, meu Deus,\nponde em mim um espírito resoluto." },
        { tipo: "Refrão", texto: "Derramarei sobre vós uma água pura,\nsereis purificados de todas as faltas,\neu vos darei um coração novo, diz o Senhor." },
      ]
    },
  ],

  penitencial: [
    {
      titulo: "Senhor, que pela Água e pelo Espírito",
      autor: "",
      secoes: [
        { tipo: "Estrofe", texto: "Senhor, que pela água e pelo Espírito Santo\nnos regenerastes à vossa imagem,\ntende piedade de nós!\nSenhor, tende piedade de nós! Piedade de nós!" },
        { tipo: "Estrofe", texto: "Cristo, que enviais o vosso Espírito\npara criar dentro de nós um coração novo,\ntende piedade de nós!\nCristo, tende piedade de nós! Piedade de nós!" },
        { tipo: "Estrofe", texto: "Senhor, que nos tornastes participantes\ndo Vosso Corpo e do Vosso Sangue,\ntende piedade de nós!\nSenhor, tende piedade de nós! Piedade de nós!" },
      ]
    },
    {
      titulo: "Senhor, que Subindo ao Céu",
      autor: "",
      secoes: [
        { tipo: "Estrofe", texto: "1. Senhor, que subindo ao céu\nnos presenteastes com o dom do Espírito,\ntende piedade de nós.\nKyrie eleison! Kyrie eleison! (Bis)" },
        { tipo: "Estrofe", texto: "2. Cristo, que dais vida a todas as coisas\ncom o poder da vossa palavra,\ntende piedade de nós.\nChriste eleison! Christe eleison! (Bis)" },
        { tipo: "Estrofe", texto: "3. Senhor, Rei do universo e Senhor dos séculos,\ntende piedade de nós.\nKyrie eleison! Kyrie eleison! (Bis)" },
      ]
    },
    {
      titulo: "Sacerdote da Nova Aliança",
      autor: "",
      secoes: [
        { tipo: "Estrofe", texto: "1. Senhor, que sois o eterno sacerdote\nda nova Aliança,\ntende piedade de nós.\nSenhor, tende piedade de nós. (Bis)" },
        { tipo: "Estrofe", texto: "2. Cristo, que nos edificais como pedras vivas\nno templo santo de Deus,\ntende piedade de nós.\nCristo, tende piedade de nós. (Bis)" },
        { tipo: "Estrofe", texto: "3. Senhor, que nos tornais concidadãos dos santos\nno reino dos céus,\ntende piedade de nós.\nSenhor, tende piedade de nós. (Bis)" },
      ]
    },
    {
      titulo: "Senhor, Senhor, Tende Piedade de Nós",
      autor: "",
      secoes: [
        { tipo: "Estrofe", texto: "1. Senhor, nossa paz,\ntende piedade de nós.\nSenhor, Senhor, tende piedade de nós. (Bis)" },
        { tipo: "Estrofe", texto: "2. Cristo, nossa Páscoa,\ntende piedade de nós.\nCristo, Cristo, tende piedade de nós. (Bis)" },
        { tipo: "Estrofe", texto: "3. Senhor, nossa vida,\ntende piedade de nós.\nSenhor, Senhor, tende piedade de nós. (Bis)" },
      ]
    },
  ],

  gloria: [
    { titulo: "Glória (Vinde Cristãos)", autor: "", secoes: [
      { tipo: "Refrão", texto: "Glória a Deus nas alturas! (Bis)" },
      { tipo: "Estrofe", texto: "1. Glória a Deus nos altos céus,\npaz na terra aos seus amados.\nA vós louvam, Rei celeste,\nos que foram libertados." },
      { tipo: "Refrão", texto: "Glória a Deus nas alturas! (Bis)" },
      { tipo: "Estrofe", texto: "2. Deus e Pai nós vos louvamos,\nadoramos, bendizemos.\nDamos glória ao vosso nome,\nvossos dons agradecemos." },
      { tipo: "Refrão", texto: "Glória a Deus nas alturas! (Bis)" },
      { tipo: "Estrofe", texto: "3. Senhor nosso Jesus Cristo,\nunigênito do Pai.\nVós de Deus Cordeiro Santo,\nnossas culpas perdoai." },
      { tipo: "Refrão", texto: "Glória a Deus nas alturas! (Bis)" },
      { tipo: "Estrofe", texto: "4. Vós que estais junto do Pai,\ncomo nosso intercessor.\nAcolhei nossos pedidos,\natendei nosso clamor." },
      { tipo: "Refrão", texto: "Glória a Deus nas alturas! (Bis)" },
      { tipo: "Estrofe", texto: "5. Vós somente sois o Santo,\no Altíssimo, o Senhor,\ncom o Espírito Divino,\nde Deus Pai no esplendor." },
      { tipo: "Refrão", texto: "Glória a Deus nas alturas! (Bis)" },
    ]},
    { titulo: "Glória, Glória, Glória a Deus", autor: "", secoes: [
      { tipo: "Refrão", texto: "Glória, glória, glória a Deus\nnas alturas e na terra paz aos homens!" },
      { tipo: "Estrofe", texto: "1. Senhor Deus, Rei dos céus,\nDeus Pai todo-poderoso.\nNós vos louvamos, nós vos bendizemos,\nnós vos adoramos, nós vos glorificamos.\nNós vos damos graças por vossa imensa glória." },
      { tipo: "Refrão", texto: "Glória, glória, glória a Deus\nnas alturas e na terra paz aos homens!" },
      { tipo: "Estrofe", texto: "2. Senhor Jesus Cristo, Filho Unigênito,\nSenhor Deus, Cordeiro de Deus, Filho de Deus Pai.\nVós que tirais o pecado do mundo,\ntende piedade de nós.\nVós que estais à direita do Pai,\ntende piedade de nós." },
      { tipo: "Refrão", texto: "Glória, glória, glória a Deus\nnas alturas e na terra paz aos homens!" },
      { tipo: "Estrofe", texto: "3. Só vós sois Santo, só vós o Senhor,\nsó vós o Altíssimo, Jesus Cristo!\nCom o Espírito Santo,\nna glória de Deus Pai. Amém." },
      { tipo: "Refrão", texto: "Glória, glória, glória a Deus\nnas alturas e na terra paz aos homens!" },
    ]},
    { titulo: "Glória a Deus nas Alturas", autor: "Casemiro Nogueira", secoes: [
      { tipo: "Refrão", texto: "Glória a Deus nas alturas,\nglória a Deus nas alturas,\ne paz na terra aos homens por Ele amados." },
      { tipo: "Estrofe", texto: "1. Senhor Deus, Rei dos céus, Deus Pai todo-poderoso.\nNós vos louvamos, nós vos bendizemos,\nnós vos adoramos, nós vos glorificamos,\nnós vos damos graças por vossa imensa glória." },
      { tipo: "Refrão", texto: "Glória a Deus nas alturas,\nglória a Deus nas alturas,\ne paz na terra aos homens por Ele amados." },
      { tipo: "Estrofe", texto: "2. Só vós sois Santo.\nSó vós o Senhor,\nsó vós o Altíssimo, Jesus Cristo!\nCom o Espírito Santo,\nna glória de Deus Pai. Amém!" },
      { tipo: "Refrão", texto: "Glória a Deus nas alturas,\nglória a Deus nas alturas,\ne paz na terra aos homens por Ele amados." },
    ]},
    { titulo: "Glória a Deus nos Altos Céus", autor: "", secoes: [
      { tipo: "Refrão", texto: "Glória a Deus lá nos céus\ne paz aos seus, amém. (Bis)" },
      { tipo: "Estrofe", texto: "1. Glória a Deus nos altos céus,\npaz na terra a seus amados.\nA vós louvam, Rei celeste,\nos que foram libertados." },
      { tipo: "Refrão", texto: "Glória a Deus lá nos céus\ne paz aos seus, amém. (Bis)" },
      { tipo: "Estrofe", texto: "2. Deus e Pai nós vos louvamos,\nadoramos, bendizemos.\nDamos glória ao vosso nome,\nvossos dons agradecemos." },
      { tipo: "Refrão", texto: "Glória a Deus lá nos céus\ne paz aos seus, amém. (Bis)" },
      { tipo: "Estrofe", texto: "3. Vós somente sois o Santo,\no Altíssimo, o Senhor,\ncom o Espírito Divino,\nde Deus Pai no esplendor!" },
      { tipo: "Refrão", texto: "Glória a Deus lá nos céus\ne paz aos seus, amém. (Bis)" },
    ]},
    { titulo: "Glória a Deus lá nos Céus, e Paz na Terra aos Seus!", autor: "", secoes: [
      { tipo: "Refrão", texto: "Glória a Deus lá nos céus,\ne paz na terra aos seus! (Bis)" },
      { tipo: "Estrofe", texto: "1. Glória a Deus nos altos céus,\npaz na terra a seus amados,\na vós louvam, Rei celeste,\nos que foram libertados!" },
      { tipo: "Refrão", texto: "Glória a Deus lá nos céus,\ne paz na terra aos seus! (Bis)" },
      { tipo: "Estrofe", texto: "2. Vós somente sois o Santo,\no Altíssimo, o Senhor,\ncom o Espírito Divino,\nde Deus Pai no esplendor!" },
      { tipo: "Refrão", texto: "Glória a Deus lá nos céus,\ne paz na terra aos seus! (Bis)" },
    ]},
    { titulo: "Glória, Glória a Deus nas Alturas", autor: "", secoes: [
      { tipo: "Refrão", texto: "Glória, glória, a Deus nas alturas. (Bis)" },
      { tipo: "Estrofe", texto: "1. Glória a Deus nas alturas\ne paz na terra aos homens por Ele amados.\nSenhor Deus, Rei dos céus, Deus Pai todo-poderoso.\nNós vos louvamos, nós vos bendizemos,\nnós vos adoramos, nós vos glorificamos.\nNós vos damos graças por vossa imensa glória." },
      { tipo: "Refrão", texto: "Glória, glória, a Deus nas alturas. (Bis)" },
      { tipo: "Estrofe", texto: "2. Só vós sois Santo. Só vós o Senhor,\nsó vós o Altíssimo, Jesus Cristo.\nCom o Espírito Santo,\nna glória de Deus Pai, amém!" },
      { tipo: "Refrão", texto: "Glória, glória, a Deus nas alturas. (Bis)" },
    ]},
    { titulo: "Glória Shalom", autor: "", secoes: [
      { tipo: "Refrão", texto: "Glória, glória a Deus nas alturas,\nô, ô glória e a nós a Sua paz. (Bis)" },
      { tipo: "Estrofe", texto: "1. Senhor Deus, Rei dos céus, Deus Pai onipotente,\nvos louvamos, bendizemos, adoramos.\nNós vos glorificamos\ne nós vos damos graças por vossa imensa glória." },
      { tipo: "Refrão", texto: "Glória, glória a Deus nas alturas,\nô, ô glória e a nós a Sua paz. (Bis)" },
      { tipo: "Estrofe", texto: "2. Só vós sois o Santo, o Senhor,\no Altíssimo, só vós,\nJesus Cristo, com o Espírito e o Pai." },
      { tipo: "Refrão", texto: "Glória, glória a Deus nas alturas,\nô, ô glória e a nós a Sua paz. (Bis)" },
    ]},
    { titulo: "Ô Ô Glória", autor: "", secoes: [
      { tipo: "Refrão", texto: "Ô ô glória, ô glória!\nA Deus Pai, a Jesus Cristo\ne ao Espírito de amor. (Bis)" },
      { tipo: "Estrofe", texto: "1. Deus e Pai nós vos louvamos,\nadoramos, bendizemos.\nDamos glória a vosso nome,\nvossos dons agradecemos." },
      { tipo: "Refrão", texto: "Ô ô glória, ô glória!\nA Deus Pai, a Jesus Cristo\ne ao Espírito de amor. (Bis)" },
      { tipo: "Estrofe", texto: "2. Vós somente sois o Santo,\no Altíssimo, o Senhor,\ncom o Espírito Divino,\nde Deus Pai no esplendor." },
      { tipo: "Refrão", texto: "Ô ô glória, ô glória!\nA Deus Pai, a Jesus Cristo\ne ao Espírito de amor. (Bis)" },
    ]},
    { titulo: "Glória, Glória, Anjos do Céu", autor: "", secoes: [
      { tipo: "Refrão", texto: "Glória, glória, anjos do céu,\ncantam todos Seu amor.\nE na terra, homens de paz,\nDeus merece o louvor. (Bis)" },
      { tipo: "Estrofe", texto: "1. Deus e Pai, nós vos louvamos,\nadoramos, bendizemos.\nDamos glória ao vosso nome,\nvossos dons agradecemos." },
      { tipo: "Refrão", texto: "Glória, glória, anjos do céu,\ncantam todos Seu amor.\nE na terra, homens de paz,\nDeus merece o louvor. (Bis)" },
      { tipo: "Estrofe", texto: "2. Vós somente sois o Santo,\no Altíssimo, o Senhor,\ncom o Espírito Divino,\nde Deus Pai no esplendor." },
      { tipo: "Refrão", texto: "Glória, glória, anjos do céu,\ncantam todos Seu amor.\nE na terra, homens de paz,\nDeus merece o louvor. (Bis)" },
    ]},
    { titulo: "Glória a Deus nas Alturas", autor: "Pe. Ney Brasil", secoes: [
      { tipo: "Refrão", texto: "Glória a Deus nas alturas,\ne paz na terra aos homens por Ele amados. (Bis)" },
      { tipo: "Estrofe", texto: "1. Senhor Deus, Rei dos céus, Deus Pai todo-poderoso:\nnós vos louvamos, nós vos bendizemos,\nnós vos adoramos, nós vos glorificamos." },
      { tipo: "Refrão", texto: "Glória a Deus nas alturas,\ne paz na terra aos homens por Ele amados. (Bis)" },
      { tipo: "Estrofe", texto: "2. Só vós sois o Santo, só vós, o Senhor,\nsó vós, o Altíssimo, Jesus Cristo,\ncom o Espírito Santo, na glória de Deus Pai." },
      { tipo: "Refrão", texto: "Glória a Deus nas alturas,\ne paz na terra aos homens por Ele amados. (Bis)" },
    ]},
  ],

  aclamacao: [
    { titulo: "Aleluia! Glória ao Pai e ao Filho (1)", autor: "", secoes: [
      { tipo: "Refrão", texto: "Aleluia! Aleluia! Aleluia! (Bis)" },
      { tipo: "Estrofe", texto: "Glória ao Pai e ao Filho e ao Espírito Divino,\nao Deus que é, que era e que vem,\npelos séculos. Amém!" },
      { tipo: "Refrão", texto: "Aleluia! Aleluia! Aleluia! (Bis)" },
    ]},
    { titulo: "Aleluia! Glória ao Pai e ao Filho (2)", autor: "", secoes: [
      { tipo: "Refrão", texto: "Aleluia! Aleluia! Aleluia! Aleluia! (Bis)" },
      { tipo: "Estrofe", texto: "Glória ao Pai e ao Filho e ao Espírito Divino,\nao Deus que é, que era e que vem,\npelos séculos. Amém!" },
      { tipo: "Refrão", texto: "Aleluia! Aleluia! Aleluia! Aleluia! (Bis)" },
    ]},
    { titulo: "Aleluia! Glória ao Pai e ao Filho (3)", autor: "", secoes: [
      { tipo: "Refrão", texto: "Aleluia! Aleluia! Aleluia! Aleluia! (Bis)" },
      { tipo: "Estrofe", texto: "Glória ao Pai e ao Filho e ao Espírito Divino,\nao Deus que é, que era e que vem,\npelos séculos. Amém!" },
      { tipo: "Refrão", texto: "Aleluia! Aleluia! Aleluia! Aleluia! (Bis)" },
    ]},
  ],

  ofertorio: [
    { titulo: "Bendito e Louvado Seja", autor: "Ir. Custódia Maria Cardoso", secoes: [
      { tipo: "Estrofe", texto: "1. Bendito e louvado seja\no Pai nosso Criador.\nO pão que nos recebemos\né prova do seu amor.\nNa missa é transformado\nno corpo do Salvador." },
      { tipo: "Refrão", texto: "Bendito seja Deus,\nbendito seu amor.\nBendito seja Deus\nPai onipotente nosso Criador. (Bis)" },
      { tipo: "Estrofe", texto: "2. Bendito e louvado seja\no Pai nosso Criador.\nO vinho que recebemos,\né prova do seu amor.\nNa missa é transformado\nno sangue do Salvador." },
      { tipo: "Refrão", texto: "Bendito seja Deus,\nbendito seu amor.\nBendito seja Deus\nPai onipotente nosso Criador. (Bis)" },
    ]},
    { titulo: "Ó Trindade Imensa e Una", autor: "Ir. Míria T. Kolling", secoes: [
      { tipo: "Estrofe", texto: "1. Ó Trindade imensa e una,\nvossa força tudo cria;\nvossa mão que rege os tempos,\nantes deles existia." },
      { tipo: "Estrofe", texto: "2. Pai, da graça fonte viva,\nLuz da glória de Deus Pai,\nSanto Espírito da vida,\nque no amor os enlaçais." },
      { tipo: "Estrofe", texto: "3. Só por vós, Trindade Santa,\nsuma origem, todo bem,\ntodo ser, toda beleza,\ntoda vida se mantém." },
      { tipo: "Estrofe", texto: "4. Nós, os filhos adotivos,\npela graça consagrados,\nnos tornemos templos vivos,\na vós sempre dedicados." },
    ]},
    { titulo: "A Vós, Senhor, Apresentamos", autor: "Ir. Custódia Maria Cardoso", secoes: [
      { tipo: "Refrão", texto: "A vós, Senhor\napresentamos estes dons\no pão e o vinho, aleluia!" },
      { tipo: "Estrofe", texto: "1. Que poderei retribuir ao Senhor Deus\npor tudo aquilo que ele fez em meu favor?" },
      { tipo: "Refrão", texto: "A vós, Senhor\napresentamos estes dons\no pão e o vinho, aleluia!" },
      { tipo: "Estrofe", texto: "2. Elevo o cálice da minha salvação\ninvocando o nome santo do Senhor." },
      { tipo: "Refrão", texto: "A vós, Senhor\napresentamos estes dons\no pão e o vinho, aleluia!" },
      { tipo: "Estrofe", texto: "3. Vou cumprir minhas promessas ao Senhor\nna presença de seu povo reunido." },
      { tipo: "Refrão", texto: "A vós, Senhor\napresentamos estes dons\no pão e o vinho, aleluia!" },
      { tipo: "Estrofe", texto: "4. Por isso oferto um sacrifício de louvor\ninvocando o nome santo do Senhor." },
      { tipo: "Refrão", texto: "A vós, Senhor\napresentamos estes dons\no pão e o vinho, aleluia!" },
    ]},
    { titulo: "Nossa Oferta na Trindade", autor: "", secoes: [
      { tipo: "Estrofe", texto: "1. Quando esse tudo era nada\nsó três pessoas haviam\nnuma comunicação de amor\nem total e perfeita harmonia.\nUma doou-se na criação\na outra doou-se na encarnação\ne a terceira se doa e trabalha na igreja\npor nossa santificação." },
      { tipo: "Refrão", texto: "Nossa oferta, na Trindade Santa\nde nossas vidas, talentos e bens\napresentemos, pra que as consagrem\nPai, Filho e Espírito Santo, amém. (Bis)" },
      { tipo: "Estrofe", texto: "2. Temos que ser solidários\nsaber nossos bens partilhar\ntal qual o amor trinitário\nque se doa no comunicar.\nO Pai doou-se no Filho\no Filho se deu a nós por amor\ne ao Espírito Santo no Pai e no Filho\nprestemos o nosso louvor." },
      { tipo: "Refrão", texto: "Nossa oferta, na Trindade Santa\nde nossas vidas, talentos e bens\nofereçamos em nome do Pai\ndo Filho e do Espírito Santo, amém. (Bis)" },
    ]},
  ],

  santo: [
    { titulo: "Santo III", secoes: [
      { tipo: "Única", texto: "Santo, Santo, Santo.\nSenhor Deus do universo.\nO céu e a terra\nproclam a vossa glória.\nHosana nas alturas." },
      { tipo: "Única", texto: "Bendito o que vem\nem nome do Senhor.\nHosana, nas alturas." },
    ]},
    { titulo: "Santo - Hosana nas Alturas", secoes: [
      { tipo: "Única", texto: "Santo, Santo, Santo.\nSenhor Deus do universo.\nO céu e a terra proclamam\na vossa glória.\nHosana nas alturas. Hosana! (Bis)" },
      { tipo: "Única", texto: "Bendito Aquele que vem\nem nome do Senhor. (Bis)" },
    ]},
    { titulo: "Santo (I) - Pe. Ney Brasil", secoes: [
      { tipo: "Única", texto: "Santo, Santo, Santo sois Senhor,\nSanto, Santo, Santo sois Senhor Nosso Deus.\nSenhor Deus do universo.\nO céu e a terra proclamam a vossa glória.\nHosana nas alturas." },
      { tipo: "Única", texto: "Santo, Santo, Santo Sois Senhor,\nSanto, Santo, Santo Sois Senhor Nosso Deus.\nBendito o que vem em nome do Senhor.\nHosana nas alturas, hosana nas alturas." },
    ]},
    { titulo: "Santo - Pedro Ivan", secoes: [
      { tipo: "Única", texto: "Santo, Santo, Santo.\nSenhor Deus do universo.\nO céu e a terra\nproclam a vossa glória.\nHosana nas alturas." },
      { tipo: "Única", texto: "Bendito o que vem\nem nome do Senhor.\nHosana, nas alturas." },
    ]},
    { titulo: "O Senhor é Santo", secoes: [
      { tipo: "Única", texto: "O Senhor é Santo!\nO Senhor é Santo!\nO Senhor é Santo!\nSenhor Deus do universo.\nO céu e a terra\nproclam a vossa glória.\nHosana nas alturas." },
      { tipo: "Única", texto: "O Senhor é Santo!\nO Senhor é Santo!\nO Senhor é Santo!\nBendito o que vem em nome do Senhor.\nHosana! Hosana! Hosana! Hosana!" },
    ]},
    { titulo: "Santo - Padre Cleidimar Moreira", secoes: [
      { tipo: "Única", texto: "Santo, Santo, Santo é o Senhor.\nDeus do universo. (Bis)\nOs céus e a terra\nproclam a vossa glória.\nHosana nas alturas.\nHosana nas alturas. (Bis)" },
      { tipo: "Única", texto: "Santo, Santo, Santo é o Senhor.\nDeus do universo. (Bis)\nBendito o que vem\nem nome do Senhor.\nHosana, nas alturas.\nHosana nas alturas. (Bis)" },
    ]},
    { titulo: "Santo, Santo, Santo é o Senhor", secoes: [
      { tipo: "Única", texto: "Santo, Santo, Santo é o Senhor,\nDeus do universo, do céu e a terra. (Bis)\nHosana! Hosana!\nHosana! nas alturas. (Bis)" },
      { tipo: "Única", texto: "Bendito o que vem\nem nome do Senhor.\nHosana, nas alturas. (Bis)" },
    ]},
  ],

  cordeiro: [
    { titulo: "Cordeiro de Deus (1)", secoes: [
      { tipo: "Única", texto: "Cordeiro de Deus,\nque tirais o pecado do mundo,\ntende piedade de nós.\nCordeiro de Deus,\nque tirais o pecado do mundo,\ntende piedade de nós.\nCordeiro de Deus,\nque tirais o pecado do mundo,\ndai-nos a paz." },
    ]},
    { titulo: "Cordeiro de Deus (2)", secoes: [
      { tipo: "Única", texto: "Cordeiro de Deus,\nque tirais o pecado do mundo,\ntende piedade de nós.\nCordeiro de Deus,\nque tirais o pecado do mundo,\ntende piedade de nós.\nCordeiro de Deus,\nque tirais o pecado do mundo,\ndai-nos a paz, dai-nos a paz." },
    ]},
    { titulo: "Cordeiro de Deus (3)", secoes: [
      { tipo: "Única", texto: "Cordeiro de Deus,\nque tirais o pecado do mundo,\ntende piedade.\nCordeiro de Deus,\nque tirais o pecado do mundo,\ntende piedade.\nCordeiro de Deus,\nque tirais o pecado do mundo,\ndai-nos a paz, dai-nos a paz,\ndai-nos a vossa paz, dai-nos a paz. (Bis)" },
    ]},
    { titulo: "Cordeiro de Deus (4)", secoes: [
      { tipo: "Única", texto: "Cordeiro de Deus,\nque tirais o pecado do mundo,\ntende piedade de nós.\nCordeiro de Deus,\nque tirais o pecado do mundo,\ntende piedade de nós.\nCordeiro de Deus,\nque tirais o pecado do mundo,\ndai-nos a paz." },
    ]},
    { titulo: "Cordeiro de Deus (5)", secoes: [
      { tipo: "Única", texto: "Cordeiro de Deus,\nque tirais o pecado do mundo,\ntende piedade de nós.\nCordeiro de Deus,\nque tirais o pecado do mundo,\ntende piedade de nós.\nCordeiro de Deus,\nque tirais o pecado do mundo,\ndai-nos a paz, dai-nos a paz,\ndai-nos a vossa paz." },
    ]},
    { titulo: "Cordeiro de Deus (6)", secoes: [
      { tipo: "Única", texto: "Cordeiro de Deus,\nque tirais o pecado do mundo,\ntende piedade, piedade de nós.\nCordeiro de Deus,\nque tirais o pecado do mundo,\ntende piedade, piedade de nós.\nCordeiro de Deus,\nque tirais o pecado do mundo,\ndai-nos a paz, a vossa paz." },
    ]},
    { titulo: "Cordeiro de Deus (7)", secoes: [
      { tipo: "Única", texto: "Cordeiro de Deus,\nque tirais o pecado do mundo,\ntende piedade, tende piedade,\npiedade de nós.\nCordeiro de Deus,\nque tirais o pecado do mundo,\ntende piedade, tende piedade,\npiedade de nós.\nCordeiro de Deus,\nque tirais o pecado do mundo,\ndai-nos a paz." },
    ]},
  ],

  comunhao: [
    { titulo: "Porque Sois Filhos", autor: "Fr. Wanderson Luiz Freitas", secoes: [
      { tipo: "Refrão", texto: "Porque sois filhos, Deus enviou\naos vossos corações o Espírito do seu Filho,\nque clama: Abá, ó Pai! (Bis)" },
      { tipo: "Estrofe", texto: "1. Bendize, ó minh'alma, ao Senhor,\ne todo o meu ser seu santo nome!" },
      { tipo: "Refrão", texto: "Porque sois filhos, Deus enviou\naos vossos corações o Espírito do seu Filho,\nque clama: Abá, ó Pai! (Bis)" },
      { tipo: "Estrofe", texto: "2. Pois ele te perdoa toda culpa\ne cura toda a tua enfermidade;" },
      { tipo: "Refrão", texto: "Porque sois filhos, Deus enviou\naos vossos corações o Espírito do seu Filho,\nque clama: Abá, ó Pai! (Bis)" },
      { tipo: "Estrofe", texto: "3. Da sepultura ele salva a tua vida\ne te cerca de carinho e compaixão." },
      { tipo: "Refrão", texto: "Porque sois filhos, Deus enviou\naos vossos corações o Espírito do seu Filho,\nque clama: Abá, ó Pai! (Bis)" },
      { tipo: "Estrofe", texto: "4. O Senhor é indulgente, é favorável,\né paciente, é bondoso e compassivo." },
      { tipo: "Refrão", texto: "Porque sois filhos, Deus enviou\naos vossos corações o Espírito do seu Filho,\nque clama: Abá, ó Pai! (Bis)" },
      { tipo: "Estrofe", texto: "5. O amor do Senhor Deus por quem o teme\né de sempre e perdura para sempre." },
      { tipo: "Refrão", texto: "Porque sois filhos, Deus enviou\naos vossos corações o Espírito do seu Filho,\nque clama: Abá, ó Pai! (Bis)" },
    ]},
    { titulo: "Glória ao Pai e ao Filho e ao Santo Espírito", autor: "", secoes: [
      { tipo: "Refrão", texto: "Glória ao Pai, e ao Filho e ao Santo Espírito,\nao Deus que é, que era e que vem,\npelos séculos dos séculos. Amém." },
      { tipo: "Estrofe", texto: "1. Ó justos, alegrai-vos no Senhor!\nAos retos fica bem glorificá-lo.\nDai graças ao Senhor ao som da harpa,\nna lira de dez cordas celebrai-o!" },
      { tipo: "Refrão", texto: "Glória ao Pai, e ao Filho e ao Santo Espírito,\nao Deus que é, que era e que vem,\npelos séculos dos séculos. Amém." },
      { tipo: "Estrofe", texto: "2. Pois reta é a palavra do Senhor,\ne tudo o que ele faz merece fé.\nDeus ama o direito e a justiça,\ntransborda em toda a terra a sua graça." },
      { tipo: "Refrão", texto: "Glória ao Pai, e ao Filho e ao Santo Espírito,\nao Deus que é, que era e que vem,\npelos séculos dos séculos. Amém." },
    ]},
    { titulo: "Teu Amor Vai Além da Medida", autor: "José Tomaz Filho", secoes: [
      { tipo: "Estrofe", texto: "1. Teu amor vai além da medida,\nse a medida é o meu próprio pensar.\nO teu sonho é partilha e convida\ntodo ser a saber partilhar." },
      { tipo: "Refrão", texto: "Teu amor é de Pai e de Filho,\nsem limite, é de eterno vigor;\né de Espírito Santo teu brilho,\né total Comunhão teu Amor!" },
      { tipo: "Estrofe", texto: "2. Teu amor vai além da medida,\nse a medida é o que posso fazer.\nO universo confirma que a vida\né o sublime destino do ser!" },
      { tipo: "Refrão", texto: "Teu amor é de Pai e de Filho,\nsem limite, é de eterno vigor;\né de Espírito Santo teu brilho,\né total Comunhão teu Amor!" },
      { tipo: "Estrofe", texto: "3. Teu amor vai além da medida,\nse a medida é o que posso dizer.\nMinha voz é tão frágil, partida,\nsó tua voz é que ensina a viver!" },
      { tipo: "Refrão", texto: "Teu amor é de Pai e de Filho,\nsem limite, é de eterno vigor;\né de Espírito Santo teu brilho,\né total Comunhão teu Amor!" },
    ]},
    { titulo: "Deus Eterno, a Vós Louvor", autor: "Grosser Gott", secoes: [
      { tipo: "Estrofe", texto: "1. Deus eterno a vós louvor!\nGlória a vossa majestade!\nAnjos e homens com fervor,\nvos adoram, Deus Trindade." },
      { tipo: "Refrão", texto: "Cante a terra com amor!\nSanto, Santo é o Senhor. (Bis)" },
      { tipo: "Estrofe", texto: "2. Pai eterno, a criação\nque tirastes vós do nada,\nrepousando em vossa mão,\num acorde imenso brada:" },
      { tipo: "Refrão", texto: "Quem me fez foi vosso amor,\nglória a vós, Pai Criador! (Bis)" },
      { tipo: "Estrofe", texto: "3. Filho eterno, nosso irmão,\nvossa morte deu-nos vida,\nvosso sangue, salvação.\nToda a Igreja, agradecida," },
      { tipo: "Refrão", texto: "Louva, exalta a vós, Jesus,\nglória canta a vossa cruz! (Bis)" },
      { tipo: "Estrofe", texto: "4. Deus Espírito, sol de amor,\nprocedeis do Pai, do Filho.\nVossos dons sempre mandais\na nós pobres que cantamos." },
      { tipo: "Refrão", texto: "Santo, Santo é o Senhor,\nuno e trino, Deus de amor. (Bis)" },
    ]},
    { titulo: "Ó Trindade Vos Louvamos", autor: "", secoes: [
      { tipo: "Refrão", texto: "Ó Trindade, vos louvamos,\nvos louvamos pela vossa comunhão.\nQue esta mesa favoreça,\nfavoreça nossa comunicação." },
      { tipo: "Estrofe", texto: "1. Contra toda tentação da ganância e do poder.\nNossas bocas gritem juntas a palavra do viver." },
      { tipo: "Refrão", texto: "Ó Trindade, vos louvamos,\nvos louvamos pela vossa comunhão.\nQue esta mesa favoreça,\nfavoreça nossa comunicação." },
      { tipo: "Estrofe", texto: "2. Na montanha com Jesus, no encontro com o Pai.\nRecebemos a mensagem: ide ao mundo e o transformai." },
      { tipo: "Refrão", texto: "Ó Trindade, vos louvamos,\nvos louvamos pela vossa comunhão.\nQue esta mesa favoreça,\nfavoreça nossa comunicação." },
      { tipo: "Estrofe", texto: "3. Deus nos fala na história e nos chama à conversão.\nVamos ser palavras vivas proclamando a salvação." },
      { tipo: "Refrão", texto: "Ó Trindade, vos louvamos,\nvos louvamos pela vossa comunhão.\nQue esta mesa favoreça,\nfavoreça nossa comunicação." },
    ]},
  ],

  final: [
    { titulo: "Teu Nome, Senhor, É Tão Bonito", autor: "Padre Jocy Rodrigues", secoes: [
      { tipo: "Única", texto: "Teu nome, Senhor, é tão bonito\ntu moras no céu, lá nas alturas\naté criancinhas que ainda mamam\njá sabem que vences o inimigo.\nOlhando pro céu que tu fizeste\neu vejo as estrelas, vejo a lua\nentendo que o homem vale muito\npois tudo pra ele tu fizeste.\nMenor um pouquinho do que os anjos\nmas cheio de glória e de valor\nde Ti recebeu poder e força\nde tudo vencer e dominar." },
    ]},
    { titulo: "Trindade Santa", autor: "", secoes: [
      { tipo: "Estrofe", texto: "1. Trindade Santa como entender\nmistério da fé quero viver\né compreender tua imensidão\ne mesmo sem perceber estar em tuas mãos.\nViva o amor de Deus ele está aqui\ntão próximo a você pare de resistir\nele te ama, ele tudo criou\ne em nós seus filhos, sua imagem retratou." },
      { tipo: "Refrão", texto: "Só vós sois Deus, só vós o Santíssimo,\nsó vós o Senhor.\nAquele que ama no Pai, revela-se no Filho,\nnos manda o Espírito Santo. (Bis)" },
    ]},
    { titulo: "Aleluia Glória ao Senhor", autor: "", secoes: [
      { tipo: "Refrão", texto: "Aleluia, aleluia, aleluia, aleluia\nGlória ao Senhor! (Bis)" },
      { tipo: "Estrofe", texto: "1. Glória ao Senhor nosso Pai\nGlória ao Senhor nosso Pai\nGlória ao Senhor nosso Pai\nGlória ao Senhor!" },
      { tipo: "Refrão", texto: "Aleluia, aleluia, aleluia, aleluia\nGlória ao Senhor! (Bis)" },
      { tipo: "Estrofe", texto: "2. Glória a Jesus Cristo Redentor\nGlória a Jesus Cristo Redentor\nGlória a Jesus Cristo Redentor\nGlória ao Senhor!" },
      { tipo: "Refrão", texto: "Aleluia, aleluia, aleluia, aleluia\nGlória ao Senhor! (Bis)" },
      { tipo: "Estrofe", texto: "3. Glória ao Espírito Santo\nGlória ao Espírito Santo\nGlória ao Espírito Santo\nGlória ao Senhor!" },
      { tipo: "Refrão", texto: "Aleluia, aleluia, aleluia, aleluia\nGlória ao Senhor! (Bis)" },
      { tipo: "Estrofe", texto: "4. Glória à Santíssima Trindade\nGlória à Santíssima Trindade\nGlória à Santíssima Trindade\nGlória ao Senhor!" },
      { tipo: "Refrão", texto: "Aleluia, aleluia, aleluia, aleluia\nGlória ao Senhor! (Bis)" },
    ]},
  ],
};

// Leituras fixas Trindade
const slidesFixosTrindade = {
  primeiraLeitura: ["Êxodo 34,4b-6.8-9"],
  salmo: ["A vós louvor, honra e glória eternamente!"],
  segundaLeitura: ["2 Coríntios 13,11-13"],
  aclamacao: ["Aleluia!"],
  evangelho: ["João 3,16-18"],
  preces: ["Trindade Santa, ouve a nossa prece!"],
};

// Oração Eucarística I Trindade
const oracaoEucaristicaTrindade = [
  "Abençoai nossa oferenda, ó Senhor!",
  "Lembrai-vos, ó Pai, de vossos filhos!",
  "Em comunhão com os vossos Santos\nvos louvamos!",
  "Enviai o vosso Espírito Santo!",
  "Anunciamos, Senhor, a vossa morte\ne proclamamos a vossa ressurreição.\nVinde, Senhor Jesus!",
  "Todas as vezes que comemos deste pão\ne bebemos deste cálice,\nanunciamos, Senhor, a vossa morte,\nenquanto esperamos a vossa vinda!",
  "Aceitai, ó Senhor, a nossa oferta!",
  "O Espírito nos una num só corpo!",
  "Concedei-lhes, ó Senhor, a luz eterna!",
];


function gerarMissaTrindade(opcoes) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_4x3";
  pres.title = "LiturgiaPlay — Santíssima Trindade";

  // 1. CAPA
  addSlideCapa(pres, "Santíssima Trindade — 31 de maio de 2026", "Solenidade da\nSantíssima Trindade");

  // 2. ENTRADA
  const entrada = bancoTrindade.cantos_entrada.find(c => c.titulo === opcoes.entrada);
  if (entrada) renderMusica(pres, entrada);
  addSlidePreto(pres);

  // 3. ASPERSÃO (opcional)
  if (opcoes.aspersao) {
    const asp = bancoTrindade.aspersao.find(c => c.titulo === opcoes.aspersao);
    if (asp) renderMusica(pres, asp);
    addSlidePreto(pres);
  }

  // 4. ATO PENITENCIAL
  if (opcoes.penitencial) {
    const pen = bancoTrindade.penitencial.find(c => c.titulo === opcoes.penitencial);
    if (pen) renderMusica(pres, pen);
  }
  addSlidePreto(pres);

  // 5. GLÓRIA
  if (opcoes.gloria) {
    const gl = bancoTrindade.gloria.find(c => c.titulo === opcoes.gloria);
    if (gl) renderMusica(pres, gl);
  }
  addSlidePreto(pres);

  // 6. LITURGIA DA PALAVRA
  addSlideFixo(pres, "I LEITURA:", slidesFixosTrindade.primeiraLeitura);
  addSlideFixo(pres, "Salmo responsorial", slidesFixosTrindade.salmo);
  addSlideFixo(pres, "II LEITURA:", slidesFixosTrindade.segundaLeitura);

  // 7. ACLAMAÇÃO
  if (opcoes.aclamacao) {
    const ac = bancoTrindade.aclamacao.find(c => c.titulo === opcoes.aclamacao);
    if (ac) renderMusica(pres, ac);
  }

  addSlideFixo(pres, "EVANGELHO", slidesFixosTrindade.evangelho);
  addSlidePreto(pres);

  // 8. ORAÇÃO DOS FIÉIS
  addSlideFixo(pres, "Oração dos Fiéis", slidesFixosTrindade.preces);
  addSlidePreto(pres);

  // 9. OFERTÓRIO
  if (opcoes.ofertorio) {
    const of = bancoTrindade.ofertorio.find(c => c.titulo === opcoes.ofertorio);
    if (of) renderMusica(pres, of);
  }
  addSlidePreto(pres);

  // 10. SANTO
  if (opcoes.santo) {
    const st = bancoTrindade.santo.find(c => c.titulo === opcoes.santo);
    if (st) renderMusica(pres, st);
  }
  addSlidePreto(pres);

  // 11. ORAÇÃO EUCARÍSTICA I
  oracaoEucaristicaTrindade.forEach((resposta, i) => {
    if (i === 4) addSlidePreto(pres);
    addSlideTexto(pres, resposta, true);
  });
  addSlidePreto(pres);

  // 12. CORDEIRO DE DEUS
  if (opcoes.cordeiro) {
    const cord = bancoTrindade.cordeiro.find(c => c.titulo === opcoes.cordeiro);
    if (cord) renderMusica(pres, cord);
  }
  addSlidePreto(pres);

  // 13. COMUNHÃO
  if (opcoes.comunhao) {
    const com = bancoTrindade.comunhao.find(c => c.titulo === opcoes.comunhao);
    if (com) renderMusica(pres, com);
  }
  addSlidePreto(pres);

  // 14. CANTO FINAL
  if (opcoes.final) {
    const fin = bancoTrindade.final.find(c => c.titulo === opcoes.final);
    if (fin) renderMusica(pres, fin);
  }

  return pres;
}



// =============================================
// BANCO — 10º DOMINGO DO TEMPO COMUM
// =============================================

const bancoDecimo = {

  cantos_entrada: [
    {
      titulo: "O Senhor É Minha Luz e Salvação (1)",
      secoes: [
        { tipo: "Refrão", texto: "O Senhor é minha luz e salvação,\nde quem eu terei medo?\nÉ Ele a defesa da minha vida,\nperante quem eu tremerei?" },
        { tipo: "Estrofe", texto: "1. Quando perversos me atacam,\npara rasgar minha carne,\nSão inimigos e opressores\nque tropeçam e sucumbem." },
        { tipo: "Refrão", texto: "O Senhor é minha luz e salvação,\nde quem eu terei medo?\nÉ Ele a defesa da minha vida,\nperante quem eu tremerei?" },
        { tipo: "Estrofe", texto: "2. Se contra mim se acamparem,\nnão temerá o meu coração;\nSe contra mim a batalha estourar,\nmesmo assim eu confiarei." },
        { tipo: "Refrão", texto: "O Senhor é minha luz e salvação,\nde quem eu terei medo?\nÉ Ele a defesa da minha vida,\nperante quem eu tremerei?" },
        { tipo: "Estrofe", texto: "3. Peço ao Senhor uma coisa só,\ne é só isto que eu desejo:\nNo santuário do Senhor habitarei\npelos tempos infinitos." },
        { tipo: "Refrão", texto: "O Senhor é minha luz e salvação,\nde quem eu terei medo?\nÉ Ele a defesa da minha vida,\nperante quem eu tremerei?" },
        { tipo: "Estrofe", texto: "4. Em sua casa vou habitar,\nsaborear sua suavidade,\nA suavidade do meu Senhor\ne contemplá-lo no templo." },
        { tipo: "Refrão", texto: "O Senhor é minha luz e salvação,\nde quem eu terei medo?\nÉ Ele a defesa da minha vida,\nperante quem eu tremerei?" },
        { tipo: "Estrofe", texto: "5. Sob o seu teto me abrigará,\nnos dias do infortúnio;\nNo interior de sua tenda me esconde\ne sobre a rocha me protege." },
        { tipo: "Refrão", texto: "O Senhor é minha luz e salvação,\nde quem eu terei medo?\nÉ Ele a defesa da minha vida,\nperante quem eu tremerei?" },
      ]
    },
    {
      titulo: "O Senhor É Minha Luz (2)",
      secoes: [
        { tipo: "Refrão", texto: "O Senhor é minha luz,\nEle é minha salvação.\nQue poderei temer?\nDeus, minha proteção!" },
        { tipo: "Estrofe", texto: "1. O Senhor é minha luz,\nEle é minha salvação.\nO que é que eu vou temer?\nDeus é minha proteção.\nEle guarda minha vida,\neu não vou ter medo, não.\nEle guarda minha vida,\neu não vou ter medo, não." },
        { tipo: "Refrão", texto: "O Senhor é minha luz,\nEle é minha salvação.\nQue poderei temer?\nDeus, minha proteção!" },
        { tipo: "Estrofe", texto: "2. Quando os maus vêm avançando,\nprocurando me acuar,\ndesejando ver meu fim,\nquerendo me matar.\nInimigos opressores\né que vão se liquidar.\nInimigos opressores\né que vão se liquidar." },
        { tipo: "Refrão", texto: "O Senhor é minha luz,\nEle é minha salvação.\nQue poderei temer?\nDeus, minha proteção!" },
        { tipo: "Estrofe", texto: "3. Se um exército se armar\ncontra mim, não temerei.\nMeu coração está firme\ne firme ficarei.\nSe estourar uma batalha,\nmesmo assim, confiarei!\nSe estourar uma batalha,\nmesmo assim, confiarei!" },
        { tipo: "Refrão", texto: "O Senhor é minha luz,\nEle é minha salvação.\nQue poderei temer?\nDeus, minha proteção!" },
        { tipo: "Estrofe", texto: "4. Sei que hei de ver um dia\na bondade do Senhor\nlá, na terra dos viventes,\nviverei no seu amor.\nEspera em Deus! Cria coragem!\nEspera em Deus que é teu Senhor!\nEspera em Deus! Cria coragem!\nEspera em Deus que é teu Senhor!" },
        { tipo: "Refrão", texto: "O Senhor é minha luz,\nEle é minha salvação.\nQue poderei temer?\nDeus, minha proteção!" },
      ]
    },
    {
      titulo: "O Senhor É Minha Luz e Salvação (3)",
      secoes: [
        { tipo: "Refrão", texto: "O Senhor é minha luz e salvação,\nde quem eu terei medo?\nO Senhor é a proteção da minha vida,\nperante quem eu tremerei?" },
        { tipo: "Estrofe", texto: "1. Quando avançam os malvados contra mim,\nquerendo devorar-me,\nsão eles, inimigos e opressores,\nque tropeçam e sucumbem." },
        { tipo: "Refrão", texto: "O Senhor é minha luz e salvação,\nde quem eu terei medo?\nO Senhor é a proteção da minha vida,\nperante quem eu tremerei?" },
        { tipo: "Estrofe", texto: "2. Ao Senhor eu peço apenas uma coisa,\ne é só isto que eu desejo:\nHabitar no santuário do Senhor\npor toda a minha vida." },
        { tipo: "Refrão", texto: "O Senhor é minha luz e salvação,\nde quem eu terei medo?\nO Senhor é a proteção da minha vida,\nperante quem eu tremerei?" },
      ]
    },
    {
      titulo: "O Senhor É Minha Luz e Salvação (4)",
      secoes: [
        { tipo: "Refrão", texto: "O Senhor é minha luz e salvação,\na quem poderia eu temer?\nO Senhor é a proteção da minha vida,\nperante quem eu tremerei?" },
        { tipo: "Estrofe", texto: "1. Quando avançam os malvados contra mim,\nquerendo devorar-me,\nsão eles, inimigos e opressores,\nque tropeçam e sucumbem." },
        { tipo: "Refrão", texto: "O Senhor é minha luz e salvação,\na quem poderia eu temer?\nO Senhor é a proteção da minha vida,\nperante quem eu tremerei?" },
        { tipo: "Estrofe", texto: "2. Se os inimigos se acamparem contra mim,\nnão temerá meu coração;\nse contra mim uma batalha estourar,\nmesmo assim confiarei." },
        { tipo: "Refrão", texto: "O Senhor é minha luz e salvação,\na quem poderia eu temer?\nO Senhor é a proteção da minha vida,\nperante quem eu tremerei?" },
        { tipo: "Estrofe", texto: "3. Ao Senhor eu peço apenas uma coisa,\ne é só isto que eu desejo:\nHabitar no santuário do Senhor\npor toda a minha vida." },
        { tipo: "Refrão", texto: "O Senhor é minha luz e salvação,\na quem poderia eu temer?\nO Senhor é a proteção da minha vida,\nperante quem eu tremerei?" },
      ]
    },
    {
      titulo: "O Senhor É Minha Luz e Salvação (5)",
      secoes: [
        { tipo: "Refrão", texto: "O Senhor é minha luz e minha salvação,\na quem poderia eu temer?\nO Senhor é o baluarte de minha vida,\nperante quem tremerei?" },
        { tipo: "Estrofe", texto: "1. Se os inimigos se acamparem contra mim,\nnão temerá meu coração." },
        { tipo: "Refrão", texto: "O Senhor é minha luz e minha salvação,\na quem poderia eu temer?\nO Senhor é o baluarte de minha vida,\nperante quem tremerei?" },
        { tipo: "Estrofe", texto: "2. Se contra mim uma batalha estourar,\nmesmo assim confiarei." },
        { tipo: "Refrão", texto: "O Senhor é minha luz e minha salvação,\na quem poderia eu temer?\nO Senhor é o baluarte de minha vida,\nperante quem tremerei?" },
        { tipo: "Estrofe", texto: "3. Quero habitar no santuário do Senhor\npor toda a minha vida!" },
        { tipo: "Refrão", texto: "O Senhor é minha luz e minha salvação,\na quem poderia eu temer?\nO Senhor é o baluarte de minha vida,\nperante quem tremerei?" },
      ]
    },
  ],

  penitencial: [
    {
      titulo: "Senhor, Tende Piedade dos Corações Arrependidos",
      secoes: [
        { tipo: "Estrofe", texto: "Senhor, tende piedade\ndos corações arrependidos.\nTende piedade de nós! (4x)" },
        { tipo: "Estrofe", texto: "Jesus, tende piedade\ndos pecadores, tão humilhados!\nTende piedade de nós! (4x)" },
        { tipo: "Estrofe", texto: "Senhor, tende piedade\nintercedendo por nós ao Pai!\nTende piedade de nós! (4x)" },
      ]
    },
    {
      titulo: "Quero Confessar a Ti",
      secoes: [
        { tipo: "Estrofe", texto: "Quero confessar a Ti,\nilumina a minha alma.\nEu reconheço, sou pecador.\nDiante de mim eu sei\nestá sempre o meu pecado.\nFoi contra vós que eu pequei." },
        { tipo: "Refrão", texto: "Kyrie eleison!\nChriste eleison!\nKyrie eleison!" },
      ]
    },
    {
      titulo: "Senhor Que Vieste Salvar",
      secoes: [
        { tipo: "Estrofe", texto: "1. Senhor que viestes salvar\nos corações arrependidos.\nPiedade, Piedade,\nPiedade de nós. (Bis)" },
        { tipo: "Estrofe", texto: "2. Ó, Cristo que viestes chamar\nos pecadores humilhados.\nPiedade, Piedade,\nPiedade de nós. (Bis)" },
        { tipo: "Estrofe", texto: "3. Senhor que intercedeis por nós\njunto a Deus Pai que nos perdoa.\nPiedade, Piedade,\nPiedade de nós. (Bis)" },
      ]
    },
    {
      titulo: "Oh Senhor",
      secoes: [
        { tipo: "Estrofe", texto: "1. Oh Senhor, tende piedade de nós!\nOh Senhor, tende piedade de nós!" },
        { tipo: "Estrofe", texto: "2. Cristo, tende piedade de nós!\nCristo, tende piedade de nós!" },
        { tipo: "Estrofe", texto: "3. Oh Senhor, tende piedade de nós!\nOh Senhor, tende piedade de nós!" },
      ]
    },
    {
      titulo: "Kyrie Eleison (JMJ)",
      secoes: [
        { tipo: "Estrofe", texto: "1. Senhor que vieste salvar\nos corações arrependidos.\nKyrie Eleison, Eleison, Eleison. (Bis)" },
        { tipo: "Estrofe", texto: "2. O Cristo que vieste chamar\nos pecadores humilhados.\nChriste Eleison, Eleison, Eleison. (Bis)" },
        { tipo: "Estrofe", texto: "3. Senhor que intercedeis por nós\njuntos a Deus Pai que nos perdoa.\nKyrie Eleison, Eleison, Eleison. (Bis)" },
      ]
    },
    {
      titulo: "Senhor, Vós Sois o Caminho",
      secoes: [
        { tipo: "Estrofe", texto: "Senhor vós sois o caminho,\nguiai-nos ao Pai com carinho.\nDe nós tende piedade,\nSenhor, tende piedade!" },
        { tipo: "Estrofe", texto: "Ó Cristo, sois a verdade,\nenchei-nos de caridade.\nDe nós tende piedade,\nÓ Cristo, tende piedade!" },
        { tipo: "Estrofe", texto: "Senhor, vós sois nossa vida,\nbuscais a ovelha perdida.\nDe nós tende piedade,\nSenhor, tende piedade!" },
      ]
    },
  ],

  aclamacao: [
    {
      titulo: "Aleluia! (1)",
      secoes: [
        { tipo: "Refrão", texto: "Aleluia! Aleluia! Aleluia! Aleluia! (Bis)" },
        { tipo: "Estrofe", texto: "Foi o Senhor quem me mandou\nboas notícias anunciar;\nao pobre, a quem está no cativeiro,\nlibertação eu vou proclamar." },
        { tipo: "Refrão", texto: "Aleluia! Aleluia! Aleluia! Aleluia! (Bis)" },
      ]
    },
    {
      titulo: "Aleluia! (2)",
      secoes: [
        { tipo: "Refrão", texto: "Aleluia! (11x)" },
        { tipo: "Estrofe", texto: "Foi o Senhor quem me mandou\nboas notícias anunciar;\nao pobre, a quem está no cativeiro,\nlibertação eu vou proclamar." },
        { tipo: "Refrão", texto: "Aleluia! (11x)" },
      ]
    },
    {
      titulo: "Aleluia! (3)",
      secoes: [
        { tipo: "Refrão", texto: "Aleluia! Aleluia! Aleluia! (Bis)" },
        { tipo: "Estrofe", texto: "Foi o Senhor quem me mandou\nboas notícias anunciar;\nao pobre, a quem está no cativeiro,\nlibertação eu vou proclamar." },
        { tipo: "Refrão", texto: "Aleluia! Aleluia! Aleluia! (Bis)" },
      ]
    },
  ],

  ofertorio: [
    { titulo: "Bendito e Louvado Seja o Pai", secoes: [
      { tipo: "Estrofe", texto: "1. Bendito e louvado seja\no Pai nosso criador,\no pão que nós recebemos\né prova do seu amor.\nÉ o fruto de sua terra\ndo povo trabalhador,\nna missa é transformado\nno corpo do Salvador." },
      { tipo: "Refrão", texto: "Bendito seja Deus,\nbendito seu amor.\nBendito seja Deus\nPai onipotente, nosso Criador. (Bis)" },
      { tipo: "Estrofe", texto: "2. Bendito e louvado seja\no Pai nosso criador,\no vinho que recebemos\né prova do seu amor.\nÉ o fruto de sua terra\ndo povo trabalhador,\nna missa é transformado\nno sangue do Salvador." },
      { tipo: "Refrão", texto: "Bendito seja Deus,\nbendito seu amor.\nBendito seja Deus\nPai onipotente, nosso Criador. (Bis)" },
    ]},
    { titulo: "Daqui do Meu Lugar", secoes: [
      { tipo: "Estrofe", texto: "1. Daqui do meu lugar, eu olho o teu altar\ne fico a imaginar aquele pão, aquela refeição.\nPartiste aquele pão e o deste aos teus irmãos,\ncriaste a religião do pão do céu,\ndo pão que vem do céu." },
      { tipo: "Refrão", texto: "Somos a Igreja do pão,\ndo pão repartido, e do abraço e da paz. (Bis)" },
      { tipo: "Estrofe", texto: "2. Daqui do meu lugar, eu olho o teu altar\ne fico a imaginar aquela paz, aquela comunhão.\nViveste aquela paz e a deste aos teus irmãos,\ncriaste a religião do pão da paz,\nda paz que vem do céu." },
      { tipo: "Refrão", texto: "Somos a Igreja da paz,\nda paz partilhada, e do abraço e do pão. (Bis)" },
    ]},
    { titulo: "De Mãos Estendidas", secoes: [
      { tipo: "Refrão", texto: "De mãos estendidas ofertamos\no que de graça, recebemos. (Bis)" },
      { tipo: "Estrofe", texto: "1. A Natureza tão bela,\nque é louvor, que é serviço.\nO Sol que ilumina as trevas,\ntransformando-as em luz.\nO dia que nos traz o pão\ne a noite que nos dá o repouso.\nOfertemos ao Senhor\no louvor da criação." },
      { tipo: "Refrão", texto: "De mãos estendidas ofertamos\no que de graça, recebemos. (Bis)" },
      { tipo: "Estrofe", texto: "2. Nossa vida, toda inteira,\nofertamos ao Senhor.\nCom o vinho e com o pão,\nofertemos ao Senhor\nnossa vida, toda inteira,\no louvor da criação." },
      { tipo: "Refrão", texto: "De mãos estendidas ofertamos\no que de graça, recebemos. (Bis)" },
    ]},
    { titulo: "Muitos Grãos de Trigo", secoes: [
      { tipo: "Estrofe", texto: "1. Muitos grãos de trigo, se tornaram pão,\nhoje são Teu corpo, ceia e comunhão.\nMuitos grãos de trigo, se tornaram pão." },
      { tipo: "Refrão", texto: "Toma, Senhor, nossa vida em ação,\npara mudá-la em fruto e missão.\nToma, Senhor, nossa vida em ação,\npara mudá-la em missão." },
      { tipo: "Estrofe", texto: "2. Muitos cachos de uva, se tornaram vinho,\nhoje são Teu sangue, força no caminho.\nMuitos cachos de uva, se tornaram vinho." },
      { tipo: "Refrão", texto: "Toma, Senhor, nossa vida em ação,\npara mudá-la em fruto e missão.\nToma, Senhor, nossa vida em ação,\npara mudá-la em missão." },
      { tipo: "Estrofe", texto: "3. Muitas são as vidas, feitas vocação,\nhoje oferecidas em consagração.\nMuitas são as vidas, feitas vocação." },
      { tipo: "Refrão", texto: "Toma, Senhor, nossa vida em ação,\npara mudá-la em fruto e missão.\nToma, Senhor, nossa vida em ação,\npara mudá-la em missão." },
    ]},
    { titulo: "A Mesa Santa", secoes: [
      { tipo: "Estrofe", texto: "1. A mesa santa que preparamos,\nmãos que se elevam a Ti, ó Senhor.\nO pão e o vinho, frutos da terra,\nduro trabalho, carinho e amor." },
      { tipo: "Refrão", texto: "Oh, recebe, Senhor!\nOh, recebe, Senhor! (Bis)" },
      { tipo: "Estrofe", texto: "2. Flores, espinhos, dor e alegria,\npais, mães e filhos diante do altar.\nA nossa oferta em nova festa,\na nossa dor vem, Senhor, transformar." },
      { tipo: "Refrão", texto: "Oh, recebe, Senhor!\nOh, recebe, Senhor! (Bis)" },
      { tipo: "Estrofe", texto: "3. A vida nova, nova família,\nque celebramos, aqui tem lugar.\nTua bondade vem com fartura,\né só saber reunir, partilhar!" },
      { tipo: "Refrão", texto: "Oh, recebe, Senhor!\nOh, recebe, Senhor! (Bis)" },
    ]},
    { titulo: "Um Coração para Amar", secoes: [
      { tipo: "Estrofe", texto: "1. Um coração para amar, pra perdoar e sentir,\npara chorar e sorrir, ao me criar Tu me destes.\nUm coração pra sonhar, inquieto e sempre a bater,\nansioso por entender as coisas que Tu disseste." },
      { tipo: "Refrão", texto: "Eis o que eu venho Te dar,\neis o que eu ponho no altar.\nToma, Senhor, que ele é Teu,\nmeu coração não é meu. (Bis)" },
      { tipo: "Estrofe", texto: "2. Quero que o meu coração seja tão cheio de paz,\nque não se sinta capaz de sentir ódio ou rancor.\nQuero que a minha oração possa me amadurecer,\nleve-me a compreender as consequências do amor." },
      { tipo: "Refrão", texto: "Eis o que eu venho Te dar,\neis o que eu ponho no altar.\nToma, Senhor, que ele é Teu,\nmeu coração não é meu. (Bis)" },
    ]},
    { titulo: "Os Grãos que Trago Aqui", secoes: [
      { tipo: "Estrofe", texto: "1. Os dons que trago aqui\nsão o que fiz, o que vivi.\nO pão que ofertarei,\npouco depois comungarei.\nAssim tudo o que é meu,\nsinto também que é de Deus." },
      { tipo: "Refrão", texto: "Esforço, trabalhos e sonhos,\no amor concreto e feliz deste dia.\nPor Cristo, com Cristo e em Cristo,\ntudo ofertamos ao Pai na alegria." },
      { tipo: "Estrofe", texto: "2. Jesus nos quis chamar\npara o seguir e ajudar.\nE aqui nos vai dizer\ncomo servir e oferecer.\nDeus pôs nas minhas mãos\npara eu partir com meus irmãos." },
      { tipo: "Refrão", texto: "Esforço, trabalhos e sonhos,\no amor concreto e feliz deste dia.\nPor Cristo, com Cristo e em Cristo,\ntudo ofertamos ao Pai na alegria." },
    ]},
    { titulo: "Pão e Vinho Te Apresentamos", secoes: [
      { tipo: "Estrofe", texto: "Pão e vinho Te apresentamos nesse altar,\ncomo sinal que Tu recolhes nossa oferta.\nTudo o que somos deixamos aqui. (Bis)" },
      { tipo: "Refrão", texto: "É um milagre que se dá,\no pão e o vinho em corpo e sangue\nvão se transformar.\nNão há limites para o amor,\nvem transformar também minha vida,\nÓ Senhor.\nÉ Teu esse milagre de amor." },
    ]},
    { titulo: "Bendito Seja Deus Pai", secoes: [
      { tipo: "Estrofe", texto: "1. Bendito seja Deus Pai\ndo universo criador,\npelo pão que nós recebemos,\nfoi de graça e com amor." },
      { tipo: "Refrão", texto: "O homem que trabalha\nfaz a terra produzir.\nO trabalho multiplica os dons\nque nós vamos repartir." },
      { tipo: "Estrofe", texto: "2. Bendito seja Deus Pai\ndo universo o criador,\npelo vinho que nós recebemos,\nfoi de graça e com amor." },
      { tipo: "Refrão", texto: "O homem que trabalha\nfaz a terra produzir.\nO trabalho multiplica os dons\nque nós vamos repartir." },
      { tipo: "Estrofe", texto: "3. E nós participamos\nda construção do mundo novo,\ncom Deus, que jamais despreza\nnossa imensa pequenez." },
      { tipo: "Refrão", texto: "O homem que trabalha\nfaz a terra produzir.\nO trabalho multiplica os dons\nque nós vamos repartir." },
    ]},
    { titulo: "Um Consagrado para Amar", secoes: [
      { tipo: "Estrofe", texto: "Venho, Senhor, me ofertar,\na minha vida consagrar.\nQuero renovar o meu sim,\nque Tua vontade se faça em mim.\nRenova, Senhor, minha vocação. (Bis)" },
      { tipo: "Refrão", texto: "Um consagrado para amar,\num consagrado pra se doar.\nUm amor que tudo suporta,\num amor que não dá pra improvisar.\nUm consagrado para amar,\num consagrado pra se doar.\nUm amor que não busca interesses seus,\né o mais puro amor, o amor de Deus." },
    ]},
    { titulo: "Sabes, Senhor, o que Temos É Tão Pouco", secoes: [
      { tipo: "Refrão", texto: "Sabes, Senhor,\no que temos é tão pouco pra dar.\nMas este pouco\nnós queremos com os irmãos compartilhar." },
      { tipo: "Estrofe", texto: "1. Queremos nesta hora,\ndiante dos irmãos,\ncomprometer a vida,\nbuscando a união." },
      { tipo: "Refrão", texto: "Sabes, Senhor,\no que temos é tão pouco pra dar.\nMas este pouco\nnós queremos com os irmãos compartilhar." },
      { tipo: "Estrofe", texto: "2. Sabemos que é difícil\nos bens compartilhar,\nmas com a tua graça,\nSenhor, podemos dar." },
      { tipo: "Refrão", texto: "Sabes, Senhor,\no que temos é tão pouco pra dar.\nMas este pouco\nnós queremos com os irmãos compartilhar." },
      { tipo: "Estrofe", texto: "3. Olhando o teu exemplo,\nSenhor, vamos seguir,\nfazendo o bem aos homens,\nsem nada exigir." },
      { tipo: "Refrão", texto: "Sabes, Senhor,\no que temos é tão pouco pra dar.\nMas este pouco\nnós queremos com os irmãos compartilhar." },
    ]},
  ],

  comunhao: [
    { titulo: "Ó Meu Deus Sois o Rochedo que Me Abriga", secoes: [
      { tipo: "Refrão", texto: "Ó meu Deus sois o rochedo que me abriga,\nminha força e poderosa salvação.\nSois meu escudo e proteção, em vós espero." },
      { tipo: "Estrofe", texto: "1. Eu vos amo ó Senhor, Sois minha força,\nminha rocha, meu refúgio e Salvador." },
      { tipo: "Refrão", texto: "Ó meu Deus sois o rochedo que me abriga,\nminha força e poderosa salvação.\nSois meu escudo e proteção, em vós espero." },
      { tipo: "Estrofe", texto: "2. Invocarei o meu Senhor, a ele a glória,\ne dos meus perseguidores serei salvo." },
      { tipo: "Refrão", texto: "Ó meu Deus sois o rochedo que me abriga,\nminha força e poderosa salvação.\nSois meu escudo e proteção, em vós espero." },
      { tipo: "Estrofe", texto: "3. Ao Senhor eu invoquei na minha angústia\ne elevei o meu clamor para o meu Deus." },
      { tipo: "Refrão", texto: "Ó meu Deus sois o rochedo que me abriga,\nminha força e poderosa salvação.\nSois meu escudo e proteção, em vós espero." },
      { tipo: "Estrofe", texto: "4. De seu Templo ele escutou a minha voz\ne chegou a seus ouvidos o meu grito." },
      { tipo: "Refrão", texto: "Ó meu Deus sois o rochedo que me abriga,\nminha força e poderosa salvação.\nSois meu escudo e proteção, em vós espero." },
    ]},
    { titulo: "Ó Senhor, aos Doentes Vieste", secoes: [
      { tipo: "Refrão", texto: "Ó Senhor, aos doentes vieste,\npecadores, com eles sentaste.\nO Teu Corpo e Teu Sangue lhes deste,\naos famintos Tu alimentaste." },
      { tipo: "Estrofe", texto: "1. Um canto novo ao Senhor,\nó terras todas, cantai!\nLouvai Seu Nome bendito,\ndiariamente aclamai!\nSua glória, Seus grandes feitos\naos povos todos contai." },
      { tipo: "Refrão", texto: "Ó Senhor, aos doentes vieste,\npecadores, com eles sentaste.\nO Teu Corpo e Teu Sangue lhes deste,\naos famintos Tu alimentaste." },
      { tipo: "Estrofe", texto: "2. Ele é o maior dos senhores,\nmerece nosso louvor;\ne mais do que aos deuses todos,\nnós lhe devemos temor.\nOs outros deuses são nada,\nEle é, do céu, Criador." },
      { tipo: "Refrão", texto: "Ó Senhor, aos doentes vieste,\npecadores, com eles sentaste.\nO Teu Corpo e Teu Sangue lhes deste,\naos famintos Tu alimentaste." },
      { tipo: "Estrofe", texto: "3. Sabei que o Senhor é Rei\ne traz justiça a esta terra.\nAlegrem-se o mar e os peixes\ne tudo o que o mundo encerra.\nOs campos, plantas, montanhas\ne as árvores da floresta." },
      { tipo: "Refrão", texto: "Ó Senhor, aos doentes vieste,\npecadores, com eles sentaste.\nO Teu Corpo e Teu Sangue lhes deste,\naos famintos Tu alimentaste." },
    ]},
    { titulo: "Senhor Tu Vieste Mostrar aos Errados", secoes: [
      { tipo: "Refrão", texto: "Senhor, tu vieste mostrar aos errados\nde novo, o caminho da casa do Pai.\nTu deste teu corpo, tu deste teu sangue,\npra ser o sustento do Filho que cai." },
      { tipo: "Estrofe", texto: "1. Vamos juntos dar glória ao Senhor\ne a seu nome fazer louvação.\nProcurei o Senhor, me atendeu,\nme livrou de uma grande aflição." },
      { tipo: "Refrão", texto: "Senhor, tu vieste mostrar aos errados\nde novo, o caminho da casa do Pai.\nTu deste teu corpo, tu deste teu sangue,\npra ser o sustento do Filho que cai." },
      { tipo: "Estrofe", texto: "2. Olhem todos pra ele e se alegrem,\ntodo tempo sua boca sorria!\nEste pobre gritou e ele ouviu,\nfiquei livre de minha agonia." },
      { tipo: "Refrão", texto: "Senhor, tu vieste mostrar aos errados\nde novo, o caminho da casa do Pai.\nTu deste teu corpo, tu deste teu sangue,\npra ser o sustento do Filho que cai." },
      { tipo: "Estrofe", texto: "3. Acampou na batalha seu anjo,\ndefendendo seu povo e o livrando.\nProvem todos, pra ver como é bom\no Senhor que nos vai abrigando." },
      { tipo: "Refrão", texto: "Senhor, tu vieste mostrar aos errados\nde novo, o caminho da casa do Pai.\nTu deste teu corpo, tu deste teu sangue,\npra ser o sustento do Filho que cai." },
      { tipo: "Estrofe", texto: "4. Santos todos, adorem o Senhor,\naos que o amam, nenhum mal assalta.\nQuem é rico, empobrece e tem fome,\nmas, a quem busca a Deus, nada falta." },
      { tipo: "Refrão", texto: "Senhor, tu vieste mostrar aos errados\nde novo, o caminho da casa do Pai.\nTu deste teu corpo, tu deste teu sangue,\npra ser o sustento do Filho que cai." },
    ]},
  ],

  final: [
    { titulo: "A Escolhida", secoes: [
      { tipo: "Estrofe", texto: "1. Uma entre todas foi a escolhida,\nfoste tu, Maria, serva preferida,\nMãe do meu Senhor,\nMãe do meu Salvador." },
      { tipo: "Refrão", texto: "Maria, cheia de graça e consolo,\nvenha caminhar com teu povo.\nNossa Mãe sempre serás.\nMaria, cheia de graça e consolo,\nvenha caminhar com teu povo.\nNossa Mãe sempre serás." },
      { tipo: "Estrofe", texto: "2. Roga pelos pecadores desta Terra,\nroga pelo povo que em seu Deus espera,\nMãe do meu Senhor,\nMãe do meu Salvador." },
      { tipo: "Refrão", texto: "Maria, cheia de graça e consolo,\nvenha caminhar com teu povo.\nNossa Mãe sempre serás.\nMaria, cheia de graça e consolo,\nvenha caminhar com teu povo.\nNossa Mãe sempre serás." },
    ]},
    { titulo: "Pelas Estradas da Vida", secoes: [
      { tipo: "Estrofe", texto: "1. Pelas estradas da vida,\nnunca sozinho estás.\nContigo pelo caminho,\nSanta Maria vai." },
      { tipo: "Refrão", texto: "Ó, vem conosco, vem caminhar,\nSanta Maria vem.\nÓ, vem conosco, vem caminhar,\nSanta Maria vem." },
      { tipo: "Estrofe", texto: "2. Mesmo que digam os homens\ntu nada podes mudar,\nluta por um mundo novo\nde unidade e paz." },
      { tipo: "Refrão", texto: "Ó, vem conosco, vem caminhar,\nSanta Maria vem.\nÓ, vem conosco, vem caminhar,\nSanta Maria vem." },
    ]},
    { titulo: "Quem É Essa que Avança", secoes: [
      { tipo: "Estrofe", texto: "Quem é esta que avança como aurora,\ntemível como exército em ordem de batalha,\nbrilhante como o Sol e como a Lua,\nmostrando o caminho aos filhos seus." },
      { tipo: "Refrão", texto: "Ah, ah, ah, minha alma glorifica ao Senhor,\nmeu espírito exulta em Deus, meu Salvador." },
    ]},
    { titulo: "Quando Jesus Passar", secoes: [
      { tipo: "Refrão", texto: "Quando Jesus passar,\nquando Jesus passar,\nquando Jesus passar,\neu quero estar no meu lugar." },
      { tipo: "Estrofe", texto: "1. No meu telônio ou jogando a rede,\nsob a figueira ou a caminhar,\nbuscando água pra minha sede,\nquerendo ver meu Senhor passar." },
      { tipo: "Refrão", texto: "Quando Jesus passar,\nquando Jesus passar,\nquando Jesus passar,\neu quero estar no meu lugar." },
      { tipo: "Estrofe", texto: "2. No meu trabalho e na minha casa,\nno meu estudo e no meu lazer,\nno compromisso e no meu descanso,\nno meu direito e no meu dever." },
      { tipo: "Refrão", texto: "Quando Jesus passar,\nquando Jesus passar,\nquando Jesus passar,\neu quero estar no meu lugar." },
    ]},
    { titulo: "Ensina o Teu Povo a Rezar", secoes: [
      { tipo: "Estrofe", texto: "1. Ensina o teu povo a rezar,\nMaria, mãe de Jesus.\nQue um dia o teu povo desperta\ne na certa vai ver a luz.\nQue um dia o teu povo se anima\ne caminha com teu Jesus." },
      { tipo: "Estrofe", texto: "2. Maria de Jesus Cristo,\nMaria de Deus, Maria mulher.\nQue um dia o teu povo se anima\ne caminha com teu Jesus.\nQue um dia o teu povo se anima\ne caminha com teu Jesus." },
      { tipo: "Estrofe", texto: "3. Maria senhora nossa,\nMaria do povo, povo de Deus.\nEnsina o teu jeito perfeito\nde sempre escutar teu Deus.\nEnsina o teu jeito perfeito\nde sempre escutar teu Deus." },
    ]},
    { titulo: "Tomado pela Mão", secoes: [
      { tipo: "Refrão", texto: "Tomado pela mão com Jesus eu vou,\nsigo-o como ovelha que encontrou o pastor. (Bis)\nTomado pela mão com Jesus eu vou\naonde ele for. (Bis)" },
      { tipo: "Estrofe", texto: "1. Se Jesus me diz:\n\"Amigo, deixa tudo e vem comigo,\nonde tudo é mais formoso e mais feliz\". (Bis)" },
      { tipo: "Refrão", texto: "Tomado pela mão com Jesus eu vou,\nsigo-o como ovelha que encontrou o pastor. (Bis)\nTomado pela mão com Jesus eu vou\naonde ele for. (Bis)" },
      { tipo: "Estrofe", texto: "2. \"Eu te levarei amigo\na um lugar comigo,\nonde o sol e as estrelas brilham mais\". (Bis)" },
      { tipo: "Refrão", texto: "Tomado pela mão com Jesus eu vou,\nsigo-o como ovelha que encontrou o pastor. (Bis)\nTomado pela mão com Jesus eu vou\naonde ele for. (Bis)" },
    ]},
  ],
};

// Leituras fixas — 10º Domingo do Tempo Comum
const slidesFixosDecimo = {
  primeiraLeitura: ["Oséias 6,3-6"],
  salmo: ["A todo homem que procede retamente,\neu mostrarei a salvação que vem de Deus."],
  segundaLeitura: ["Romanos 4,18-25"],
  evangelho: ["Mateus 9,9-13"],
  preces: ["Salvai, Senhor, o Vosso Povo!"],
};

// 6 Orações Eucarísticas com Mistério da Fé
const oracoesEucaristicas = {
  1: {
    nome: "Oração Eucarística I",
    antesDoSanto: [],
    aposDoSanto: [
      "Abençoai nossa oferenda, ó Senhor!",
      "Lembrai-vos, ó Pai, de vossos filhos!",
      "Em comunhão com os vossos Santos\nvos louvamos!",
      "Enviai o vosso Espírito Santo!",
    ],
    aposDoMisterio: [
      "Aceitai, ó Senhor, a nossa oferta!",
      "O Espírito nos una num só corpo!",
      "Concedei-lhes, ó Senhor, a luz eterna!",
    ]
  },
  2: {
    nome: "Oração Eucarística II",
    antesDoSanto: [],
    aposDoSanto: [
      "Enviai o Vosso Espírito Santo!",
    ],
    aposDoMisterio: [
      "Aceitai, ó Senhor, a nossa oferta!",
      "O Espírito nos una num só corpo!",
      "Lembrai-vos, ó Pai, da vossa Igreja!",
      "Concedei-lhes, ó Senhor, a luz eterna!",
    ]
  },
  3: {
    nome: "Oração Eucarística III",
    antesDoSanto: [],
    aposDoSanto: [
      "Enviai o vosso Espírito Santo!",
    ],
    aposDoMisterio: [
      "Aceitai, ó Senhor, a nossa oferta!",
      "O Espírito nos una num só corpo!",
      "Fazei de nós um perfeita oferenda!",
      "Lembrai-vos, ó Pai, da vossa Igreja!",
    ]
  },
  4: {
    nome: "Oração Eucarística IV",
    antesDoSanto: [
      "Alegrai-nos, ó Pai, com a vossa luz!",
    ],
    aposDoSanto: [
      "Socorrei, com bondade, os que vos buscam!",
      "Por amor nos enviastes vosso Filho!",
      "Jesus Cristo deu-nos vida por sua morte!",
      "Santificai-nos pelo dom do vosso Espírito!",
      "Santificai nossa oferenda pelo Espírito.",
    ],
    aposDoMisterio: [
      "Recebei, ó Senhor, a nossa oferta!",
      "Fazei de nós um sacrifício de louvor!",
      "Lembrai-vos, ó Pai, dos vossos filhos!",
      "A todos saciai com vossa glória!",
      "Concedei-nos o convívio dos eleitos!",
    ]
  },
  5: {
    nome: "Oração Eucarística V",
    antesDoSanto: [],
    aposDoSanto: [
      "Mandai vosso Espírito Santo!",
    ],
    aposDoMisterio: [
      "Recebei, ó Senhor, a nossa oferta!",
      "O Espírito nos uma num só corpo!",
      "Caminhamos na estrada de Jesus!",
      "Lembrai-vos, ó Pai, da vossa Igreja!",
      "Esperamos entrar na vida eterna!",
      "A todos dai a luz que não se apaga!",
    ]
  },
  6: {
    nome: "Oração Eucarística VI",
    antesDoSanto: [],
    aposDoSanto: [
      "O vosso Filho permaneça entre nós!",
      "Mandai o vosso Espírito Santo!",
    ],
    aposDoMisterio: [
      "Aceitai, ó Senhor, a nossa oferta!",
      "Confirmai na caridade o vosso povo!",
      "Concedei-lhes, ó Senhor, a luz eterna!",
    ]
  },
};

const misteriosDaFe = {
  1: "Anunciamos, Senhor, a vossa morte\ne proclamamos a vossa ressurreição.\nVinde, Senhor Jesus!",
  2: "Todas as vezes que comemos deste pão\ne bebemos deste cálice,\nanunciamos, Senhor, a vossa morte,\nenquanto esperamos a vossa vinda!",
  3: "Salvador do mundo, salvai-nos,\nvós que nos libertastes\npela cruz e ressurreição.",
};

function gerarMissaDecimo(opcoes, addSlidePreto, addSlideTexto, addSlideFixo, addSlideCapa, renderMusica, banco, bancoTrindade) {
  const pres = new (require("pptxgenjs"))();
  pres.layout = "LAYOUT_4x3";
  pres.title = "LiturgiaPlay — 10º Domingo do Tempo Comum";

  const FONTE = "Arial";
  const COR_TEXTO = "000000";
  const MAX_LINHAS = 6;
  const FONT_MIN = 36;

  function calcFonte(linhas) {
    const n = linhas.length;
    const maxChars = Math.max(...linhas.map(l => l.length));
    let fs = n <= 2 ? 54 : n <= 3 ? 48 : n <= 4 ? 44 : n <= 5 ? 40 : 38;
    if (maxChars > 30) fs = Math.min(fs, 44);
    if (maxChars > 38) fs = Math.min(fs, 40);
    if (maxChars > 48) fs = Math.min(fs, 38);
    if (maxChars > 56) fs = Math.min(fs, 36);
    return Math.max(fs, FONT_MIN);
  }

  function addTexto(texto, negrito) {
    const linhas = texto.split("\n");
    for (let i = 0; i < linhas.length; i += MAX_LINHAS) {
      const bloco = linhas.slice(i, i + MAX_LINHAS);
      const slide = pres.addSlide();
      slide.background = { color: "FFFFFF" };
      const fs = calcFonte(bloco);
      slide.addText(bloco.map((l, idx) => ({ text: l, options: { breakLine: idx < bloco.length - 1, bold: negrito } })), {
        x: 0.3, y: 0.2, w: 9.4, h: 7.1, fontSize: fs, color: COR_TEXTO, fontFace: FONTE, align: "center", valign: "middle", lineSpacingMultiple: 1.2
      });
    }
  }

  function addPreto() {
    const s = pres.addSlide(); s.background = { color: "000000" };
  }

  function addFixo(titulo, linhas) {
    const slide = pres.addSlide();
    slide.background = { color: "FFFFFF" };
    slide.addText(titulo, { x: 0.3, y: 0.2, w: 9.4, h: 0.7, fontSize: 22, bold: true, color: COR_TEXTO, fontFace: FONTE, align: "left", valign: "top" });
    slide.addShape("rect", { x: 0.3, y: 0.95, w: 9.4, h: 0.03, fill: { color: "CCCCCC" }, line: { color: "CCCCCC" } });
    const n = linhas.length;
    const fs = n <= 1 ? 48 : n <= 2 ? 44 : n <= 3 ? 40 : n <= 4 ? 36 : 32;
    slide.addText(linhas.map((l, i) => ({ text: l, options: { breakLine: i < linhas.length - 1, bold: false } })), {
      x: 0.3, y: 1.1, w: 9.4, h: 6.3, fontSize: fs, color: COR_TEXTO, fontFace: FONTE, align: "center", valign: "middle", lineSpacingMultiple: 1.3
    });
  }

  function renderMus(musica) {
    musica.secoes.forEach(s => {
      const bold = s.tipo === "Refrão" || s.tipo === "Única";
      addTexto(s.texto, bold);
    });
  }

  // CAPA
  const capSlide = pres.addSlide();
  capSlide.background = { color: "FFFFFF" };
  capSlide.addText("10º Domingo do Tempo Comum — 07 de junho de 2026", { x: 1.0, y: 1.8, w: 8, h: 0.5, fontSize: 18, color: "888888", fontFace: FONTE, align: "center" });
  capSlide.addText("10º Domingo\ndo Tempo Comum", { x: 1.0, y: 2.4, w: 8, h: 2.5, fontSize: 48, bold: true, color: COR_TEXTO, fontFace: FONTE, align: "center", valign: "middle" });

  // ORAÇÃO EUCARÍSTICA selecionada
  const oeNum = opcoes.oracao_eucaristica || 1;
  const misterioNum = opcoes.misterio_fe || 1;
  const oe = oracoesEucaristicas[oeNum];

  // Respostas antes do Santo (Oração IV)
  if (oe.antesDoSanto.length > 0) {
    oe.antesDoSanto.forEach(r => addTexto(r, true));
    addPreto();
  }

  // ENTRADA
  const entrada = bancoDecimo.cantos_entrada.find(c => c.titulo === opcoes.entrada);
  if (entrada) renderMus(entrada);
  addPreto();

  // ATO PENITENCIAL
  if (opcoes.penitencial) {
    const pen = bancoDecimo.penitencial.find(c => c.titulo === opcoes.penitencial);
    if (pen) renderMus(pen);
  }
  addPreto();

  // GLÓRIA (reutiliza banco da Trindade)
  if (opcoes.gloria && bancoTrindade && bancoTrindade.gloria) {
    const gl = bancoTrindade.gloria.find(c => c.titulo === opcoes.gloria);
    if (gl) renderMus(gl);
  }
  addPreto();

  // LITURGIA DA PALAVRA
  addFixo("I LEITURA:", slidesFixosDecimo.primeiraLeitura);
  addFixo("Salmo responsorial 49(50)", slidesFixosDecimo.salmo);
  addFixo("II LEITURA:", slidesFixosDecimo.segundaLeitura);

  if (opcoes.aclamacao) {
    const ac = bancoDecimo.aclamacao.find(c => c.titulo === opcoes.aclamacao);
    if (ac) renderMus(ac);
  }
  addFixo("EVANGELHO", slidesFixosDecimo.evangelho);
  addPreto();

  // ORAÇÃO DOS FIÉIS
  addFixo("Oração dos Fiéis", slidesFixosDecimo.preces);
  addPreto();

  // OFERTÓRIO
  if (opcoes.ofertorio) {
    const of = bancoDecimo.ofertorio.find(c => c.titulo === opcoes.ofertorio);
    if (of) renderMus(of);
  }
  addPreto();

  // SANTO (reutiliza banco da Trindade)
  if (opcoes.santo && bancoTrindade && bancoTrindade.santo) {
    const st = bancoTrindade.santo.find(c => c.titulo === opcoes.santo);
    if (st) renderMus(st);
  }
  addPreto();

  // ORAÇÃO EUCARÍSTICA — respostas após o Santo
  oe.aposDoSanto.forEach(r => addTexto(r, true));
  addPreto();

  // MISTÉRIO DA FÉ
  const misterio = misteriosDaFe[misterioNum];
  if (misterio) addTexto(misterio, true);
  addPreto();

  // Respostas após o Mistério da Fé
  oe.aposDoMisterio.forEach(r => addTexto(r, true));
  addPreto();

  // CORDEIRO (reutiliza banco da Trindade)
  if (opcoes.cordeiro && bancoTrindade && bancoTrindade.cordeiro) {
    const cord = bancoTrindade.cordeiro.find(c => c.titulo === opcoes.cordeiro);
    if (cord) renderMus(cord);
  }
  addPreto();

  // COMUNHÃO
  if (opcoes.comunhao) {
    const com = bancoDecimo.comunhao.find(c => c.titulo === opcoes.comunhao);
    if (com) renderMus(com);
  }
  addPreto();

  // CANTO FINAL
  if (opcoes.final) {
    const fin = bancoDecimo.final.find(c => c.titulo === opcoes.final);
    if (fin) renderMus(fin);
  }

  return pres;
}


// =============================================
// ROTAS DA API
// =============================================

// Saúde — confirma que o servidor está no ar
app.get("/", (req, res) => {
  res.json({
    status: "online",
    servico: "LiturgiaPlay PPTX Generator",
    version: "2.0.0",
    domingo: "Pentecostes — 01 jun 2025"
  });
});

// Rota principal — recebe opções e devolve o PPTX completo
app.post("/gerar", async (req, res) => {
  try {
    const { opcoes, domingo } = req.body;

    if (!opcoes) {
      return res.status(400).json({ erro: "Envie as opções de cantos no campo 'opcoes'" });
    }

    // Seleciona a função correta baseada no domingo
    const isTrindade = domingo && domingo.toLowerCase().includes('trindade');
    const isDecimo = domingo && domingo.toLowerCase().includes('decimo');
    let pres;
    if (isDecimo) {
      pres = gerarMissaDecimo(opcoes, null, null, null, null, null, null, bancoTrindade);
    } else if (isTrindade) {
      pres = gerarMissaTrindade(opcoes);
    } else {
      pres = gerarMissa(opcoes);
    }
    const nomeArquivo = `liturgiaplay_${(domingo || "missa").replace(/\s/g, "_")}.pptx`;
    const buffer = await pres.write({ outputType: "nodebuffer" });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", `attachment; filename="${nomeArquivo}"`);
    res.setHeader("Access-Control-Expose-Headers", "Content-Disposition");
    res.send(buffer);

  } catch (err) {
    console.error("Erro ao gerar PPTX:", err);
    res.status(500).json({ erro: "Erro ao gerar o arquivo. Tente novamente." });
  }
});

// Lista os cantos disponíveis por momento
app.get("/cantos", (req, res) => {
  const lista = {};
  Object.keys(banco).forEach(momento => {
    lista[momento] = banco[momento].map(c => ({
      titulo: c.titulo,
      autor: c.autor || ""
    }));
  });
  res.json(lista);
});

app.listen(PORT, () => {
  console.log(`LiturgiaPlay API v2.0 rodando na porta ${PORT}`);
});
