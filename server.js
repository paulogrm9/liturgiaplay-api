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
console.log(`[${new Date().toISOString()}] PPTX gerado — domingo: ${req.body.domingo}`);
    if (!opcoes) {
      return res.status(400).json({ erro: "Envie as opções de cantos no campo 'opcoes'" });
    }

    const pres = gerarMissa(opcoes);
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
