// Código criado por Tiago Costa de Carvalho
// Ano de criação: 2020
//
// Grande ajuda da comunidade da Adobe para pequenas resoluções sobre uso dos objetos e modelos do Photoshop :)
//
// Esse script tem por idéia ser uma versão simplificada do script "Contact Sheet II" já que mesmo que ele cria uma janela mais amigável ao usuário, o mesmo
// não permite alterações posteriores ou mesmo complexifica muito como alterar coisas que achei cruciais para a melhor manipulação das imagens.

// Ele usa a mesma idéia da criação de células num plano base com um tamanho pré-determinado, em seguida, abre um prompt pedindo o caminho de uma pasta para
// gerar a lista de importações.

// As páginas são todas geradas automaticamente, e ao final da lista de arquivos, ele condensa todas as páginas na página inicial,
// Usará o caminho do molde da página ( CHECAR A FUNÇÃO >>>IMPORTA<<< )
// Pedirá o nome que vai ser inserido no Cabeçalho
// Vai multiplicar o molde da página para baixo para criar as pranchetas - Ao final do processo, estará pronto o arquivo já perfeitamente pronto para
// usar a exportação de pranchetas para PDF do próprio Photoshop

var startRulerUnits = app.preferences.rulerUnits;

app.preferences.rulerUnits = Units.PIXELS;

/////////////////////////////////////VARIÁVEIS GLOBAIS//////////////////////////////////////////

// DIMENSÕES FIXAS DA PÁGINA
var TamanhoA4 = [2480, 3308];
var MM300DPI = 12;

var largura = TamanhoA4[0];
var altura = TamanhoA4[1];
var margem = MM300DPI * 3;

var TamanhoA4Total = altura + 200;

// NÚMERO DE REFERÊNCIAS POR PÁGINA
var NUMColunas = 3;
var NUMLinhas = 3;

var LarguraColuna = largura / NUMColunas;
var AlturaColuna = altura / NUMLinhas;

var colunas = [];
var linhas = [];

var texto;
var nome;

var arial = app.fonts.getByName("ArialMT");

//NOMEANDO COLUNAS PARA ACESSIBILIDADE POSTERIOR
for (i = 0; i < NUMColunas; i++) {
  var texto = "col" + [i];
  colunas.push(texto);
}

//NOMEANDO LINHAS PARA ACESSIBILIDADE POSTERIOR
for (i = 0; i < NUMLinhas; i++) {
  var texto = "lin" + [i];
  colunas.push(texto);
}

//////////////////////////////////////FUNÇÕES///////////////////////////////////////////////////

// Apenas abre o arquivo em função do caminho
function Abre(caminho) {
  var idOpn = charIDToTypeID("Opn ");
  var desc140 = new ActionDescriptor();
  var iddontRecord = stringIDToTypeID("dontRecord");
  desc140.putBoolean(iddontRecord, false);
  var idforceNotify = stringIDToTypeID("forceNotify");
  desc140.putBoolean(idforceNotify, true);
  var idnull = charIDToTypeID("null");
  desc140.putPath(idnull, new File(caminho));
  var idDocI = charIDToTypeID("DocI");
  desc140.putInteger(idDocI, 337);
  executeAction(idOpn, desc140, DialogModes.NO);
}

// Faz o texto de nome do arquivo que foi importado, necessitando receber um nome e uma camada
function FazTexto(nome, camada, base) {
  var nomeSemExtensao = nome.split(".")[0];

  var layerAtual = app.activeDocument.artLayers.getByName(camada);
  var posicaoLayerAtual = layerAtual.bounds;

  var tamanhoX = posicaoLayerAtual[2] - posicaoLayerAtual[0];
  var tamanhoY = posicaoLayerAtual[3] - posicaoLayerAtual[1];

  var meioX = posicaoLayerAtual[0] + tamanhoX / 2;
  var meioY = posicaoLayerAtual[1] + tamanhoY / 2;

  var camadaTexto = app.activeDocument.artLayers.add();
  camadaTexto.kind = LayerKind.TEXT;
  camadaTexto.name = nome;

  var conteudo = camadaTexto.textItem;
  conteudo.contents = nomeSemExtensao + " - Base " + base;
  conteudo.justification = Justification.CENTER;

  conteudo.position = new Array(meioX, posicaoLayerAtual[3] + 55);
  conteudo.size = 15;
  conteudo.fauxBold = true;
}

//Adiciona, em função da largura em PX do arquivo, de qual categoria ele é:
//Retorna no nome para ser inserido na função "FazTexto"
function ChecaBase() {

  app.preferences.rulerUnits = Units.PIXELS;

  var larguraEstampa = app.activeDocument.width;
  var nomeBase;

  if (larguraEstampa == 5896) {
    nomeBase = "Prata";
  } else {
    if (larguraEstampa == 4124) {
      nomeBase = "Pontos";
    } else {
      if (larguraEstampa == 8249) {
        nomeBase = "Pontos";
      } else {
        if (larguraEstampa == 3827) {
          nomeBase = "Linho";
        } else {
          if (larguraEstampa == 3855) {
            nomeBase = "Seda";
          } else {
            if (larguraEstampa >= 5272 && larguraEstampa <= 5273) {
              nomeBase = "Palha";
            } else {
              if (larguraEstampa >= 4506 && larguraEstampa <= 4508) {
                nomeBase = "Jateado";
              } else {
                if (larguraEstampa >= 6008 && larguraEstampa <= 6010) {
                  nomeBase = "Jateado";
                } else {
                  if (larguraEstampa == 7512) {
                    nomeBase = "Jateado";
                  } else {
                    if (larguraEstampa >= 9014 && larguraEstampa <= 9015) {
                      nomeBase = "Jateado";
                    } else {
                      if (larguraEstampa >= 12018 && larguraEstampa <= 12019) {
                        nomeBase = "Jateado";
                      } else {
                        if (
                          larguraEstampa >= 10516 &&
                          larguraEstampa <= 10517
                        ) {
                          nomeBase = "Jateado";
                        } else {
                          if (
                            larguraEstampa >= 1501 &&
                            larguraEstampa <= 1503
                          ) {
                            nomeBase = "Jateado";
                          } else {
                            if (
                              larguraEstampa >= 3797 &&
                              larguraEstampa <= 3799
                            ) {
                              nomeBase = "Vinil";
                            } else {
                              if (
                                larguraEstampa >= 7596 &&
                                larguraEstampa <= 7598
                              ) {
                                nomeBase = "Vinil";
                              } else {
                                if (
                                  larguraEstampa >= 11394 &&
                                  larguraEstampa <= 11396
                                ) {
                                  nomeBase = "Vinil";
                                } else {
                                  if (
                                    larguraEstampa >= 7908 &&
                                    larguraEstampa <= 7910
                                  ) {
                                    nomeBase = "Palha";
                                  } else {
                                    nomeBase = "";
                                  }
                                }
                              }
                            }
                          }
                        }
                      }
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
  }

  return nomeBase;
}

// Redimensiona o arquivo aberto para a largura correta - ou altura, caso a maior dimensão do arquivo seja;
// Indice ---> input vai ser dado pelo FOR
// Cria uma seleção no arquivo inteiro
// Copia tudo para a Área de seleção
// Fecha e volta para a página aberta
// Cola o arquivo copiado na página
// Adiciona o texto com o nome + informações extras definidas na função "ChecaBase"
// Insere o texto por conta do escopo da abertura do arquivo
function ArquivoHANDLER(indice, posXCelula, posYCelula) {
  
  var nomeBase = ChecaBase();

  app.activeDocument.resizeImage(largura / NUMColunas - margem);

  if (app.activeDocument.width / app.activeDocument.height < 0.9) {
    app.activeDocument.resizeImage(undefined, altura / NUMLinhas - margem * 2);
  }

  app.activeDocument.selection.selectAll();
  app.activeDocument.flatten();

  var nome = app.activeDocument.name;

  app.activeDocument.selection.copy();
  app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);

  var Celula = posCelula(posXCelula, posYCelula);

  app.activeDocument.selection.select(Celula, SelectionType.REPLACE, 0, false);

  app.activeDocument.paste(true);

  var formulaNomeCamada = "Camada " + indice;

  var nomeCamada = formulaNomeCamada.toString();

  SelecionaMascara();
  ExcluiMascara();

  var camada = app.activeDocument.artLayers.getByName(nomeCamada);
  camada.applyOffset(0, -30, OffsetUndefinedAreas.SETTOBACKGROUND);

  FazTexto(nome, nomeCamada, nomeBase);
}

// Cria uma nova página A4
function CriaDoc(indice) {
  var novoDoc = app.documents.add(largura, altura, 300, "PaginaBoard" + indice);
}

// Entende a largura e altura das células da imagem
function PegaColuna() {
  var LarguraColuna = largura / colunas.length;
  var AlturaLinha = altura / linhas.length;

  return LarguraColuna, AlturaLinha;
}

//Cria o grid na página
function FazGrid(colunas, linhas) {
  var idnewGuideLayout = stringIDToTypeID("newGuideLayout");
  var desc178 = new ActionDescriptor();
  var idpresetKind = stringIDToTypeID("presetKind");
  var idpresetKindType = stringIDToTypeID("presetKindType");
  var idpresetKindCustom = stringIDToTypeID("presetKindCustom");
  desc178.putEnumerated(idpresetKind, idpresetKindType, idpresetKindCustom);
  var idguideLayout = stringIDToTypeID("guideLayout");
  var desc179 = new ActionDescriptor();
  var idcolCount = stringIDToTypeID("colCount");
  desc179.putInteger(idcolCount, colunas);
  var idrowCount = stringIDToTypeID("rowCount");
  desc179.putInteger(idrowCount, linhas);
  var idmarginTop = stringIDToTypeID("marginTop");
  var idPxl = charIDToTypeID("#Pxl");
  desc179.putUnitDouble(idmarginTop, idPxl, 0.0);
  var idmarginLeft = stringIDToTypeID("marginLeft");
  var idPxl = charIDToTypeID("#Pxl");
  desc179.putUnitDouble(idmarginLeft, idPxl, 0.0);
  var idmarginBottom = stringIDToTypeID("marginBottom");
  var idPxl = charIDToTypeID("#Pxl");
  desc179.putUnitDouble(idmarginBottom, idPxl, 0.0);
  var idmarginRight = stringIDToTypeID("marginRight");
  var idPxl = charIDToTypeID("#Pxl");
  desc179.putUnitDouble(idmarginRight, idPxl, 0.0);
  var idguideLayout = stringIDToTypeID("guideLayout");
  desc178.putObject(idguideLayout, idguideLayout, desc179);
  var idguideTarget = stringIDToTypeID("guideTarget");
  var idguideTarget = stringIDToTypeID("guideTarget");
  var idguideTargetCanvas = stringIDToTypeID("guideTargetCanvas");
  desc178.putEnumerated(idguideTarget, idguideTarget, idguideTargetCanvas);
  executeAction(idnewGuideLayout, desc178, DialogModes.NO);
}

// Pega a posição da celula
function posCelula(indiceX, indiceY) {
  posicao = [];
  CelulaBounds1 = [
    Math.round(LarguraColuna * indiceX),
    Math.round(AlturaColuna * indiceY),
  ];
  CelulaBounds2 = [
    Math.round(LarguraColuna * indiceX),
    Math.round(AlturaColuna * indiceY + AlturaColuna),
  ];
  CelulaBounds3 = [
    Math.round(LarguraColuna * indiceX) + LarguraColuna,
    Math.round(AlturaColuna * indiceY + AlturaColuna),
  ];
  CelulaBounds4 = [
    Math.round(LarguraColuna * indiceX) + LarguraColuna,
    Math.round(AlturaColuna * indiceY),
  ];

  posicao.push(CelulaBounds1);
  posicao.push(CelulaBounds2);
  posicao.push(CelulaBounds3);
  posicao.push(CelulaBounds4);

  return posicao;
}

// Seleciona a Mascara que o Photoshop cria automaticamente
function SelecionaMascara() {
  var idslct = charIDToTypeID("slct");
  var desc2188 = new ActionDescriptor();
  var idnull = charIDToTypeID("null");
  var ref708 = new ActionReference();
  var idChnl = charIDToTypeID("Chnl");
  var idChnl = charIDToTypeID("Chnl");
  var idMsk = charIDToTypeID("Msk ");
  ref708.putEnumerated(idChnl, idChnl, idMsk);
  desc2188.putReference(idnull, ref708);
  var idMkVs = charIDToTypeID("MkVs");
  desc2188.putBoolean(idMkVs, false);
  executeAction(idslct, desc2188, DialogModes.NO);
}

// Exclui a mascara que o Photoshop cria automaticamente
function ExcluiMascara() {
  var idDlt = charIDToTypeID("Dlt ");
  var desc2186 = new ActionDescriptor();
  var idnull = charIDToTypeID("null");
  var ref706 = new ActionReference();
  var idChnl = charIDToTypeID("Chnl");
  var idOrdn = charIDToTypeID("Ordn");
  var idTrgt = charIDToTypeID("Trgt");
  ref706.putEnumerated(idChnl, idOrdn, idTrgt);
  desc2186.putReference(idnull, ref706);
  executeAction(idDlt, desc2186, DialogModes.NO);
}

function DuplicaProInicio(nomeDoc) {
  var idDplc = charIDToTypeID("Dplc");
  var desc15350 = new ActionDescriptor();
  var idnull = charIDToTypeID("null");
  var ref1141 = new ActionReference();
  var idLyr = charIDToTypeID("Lyr ");
  var idOrdn = charIDToTypeID("Ordn");
  var idTrgt = charIDToTypeID("Trgt");
  ref1141.putEnumerated(idLyr, idOrdn, idTrgt);
  desc15350.putReference(idnull, ref1141);
  var idT = charIDToTypeID("T   ");
  var ref1142 = new ActionReference();
  var idDcmn = charIDToTypeID("Dcmn");
  ref1142.putName(idDcmn, "PaginaBoard0");
  desc15350.putReference(idT, ref1142);
  var idNm = charIDToTypeID("Nm  ");
  desc15350.putString(idNm, nomeDoc);
  var idVrsn = charIDToTypeID("Vrsn");
  desc15350.putInteger(idVrsn, 5);
  executeAction(idDplc, desc15350, DialogModes.NO);
}

// Importa molde da página
// Sempre lembrar de colocar o caminho correto para o arquivo do molde da página
function Importa(caminho) {
  var idPlc = charIDToTypeID("Plc ");
  var desc15820 = new ActionDescriptor();
  var idIdnt = charIDToTypeID("Idnt");
  desc15820.putInteger(idIdnt, 40);
  var idnull = charIDToTypeID("null");
  desc15820.putPath(idnull, new File(caminho));
  var idFTcs = charIDToTypeID("FTcs");
  var idQCSt = charIDToTypeID("QCSt");
  var idQcsa = charIDToTypeID("Qcsa");
  desc15820.putEnumerated(idFTcs, idQCSt, idQcsa);
  var idOfst = charIDToTypeID("Ofst");
  var desc15821 = new ActionDescriptor();
  var idHrzn = charIDToTypeID("Hrzn");
  var idPxl = charIDToTypeID("#Pxl");
  desc15821.putUnitDouble(idHrzn, idPxl, 0.0);
  var idVrtc = charIDToTypeID("Vrtc");
  var idPxl = charIDToTypeID("#Pxl");
  desc15821.putUnitDouble(idVrtc, idPxl, 0.0);
  var idOfst = charIDToTypeID("Ofst");
  desc15820.putObject(idOfst, idOfst, desc15821);
  executeAction(idPlc, desc15820, DialogModes.NO);
}

//Rasteriza a layer atual
function Rasteriza() {
  var idrasterizeLayer = stringIDToTypeID("rasterizeLayer");
  var desc11701 = new ActionDescriptor();
  var idnull = charIDToTypeID("null");
  var ref2313 = new ActionReference();
  var idLyr = charIDToTypeID("Lyr ");
  var idOrdn = charIDToTypeID("Ordn");
  var idTrgt = charIDToTypeID("Trgt");
  ref2313.putEnumerated(idLyr, idOrdn, idTrgt);
  desc11701.putReference(idnull, ref2313);
  executeAction(idrasterizeLayer, desc11701, DialogModes.NO);
}

//Move a camada para a posição acima da página base para incorporar em todas as páginas
function MOVE(Pos) {
  var idmove = charIDToTypeID("move");
  var desc16925 = new ActionDescriptor();
  var idnull = charIDToTypeID("null");
  var ref1390 = new ActionReference();
  var idLyr = charIDToTypeID("Lyr ");
  var idOrdn = charIDToTypeID("Ordn");
  var idTrgt = charIDToTypeID("Trgt");
  ref1390.putEnumerated(idLyr, idOrdn, idTrgt);
  desc16925.putReference(idnull, ref1390);
  var idT = charIDToTypeID("T   ");
  var ref1391 = new ActionReference();
  var idLyr = charIDToTypeID("Lyr ");
  ref1391.putIndex(idLyr, Pos);
  desc16925.putReference(idT, ref1391);
  var idAdjs = charIDToTypeID("Adjs");
  desc16925.putBoolean(idAdjs, false);
  var idVrsn = charIDToTypeID("Vrsn");
  desc16925.putInteger(idVrsn, 5);
  var idLyrI = charIDToTypeID("LyrI");
  var list446 = new ActionList();
  list446.putInteger(36);
  desc16925.putList(idLyrI, list446);
  executeAction(idmove, desc16925, DialogModes.NO);
}

//Insere um texto com base numa caixa de diálogo que você escreve o quer que entre
function Cabecalho(tipo, X, Y) {
  var input = prompt("Insira o nome do(a)" + tipo, " ");

  var camadaTexto = app.activeDocument.artLayers.add();
  camadaTexto.kind = LayerKind.TEXT;
  camadaTexto.name = tipo;

  var conteudo = camadaTexto.textItem;
  conteudo.contents = input;

  conteudo.position = new Array(X, Y);
  conteudo.font = arial;
  conteudo.size = 18;
}

//Chama a criação de um texto usando o que vai ser escrito e a posição a ser colocada no canvas
function ChamaTexto(text, X, Y) {
  var camadaTexto = app.activeDocument.artLayers.add();
  camadaTexto.kind = LayerKind.TEXT;
  camadaTexto.name = "Contador de artes";

  var conteudo = camadaTexto.textItem;
  conteudo.contents = text;

  conteudo.position = new Array(X, Y);
  conteudo.font = arial;
  conteudo.size = 18;
}

//Transforma a camada alvo em uma prancheta
//Recebe o ID que é o nome da camada que vai virar a prancheta - Note que as dimensões da prancheta viram as dimensões dessa camada base
function FazPrancheta(ID) {
  var idMk = charIDToTypeID("Mk  ");
  var desc17011 = new ActionDescriptor();
  var idnull = charIDToTypeID("null");
  var ref1421 = new ActionReference();
  var idartboardSection = stringIDToTypeID("artboardSection");
  ref1421.putClass(idartboardSection);
  desc17011.putReference(idnull, ref1421);
  var idFrom = charIDToTypeID("From");
  var ref1422 = new ActionReference();
  var idLyr = charIDToTypeID("Lyr ");
  var idOrdn = charIDToTypeID("Ordn");
  var idTrgt = charIDToTypeID("Trgt");
  ref1422.putEnumerated(idLyr, idOrdn, idTrgt);
  desc17011.putReference(idFrom, ref1422);
  var idUsng = charIDToTypeID("Usng");
  var desc17012 = new ActionDescriptor();
  var idNm = charIDToTypeID("Nm  ");
  desc17012.putString(idNm, ID);
  var idartboardSection = stringIDToTypeID("artboardSection");
  desc17011.putObject(idUsng, idartboardSection, desc17012);
  var idlayerSectionStart = stringIDToTypeID("layerSectionStart");
  desc17011.putInteger(idlayerSectionStart, 68);
  var idlayerSectionEnd = stringIDToTypeID("layerSectionEnd");
  desc17011.putInteger(idlayerSectionEnd, 69);
  var idNm = charIDToTypeID("Nm  ");
  desc17011.putString(idNm, ID);
  executeAction(idMk, desc17011, DialogModes.NO);
}

/////////////////////////////////// ROTINA DE EXECUÇÃO ////////////////////////////////

// Retorna um Array de caminhos das estampas da pasta
var PastaEstampas = Folder.selectDialog("Selecione");
if (PastaEstampas != null) {
  var listaArquivos = PastaEstampas.getFiles();
} else {
  alert("Pasta vazia!");
}

//Faz o número de páginas ( adiciona um para dar conta do resto da divisão arredondada pra baixo)
var NUMPaginas =
  Math.floor(listaArquivos.length / (NUMColunas * NUMLinhas)) + 1;

///////MOTOR

//FATOR DE POSICAO
var posColuna = 0;
var posLinha = 0;

// Iterador que marca o index da lista de arquivos na inserção das imagens
var o = 0;

// Marcador de página
var indice;

// Iterador de páginas
var PagID = 0;

while (PagID < NUMPaginas) {
  for (o = 0; o < listaArquivos.length; o++) {
    if (o == NUMColunas * NUMLinhas * PagID) {
      posColuna = 0;
      posLinha = 0;
      indice = 1;
      try {
        app.activeDocument.flatten();
      } catch (e) {}
      CriaDoc(PagID);
      PagID++;
    }
    Abre(listaArquivos[o]);
    ArquivoHANDLER(indice, posColuna, posLinha);
    indice++;
    posColuna++;
    if (posColuna == NUMColunas) {
      posLinha++;
      posColuna = 0;
    }
  }
}

app.activeDocument.flatten();

var TodosDocumentos = app.documents.length;

for (n = 1; n < TodosDocumentos; n++) {
  app.activeDocument = app.documents[n];
  DuplicaProInicio("Pagina " + n);
}

for (m = 1; m < TodosDocumentos; m++) {
  app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);
}

var camada = app.activeDocument.artLayers[PagID - 1];

camada.isBackgroundLayer = false;
app.activeDocument.resizeCanvas(
  undefined,
  TamanhoA4Total,
  AnchorPosition.BOTTOMCENTER
);
camada.name = "Pagina 0";

app.activeDocument.resizeCanvas(
  undefined,
  TamanhoA4Total * PagID,
  AnchorPosition.TOPCENTER
);

for (j = 1; j < PagID; j++) {
  var camadita = app.activeDocument.artLayers.getByName("Pagina " + j);
  camadita.applyOffset(
    0,
    TamanhoA4Total * j,
    OffsetUndefinedAreas.SETTOBACKGROUND
  );
}

//Parte final do script em que produz as pranchetas à serem exportadas em PDF

Importa(
  "D:\\Creative Cloud\\Creative Cloud Files\\Branco\\PaginaBoard_Contrato.tif"
);

Rasteriza();

app.activeDocument.activeLayer.applyOffset(
  0,
  ((TamanhoA4Total - app.activeDocument.height) / 2) * -1,
  OffsetUndefinedAreas.SETTOBACKGROUND
);
app.activeDocument.activeLayer.blendMode = BlendMode.MULTIPLY;

ChamaTexto("artes: " + listaArquivos.length, 1420, 126);

Importa("D:\\Creative Cloud\\Creative Cloud Files\\Branco\\LOGO_BRANCO.psd");

Rasteriza();

MOVE(PagID + 1);

app.activeDocument.activeLayer.applyOffset(
  924,
  (app.activeDocument.height / 2 - 100) * -1,
  OffsetUndefinedAreas.SETTOBACKGROUND
);

Cabecalho("Artista", 422, 126);

Rasteriza();

MOVE(PagID + 1);

app.activeDocument.activeLayer.merge();

for (x = 1; x < PagID; x++) {
  app.activeDocument.activeLayer.name = "Pag" + (x - 1);
  app.activeDocument.activeLayer.duplicate();
  app.activeDocument.activeLayer.applyOffset(
    0,
    TamanhoA4Total,
    OffsetUndefinedAreas.SETTOBACKGROUND
  );
  app.activeDocument.activeLayer.name = "Pag" + x;
}

app.activeDocument.activeLayer.name =
  app.activeDocument.activeLayer.name + " copiar";

for (y = 0; y < PagID; y++) {
  app.activeDocument.activeLayer = app.activeDocument.artLayers.getByName(
    "Pag" + y + " copiar"
  );
  FazPrancheta("Pag " + y);
}

app.preferences.rulerUnits = startRulerUnits;
