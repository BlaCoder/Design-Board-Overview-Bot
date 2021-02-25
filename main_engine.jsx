// Code created by Tiago Costa de Carvalho
// Date of creation: 10/2020
//
// Within de great help from Adobe Community from little tweeks from other codes and doubts :)
//
// This code works like "Contact Sheet II" but it works in a more straightfoward method were it create the layout for a normal A4 print. 
// It do not allow further editing in Photoshop interface besides a name that goes at the first page. ( or it could go blank anyways );

// It uses the same idea of creating a series of cells based on the size of the document. As it follows, it opens a prompt asking for the path where the images are.

// Pages are all generated automatically and at the end, it joins all the pages to the initial document, puts that name of the designer ( or blank if wanted )
// and also import a model sheet and a logo if is desired to make the overview board more appropriated for a business to use; 
// TODO: use a "import" command for calling the build-in script for exporting all worktables to pdf

#target photoshop

var startRulerUnits = app.preferences.rulerUnits;

app.preferences.rulerUnits = Units.PIXELS;

/////////////////////////////////////GLOBAL VARIABLES//////////////////////////////////////////
// note: variables will go all in portuguese name

// A4 sheet dimensions in pixels ( at 300 dpi )
var TamanhoA4 = [2480, 3308];
var MM300DPI = 12;

var largura = TamanhoA4[0];
var altura = TamanhoA4[1];
var margem = MM300DPI * 3;

var TamanhoA4Total = altura + 200;

// Number of images per page
var NUMColunas = 3;
var NUMLinhas = 3;

var LarguraColuna = largura / NUMColunas;
var AlturaColuna = altura / NUMLinhas;

var colunas = [];
var linhas = [];

var texto;
var nome;

var arial = app.fonts.getByName("ArialMT");

//Naming columns for further access
for (i = 0; i < NUMColunas; i++) {
  var texto = "col" + [i];
  colunas.push(texto);
}

//Naming lines for further access
for (i = 0; i < NUMLinhas; i++) {
  var texto = "lin" + [i];
  colunas.push(texto);
}

//////////////////////////////////////FUNCTIONS///////////////////////////////////////////////////

// Opens a document within a path
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

// Creates a full layer of text using the opened document's name, the inputs are: name of the document, name of the text layer and the addictional info based in some aspect of the image. 
// note: In this version, it uses the size in pixels to check the width of the image to categorize in what thing are going to be printed; 
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

//Adds depending info to the name that goes under the image depending on the width of the opened image.
//Returns the name to the function "FazTexto" uses it.
function ChecaBase() {

  app.preferences.rulerUnits = Units.PIXELS;

  var larguraEstampa = app.activeDocument.width;
  var nomeBase;

  if(larguraEstampa =! 0) {
    nomeBase = ""
  }

  return nomeBase;
}

// The image handler opens the image, make the necessary transformations to fit in the overview board
// Indice ---> input important for the engine to work
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

// Create an A4 sheet
function CriaDoc(indice) {
  var novoDoc = app.documents.add(largura, altura, 300, "PaginaBoard" + indice);
}

// Calculates the size of lines and columns
function PegaColuna() {
  var LarguraColuna = largura / colunas.length;
  var AlturaLinha = altura / linhas.length;

  return LarguraColuna, AlturaLinha;
}

//Creates a grid
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

// Get the position of a cell
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

// The following two functions works like a "prevent default" to the auto clipping mask that photoshop create when you paste an image in a predefined selection. This is made because images were invisible when pasted.
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

//Duplicate an intire page to the initial document
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

// Imports the page model
// Check for the right path for the page model
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

//Raster any special object or text layer
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

//Reorganize the layer in the layer set;
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

//prompt for inserting the name in de overview board
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

//Create any kind of text and gives its position
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

//Turns a layer in to a worktable set
//This is importat cause is easy to export worktables to PDF and when created using a layer, the dimensions always goes correct for print.
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

/////////////////////////////////// RUN TIME ////////////////////////////////

// Returns an array of file paths
var PastaEstampas = Folder.selectDialog("Selecione");
if (PastaEstampas != null) {
  var listaArquivos = PastaEstampas.getFiles();
} else {
  alert("Empty folder!");
}

//Determines the number of times that engine will work
var NUMPaginas =
  Math.floor(listaArquivos.length / (NUMColunas * NUMLinhas)) + 1;

///////ENGINE

//Position factor
var posColuna = 0;
var posLinha = 0;

// Index of image in a page
var o = 0;

// Page index
var indice;

// Identify the page
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

//TODO: there's a bug if the number of images on the entire set if perfect divisible by the product of "NUMColunas" and "NUMLinhas" its call an error message and breaks at this point.

/////////////////////// END OF ENGINE //////////////////////////

//////////////////////
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

//Final part where worktables are made

Importa(
  // ************************************************PATH TO THE MODEL SHEET************************************************************
);

Rasteriza();

app.activeDocument.activeLayer.applyOffset(
  0,
  ((TamanhoA4Total - app.activeDocument.height) / 2) * -1,
  OffsetUndefinedAreas.SETTOBACKGROUND
);
app.activeDocument.activeLayer.blendMode = BlendMode.MULTIPLY;

ChamaTexto("artes: " + listaArquivos.length, 1420, 126);

Importa(
  // ************************************************PATH TO A LOGO************************************************************
  // Charge this path with a blank white image if not wanted;
);

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
