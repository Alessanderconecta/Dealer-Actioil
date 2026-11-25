const ABA_VISITAS = "Visitas";
const ABA_MARCAS = "Marcas";
const ABA_REPS = "Representantes";
const ABA_EMPRESAS = "Empresas";

const COLUNAS_VISITAS = [
  "ID", "Registrado em", "Data", "Representante", "Regional", "UF", "Cidade", "Segmento", "Marca",
  "Grupo econômico", "Matriz/Filial", "ID ponto de venda", "Tipo da visita", "Valor da venda",
  "Pessoa de contato", "Telefone", "Descrição"
];
const COLUNAS_MARCAS = ["Segmento", "Marca", "Ativo"];
const COLUNAS_REPS = ["Nome", "Regional", "Ativo"];
const COLUNAS_EMPRESAS = ["ID", "Grupo econômico", "Segmento", "Marca", "UF", "Cidade", "Matriz/Filial", "Ativo"];

const SEGMENTOS = ["Caminhões", "Máquinas"];
const UFS = ["AC","AL","AM","AP","BA","CE","DF","ES","GO","MA","MG","MS","MT","PA","PB","PE","PI","PR","RJ","RN","RO","RR","RS","SC","SE","SP","TO"];
const MARCAS_INICIAIS = {
  "Caminhões": ["VW Truck e Bus", "Mercedes-Benz", "Volvo", "Scania", "Iveco", "DAF", "Foton", "Agrale", "JAC", "Ford",
    "Sany", "Nanjing", "Hyundai", "Tesla", "BYD", "XCMG", "Sinotruk", "CNHTC", "International", "Peterbilt"],
  "Máquinas": ["John Deere", "Massey Ferguson", "Valtra", "Fendt", "New Holland", "Case IH", "Agrale", "Agritech",
    "Mahindra", "LS Tractor", "Yanmar", "Kioti", "Landini", "Jacto", "Tatu Marchesan", "Horsch"]
};

// Rode uma vez (menu Executar > prepararPlanilha) para criar as abas, cabeçalhos e listas suspensas.
function prepararPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const marcas = abaOuCria(ss, ABA_MARCAS, COLUNAS_MARCAS);
  if (marcas.getLastRow() < 2) {
    const linhas = [];
    SEGMENTOS.forEach(seg => MARCAS_INICIAIS[seg].forEach(m => linhas.push([seg, m, "SIM"])));
    marcas.getRange(2, 1, linhas.length, 3).setValues(linhas);
  }
  lista(marcas, 1, SEGMENTOS);
  lista(marcas, 3, ["SIM", "NÃO"]);

  const reps = abaOuCria(ss, ABA_REPS, COLUNAS_REPS);
  lista(reps, 2, ["1", "2", "3", "4"]);
  lista(reps, 3, ["SIM", "NÃO"]);

  const emp = abaOuCria(ss, ABA_EMPRESAS, COLUNAS_EMPRESAS);
  lista(emp, 3, SEGMENTOS);
  lista(emp, 5, UFS);
  lista(emp, 7, ["Matriz", "Filial"]);
  lista(emp, 8, ["SIM", "NÃO"]);

  abaOuCria(ss, ABA_VISITAS, COLUNAS_VISITAS);
}

function abaOuCria(ss, nome, colunas) {
  let sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, colunas.length).setValues([colunas]).setFontWeight("bold").setBackground("#e8eefb");
    sh.setFrozenRows(1);
  }
  return sh;
}

function lista(sh, coluna, valores) {
  const regra = SpreadsheetApp.newDataValidation().requireValueInList(valores, true).setAllowInvalid(false).build();
  sh.getRange(2, coluna, 4999, 1).setDataValidation(regra);
}

function paraData(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ""));
  return m ? Utilities.parseDate(String(s), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd") : seguro(s);
}

function seguro(v) {
  const s = String(v == null ? "" : v).trim();
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const v = JSON.parse(e.postData.contents);
    if (!v.id) return resposta({ ok: false, erro: "sem id" });

    const sh = abaOuCria(SpreadsheetApp.getActiveSpreadsheet(), ABA_VISITAS, COLUNAS_VISITAS);
    const n = sh.getLastRow() - 1;
    if (n > 0) {
      const ids = sh.getRange(2, 1, n, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (ids[i][0] === v.id) return resposta({ ok: true, duplicado: true });
      }
    }

    const linha = sh.getLastRow() + 1;
    const colTel = COLUNAS_VISITAS.indexOf("Telefone") + 1;
    const colValor = COLUNAS_VISITAS.indexOf("Valor da venda") + 1;
    sh.getRange(linha, colTel).setNumberFormat("@");
    sh.getRange(linha, colValor).setNumberFormat("R$ #,##0.00");
    sh.getRange(linha, 2).setNumberFormat("dd/MM/yyyy HH:mm");
    sh.getRange(linha, 3).setNumberFormat("dd/MM/yyyy");
    sh.getRange(linha, 1, 1, COLUNAS_VISITAS.length).setValues([[
      v.id,
      new Date(),
      paraData(v.data),
      seguro(v.representante),
      seguro(v.regional),
      seguro(v.uf),
      seguro(v.cidade),
      seguro(v.segmento),
      seguro(v.marca),
      seguro(v.grupo),
      seguro(v.matrizFilial),
      seguro(v.pdvId),
      seguro(v.tipo),
      v.valor === "" || v.valor == null ? "" : Number(v.valor),
      seguro(v.contato),
      seguro(v.telefone),
      seguro(v.descricao)
    ]]);
    return resposta({ ok: true });
  } catch (err) {
    return resposta({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// O app só lê as listas-base. A aba Visitas nunca é devolvida por aqui.
function doGet(e) {
  if (e && e.parameter && e.parameter.action === "config") return resposta(configuracao());
  return resposta({ ok: true, app: "dealer-actioil" });
}

function lerAba(nome, nCols) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, nCols).getValues();
}

const ativo = v => !/^N[ÃA]O$/i.test(String(v || "").trim());
const txt = v => String(v == null ? "" : v).trim();

function configuracao() {
  const marcas = lerAba(ABA_MARCAS, 3)
    .filter(r => txt(r[0]) && txt(r[1]) && ativo(r[2]))
    .map(r => ({ segmento: txt(r[0]), marca: txt(r[1]) }));

  const representantes = lerAba(ABA_REPS, 3)
    .filter(r => txt(r[0]) && ativo(r[2]))
    .map(r => ({ nome: txt(r[0]), regional: txt(r[1]) }));

  const empresas = [];
  lerAba(ABA_EMPRESAS, 8).forEach((r, i) => {
    if (!txt(r[1]) || !txt(r[2]) || !txt(r[3]) || !txt(r[4]) || !txt(r[5]) || !ativo(r[7])) return;
    empresas.push({
      id: txt(r[0]) || "L" + (i + 2),
      grupo: txt(r[1]),
      segmento: txt(r[2]),
      marca: txt(r[3]),
      uf: txt(r[4]).toUpperCase(),
      cidade: txt(r[5]),
      mf: txt(r[6]) || "A confirmar"
    });
  });

  return { ok: true, marcas, representantes, empresas };
}

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
