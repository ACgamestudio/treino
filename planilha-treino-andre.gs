/**
 * Treino do André — guarda o desempenho numa planilha do Google
 *
 * COMO INSTALAR (5 minutos, só se faz uma vez)
 *
 * 1. Abra sheets.new e dê um nome à planilha, tipo "Treino do André".
 * 2. Menu Extensões → Apps Script.
 * 3. Apague o código que aparecer e cole TODO este arquivo no lugar.
 * 4. Clique no disquete pra salvar.
 * 5. Botão azul "Implantar" → "Nova implantação".
 *      - No engrenagem ao lado de "Selecionar tipo", escolha "App da Web".
 *      - Executar como: Eu (seu e-mail).
 *      - Quem pode acessar: QUALQUER PESSOA.   ← isso é obrigatório
 *      - Clique em Implantar e autorize (vai aparecer um aviso do Google:
 *        "Avançado" → "Acessar projeto sem título (não seguro)" → Permitir.
 *        É seguro, o script é seu e só mexe na sua planilha).
 * 6. Copie o "URL do app da Web" (termina em /exec).
 * 7. No app: engrenagem → Planilha do Google → cole o link → Testar conexão.
 *
 * As abas Treinos, Séries e Cargas são criadas sozinhas no primeiro envio.
 *
 * MUDOU O CÓDIGO DEPOIS? Implantar → Gerenciar implantações → lápis →
 * Versão: Nova versão → Implantar. O link continua o mesmo.
 */

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);

    if (d.tipo === 'treino') {
      gravarTreino(d);
    } else if (d.tipo === 'cargas') {
      gravarCargas(d.dados);
    } else if (d.tipo === 'lote') {
      (d.treinos || []).forEach(gravarTreino);
      if (d.cargas) gravarCargas(d.cargas);
    } else {
      return resposta({ ok: false, erro: 'tipo desconhecido: ' + d.tipo });
    }
    return resposta({ ok: true });

  } catch (err) {
    return resposta({ ok: false, erro: String(err) });
  }
}

function doGet(e) {
  try {
    var tipo = (e && e.parameter && e.parameter.tipo) || 'ping';

    if (tipo === 'estado') {
      return resposta({ ok: true, cargas: lerCargas(), treinos: lerTreinos() });
    }
    return resposta({
      ok: true,
      planilha: SpreadsheetApp.getActiveSpreadsheet().getName()
    });

  } catch (err) {
    return resposta({ ok: false, erro: String(err) });
  }
}

/* ---------------------------------------------------------------- gravar */

function gravarTreino(d) {
  var t = aba('Treinos', ['Data', 'Treino', 'Foco', 'Séries', 'Volume (kg)', 'Duração (min)']);

  // não duplica o mesmo treino no mesmo dia se você mandar de novo
  var vals = t.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (texto(vals[i][0]) === d.data && String(vals[i][1]) === String(d.treino)) return;
  }

  t.appendRow([d.data, d.treino, d.foco || '', d.series_total || 0, d.volume || 0, d.min || 0]);

  var s = aba('Séries', ['Data', 'Treino', 'Exercício', 'Série', 'Peso (kg)', 'Reps']);
  var linhas = (d.series || []).map(function (x) {
    return [d.data, d.treino, x.ex, x.serie, x.peso === '' ? '' : Number(x.peso), x.reps];
  });
  if (linhas.length) {
    s.getRange(s.getLastRow() + 1, 1, linhas.length, 6).setValues(linhas);
  }
}

function gravarCargas(dados) {
  var s = aba('Cargas', ['ID', 'Exercício', 'Série', 'Peso (kg)', 'Atualizado']);

  if (s.getLastRow() > 1) {
    s.getRange(2, 1, s.getLastRow() - 1, 5).clearContent();
  }

  var agora = new Date(), linhas = [];
  Object.keys(dados || {}).forEach(function (id) {
    var ex = dados[id];
    Object.keys(ex.series || {}).forEach(function (i) {
      linhas.push([id, ex.nome, Number(i) + 1, Number(ex.series[i]), agora]);
    });
  });

  if (linhas.length) {
    linhas.sort(function (a, b) { return a[0] === b[0] ? a[2] - b[2] : (a[0] < b[0] ? -1 : 1); });
    s.getRange(2, 1, linhas.length, 5).setValues(linhas);
  }
}

/* ------------------------------------------------------------------ ler */

function lerCargas() {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cargas');
  if (!s || s.getLastRow() < 2) return {};

  var vals = s.getRange(2, 1, s.getLastRow() - 1, 4).getValues(), out = {};
  vals.forEach(function (r) {
    var id = String(r[0]).trim();
    if (!id) return;
    out[id] = out[id] || {};
    out[id][Number(r[2]) - 1] = Number(r[3]);
  });
  return out;
}

function lerTreinos() {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Treinos');
  if (!s || s.getLastRow() < 2) return [];

  var vals = s.getRange(2, 1, s.getLastRow() - 1, 6).getValues();
  var out = vals.filter(function (r) { return r[0]; }).map(function (r) {
    return { d: texto(r[0]), w: String(r[1]), sets: Number(r[3]) || 0, vol: Number(r[4]) || 0, min: Number(r[5]) || 0 };
  });
  out.sort(function (a, b) { return a.d < b.d ? 1 : -1; });   // mais recente primeiro
  return out;
}

/* -------------------------------------------------------------- ajudinha */

function aba(nome, cabecalho) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(nome);
  if (!s) {
    s = ss.insertSheet(nome);
    s.appendRow(cabecalho);
    s.getRange(1, 1, 1, cabecalho.length).setFontWeight('bold').setBackground('#0F1B35').setFontColor('#E9F1FF');
    s.setFrozenRows(1);
    s.autoResizeColumns(1, cabecalho.length);
  }
  return s;
}

function texto(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v).trim();
}

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
