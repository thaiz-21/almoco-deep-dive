/**
 * Almoço Deep Dive – recebe os pedidos do site e grava nesta planilha.
 *
 * Aba "Pedidos": uma linha por pessoa (um pedido por nome; enviar de novo
 * com o mesmo nome substitui o pedido anterior).
 * Aba "Totais": quantidades por lanche, batata chips e bebidas (fórmulas).
 *
 * Como publicar (uma vez):
 *   Implantar → Nova implantação → Tipo: App da Web
 *   Executar como: Eu · Quem pode acessar: Qualquer pessoa
 *   Autorize e copie a URL do app da Web.
 */

const DEADLINE = new Date('2026-10-18T20:00:00-03:00'); // domingo, 20h (Brasília)
const MAX_TOTAL = 10; // lanches por pedido

const ITEMS = [
  ['burger-simples', 'Cheese Burguer Simples'],
  ['burger-classico', 'Cheese Salada Clássico 57 Anos'],
  ['burger-egg-salada', 'Cheese Egg Salada'],
  ['burger-magricelo', 'Magricelo'],
  ['frango-crispy', 'Frango Crispy'],
  ['frango-catupiry', 'Cheese Frango Catupiry'],
  ['calabresa-simples', 'Cheese Calabresa Simples'],
  ['calabresa-maionese', 'Cheese Calabresa Maionese'],
  ['calabresa-salada', 'Cheese Calabresa Salada'],
  ['calabresa-tartaro', 'Cheese Calabresa Tártaro'],
  ['calabresa-vinagrete', 'Cheese Calabresa Vinagrete'],
  ['calabresa-bacon', 'Cheese Calabresa Bacon'],
  ['calabresa-egg', 'Cheese Calabresa Egg'],
  ['calabresa-egg-salada', 'Cheese Calabresa Egg Salada'],
  ['calabresa-acebolado', 'Cheese Calabresa Acebolado'],
  ['calabresa-cheddar', 'Cheese Calabresa Cheddar'],
  ['calabresa-catupiry', 'Cheese Calabresa Catupiry'],
  ['calabresa-creme-milho', 'Cheese Calabresa Creme de Milho'],
  ['veg-cheese', 'Cheese Vegetariano'],
  ['veg-salada', 'Cheese Salada Vegetariano']
];
const ITEM_IDS = ITEMS.map(function (x) { return x[0]; });

const DRINKS = { coca: 'Coca-Cola', zero: 'Coca Zero', suco: 'Suco de laranja', nenhuma: 'Sem bebida' };

const BASE_HEADERS = ['Enviado em', 'Nome', 'Lanches', 'Total de lanches', 'Batata chips', 'Bebida', 'Observação', 'Chave'];
const KEY_COL = 8; // coluna "Chave" (oculta)

function doGet() {
  return out_({ ok: true, open: new Date() < DEADLINE });
}

function doPost(e) {
  var data;
  try { data = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'invalid' }); }
  if (new Date() >= DEADLINE) return out_({ ok: false, error: 'closed' });

  var name = String(data.name || '').replace(/\s+/g, ' ').trim().slice(0, 60);
  if (!name) return out_({ ok: false, error: 'name' });

  var qty = {}, total = 0;
  (Array.isArray(data.items) ? data.items : []).slice(0, 40).forEach(function (x) {
    if (!x || ITEM_IDS.indexOf(x.item) === -1) return;
    var q = Math.round(Number(x.qty)) || 0;
    if (q < 1) return;
    q = Math.min(q, MAX_TOTAL - total);
    if (q < 1) return;
    qty[x.item] = (qty[x.item] || 0) + q;
    total += q;
  });
  if (!total) return out_({ ok: false, error: 'items' });
  if (typeof data.chips !== 'boolean') return out_({ ok: false, error: 'chips' });
  if (!Object.prototype.hasOwnProperty.call(DRINKS, data.drink)) return out_({ ok: false, error: 'drink' });
  var note = String(data.note || '').trim().slice(0, 140);

  var lanches = ITEMS.filter(function (x) { return qty[x[0]]; })
    .map(function (x) { return qty[x[0]] + 'x ' + x[1]; }).join(', ');
  var key = keyOf_(name);
  var row = [new Date(), safe_(name), lanches, total, data.chips ? 'Sim' : 'Não', DRINKS[data.drink], safe_(note), key]
    .concat(ITEMS.map(function (x) { return qty[x[0]] || 0; }));

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = setup_();
    var last = sh.getLastRow();
    var updated = false;
    if (last > 1) {
      var keys = sh.getRange(2, KEY_COL, last - 1, 1).getValues();
      for (var i = 0; i < keys.length; i++) {
        if (keys[i][0] === key) {
          sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
          updated = true;
          break;
        }
      }
    }
    if (!updated) sh.appendRow(row);
    return out_({ ok: true, updated: updated });
  } finally {
    lock.releaseLock();
  }
}

/** Cria as abas "Pedidos" e "Totais" se ainda não existirem. Pode ser rodada à mão pelo editor. */
function setup_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Pedidos');
  if (!sh) {
    var sheets = ss.getSheets();
    sh = (sheets.length === 1 && sheets[0].getLastRow() === 0) ? sheets[0].setName('Pedidos') : ss.insertSheet('Pedidos');
    var headers = BASE_HEADERS.concat(ITEMS.map(function (x) { return x[1]; }));
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A2:A').setNumberFormat('dd/MM HH:mm');
    sh.hideColumns(KEY_COL);
  }
  if (!ss.getSheetByName('Totais')) buildTotals_(ss);
  return sh;
}

function setup() { setup_(); }

function buildTotals_(ss) {
  var t = ss.insertSheet('Totais');
  var rows = [['Lanche', 'Quantidade', '']];
  ITEMS.forEach(function (x, i) {
    var col = colLetter_(BASE_HEADERS.length + 1 + i);
    rows.push([x[1], '=SUM(Pedidos!' + col + '2:' + col + ')', '']);
  });
  rows.push(['Total de lanches', '=SUM(Pedidos!D2:D)', '']);
  rows.push(['', '', '']);
  rows.push(['Batata chips (compartilhada)', 'Pessoas', 'Porções inteiras (serve 2 a 3)']);
  var r = rows.length + 1;
  rows.push(['Querem batata', '=COUNTIF(Pedidos!E2:E,"Sim")', '=ROUNDUP(B' + r + '/2.5,0)']);
  rows.push(['', '', '']);
  rows.push(['Bebida (mercado)', 'Pessoas', 'Garrafas sugeridas (~600 ml por pessoa)']);
  [['Coca-Cola', 2], ['Coca Zero', 2], ['Suco de laranja', 1.5]].forEach(function (d) {
    var rr = rows.length + 1;
    rows.push([d[0] + (d[1] === 2 ? ' (2 L)' : ' (1,5 L)'), '=COUNTIF(Pedidos!F2:F,"' + d[0] + '")', '=ROUNDUP(B' + rr + '*0.6/' + d[1] + ',0)']);
  });
  rows.push(['Sem bebida', '=COUNTIF(Pedidos!F2:F,"Sem bebida")', '']);
  t.getRange(1, 1, rows.length, 3).setValues(rows);
  t.getRange(1, 1, 1, 3).setFontWeight('bold');
  [ITEMS.length + 2, ITEMS.length + 4, ITEMS.length + 7].forEach(function (row) {
    t.getRange(row, 1, 1, 3).setFontWeight('bold');
  });
  t.setColumnWidth(1, 280);
  t.setColumnWidth(3, 280);
  t.setFrozenRows(1);
}

function keyOf_(name) {
  return name.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
}

// Evita que um texto começando com =, +, - ou @ vire fórmula na planilha.
function safe_(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function colLetter_(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
