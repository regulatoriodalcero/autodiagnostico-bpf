/**
 * ============================================================================
 *  AUTODIAGN\u00d3STICO REGULAT\u00d3RIO BPF \u2014 Dal Cero Consultoria
 *  Backend (Google Apps Script Web App): e-mail (relat\u00f3rio no CORPO) + base de dados
 *
 *  A cada preenchimento, este script:
 *    1) envia o relat\u00f3rio por e-mail, no CORPO do e-mail (HTML), para a pessoa
 *       e para o time comercial (sem anexo PDF);
 *    2) grava os dados em uma Planilha Google (abas "Diagnosticos" e "Respostas").
 *
 *  PUBLICA\u00c7\u00c3O / ATUALIZA\u00c7\u00c3O (a URL /exec n\u00e3o muda ao publicar nova vers\u00e3o):
 *    - script.google.com -> seu projeto -> cole este arquivo por cima.
 *    - Rode testarEnvio() uma vez (autoriza e-mail + planilha) e confira.
 *    - Implantar -> Gerenciar implanta\u00e7\u00f5es -> editar (l\u00e1pis) -> Vers\u00e3o: Nova vers\u00e3o.
 * ============================================================================
 */

// ID da Planilha Google que vai guardar a base de dados (parte entre /d/ e /edit da URL)
var SHEET_ID = '1cxBFUQ1Vf50orfc86f11S50v2I3-Gww-3kIEzIwG0ic';

// Quem recebe a c\u00f3pia interna de cada autodiagn\u00f3stico (al\u00e9m da pr\u00f3pria pessoa)
var EMAILS_INTERNOS = ['comercial@dalceroconsultoria.com.br', 'regulatorio@dalceroconsultoria.com.br'];

var REMETENTE_NOME = 'Dal Cero Consultoria';

// Paleta da marca
var COR = {
  navy: '#041F47', azul: '#0043D8', azulClaro: '#177fe5', dourado: '#E8B461',
  cinza: '#6b7280', verde: '#10b981', amarelo: '#f59e0b', laranja: '#f97316', vermelho: '#ef4444'
};

function doPost(e) {
  var resultado = { ok: true, etapas: {} };
  var p = (e && e.parameter) ? e.parameter : {};
  var dados = lerDados(p);
  var id = String(Date.now());

  // 1) Grava na planilha (um erro aqui n\u00e3o impede o e-mail)
  try {
    gravarNaPlanilha(dados, id);
    resultado.etapas.planilha = 'ok';
  } catch (err) {
    resultado.ok = false;
    resultado.etapas.planilha = String(err);
  }

  // 2) E-mails com o relat\u00f3rio NO CORPO (HTML), sem anexo
  try {
    if (dados.email) {
      MailApp.sendEmail({
        to: dados.email,
        name: REMETENTE_NOME,
        subject: 'Seu relat\u00f3rio de Autodiagn\u00f3stico Regulat\u00f3rio BPF \u2014 Dal Cero',
        htmlBody: corpoEmailLead(dados)
      });
    }
    if (EMAILS_INTERNOS && EMAILS_INTERNOS.length) {
      MailApp.sendEmail({
        to: EMAILS_INTERNOS.join(','),
        name: REMETENTE_NOME,
        subject: 'Novo autodiagn\u00f3stico: ' + dados.empresa + ' (score ' + dados.score + ')',
        htmlBody: corpoEmailInterno(dados)
      });
    }
    resultado.etapas.email = 'ok';
  } catch (err2) {
    resultado.ok = false;
    resultado.etapas.email = String(err2);
  }

  return ContentService.createTextOutput(JSON.stringify(resultado))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput('Autodiagn\u00f3stico BPF \u2014 endpoint ativo (e-mail + base de dados).');
}

/* ---------- Leitura e normaliza\u00e7\u00e3o dos dados recebidos ---------- */
function lerDados(p) {
  var secoes = [], riscos = [], respostas = [];
  try { secoes = JSON.parse(p.secoes_json || '[]'); } catch (e) {}
  try { riscos = JSON.parse(p.riscos_json || '[]'); } catch (e) {}
  try { respostas = JSON.parse(p.respostas_json || '[]'); } catch (e) {}
  return {
    nome: p.nome || '',
    empresa: p.empresa || '(empresa n\u00e3o informada)',
    email: p.email || '',
    registro: p.registro || '',
    marketing: p.marketing || 'nao',
    dataISO: p.data || '',
    data: formatarData(p.data),
    score: p.score_geral || '0',
    classificacao: p.classificacao || '',
    perfil: p.perfil || '',
    secoes: secoes,
    riscos: riscos,
    respostas: respostas
  };
}

/* ---------- Grava\u00e7\u00e3o na Planilha Google (base de dados) ---------- */
function gravarNaPlanilha(d, id) {
  if (!SHEET_ID || SHEET_ID.indexOf('COLE_AQUI') === 0) return; // sem planilha configurada
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // Aba "Diagnosticos": 1 linha por preenchimento
  var aDiag = getOrCreateSheet(ss, 'Diagnosticos',
    ['ID', 'Data', 'Empresa', 'Respons\u00e1vel', 'E-mail', 'Registro MAPA',
     'Score', 'Classifica\u00e7\u00e3o', 'Qtd. riscos O/RR', 'Aceita marketing', 'Perfil']);
  aDiag.appendRow([
    id, d.data, d.empresa, d.nome, d.email, d.registro,
    Number(d.score), d.classificacao, d.riscos.length,
    (d.marketing === 'sim' ? 'Sim' : 'N\u00e3o'), d.perfil
  ]);

  // Aba "Respostas": 1 linha por pergunta respondida (base p/ \u00edndice de n\u00e3o conformidade)
  var aResp = getOrCreateSheet(ss, 'Respostas',
    ['ID', 'Data', 'Empresa', 'Item', 'Se\u00e7\u00e3o', 'Tipo', 'Conformidade', 'Pergunta']);
  if (d.respostas && d.respostas.length) {
    var linhas = d.respostas.map(function (r) {
      return [id, d.data, d.empresa, r.id, r.secao, r.tipo, r.conformidade, r.pergunta];
    });
    aResp.getRange(aResp.getLastRow() + 1, 1, linhas.length, linhas[0].length).setValues(linhas);
  }
}

function getOrCreateSheet(ss, nome, cabecalho) {
  var sh = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome);
    sh.appendRow(cabecalho);
    sh.getRange(1, 1, 1, cabecalho.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ---------- Cores ---------- */
function corClassificacao(nome) {
  if (/baixo/i.test(nome)) return COR.verde;
  if (/m\u00e9dio|medio/i.test(nome)) return COR.amarelo;
  if (/alt\u00edssimo|altissimo/i.test(nome)) return COR.vermelho;
  if (/alto/i.test(nome)) return COR.laranja;
  return COR.azul;
}
function corBarra(pct) {
  if (pct >= 90) return COR.verde;
  if (pct >= 70) return COR.amarelo;
  if (pct >= 50) return COR.laranja;
  return COR.vermelho;
}

/* ---------- Relat\u00f3rio (vai NO CORPO do e-mail) ---------- */
// Retorna um fragmento HTML (sem <html>/<body>) que os clientes de e-mail renderizam bem.
function montarRelatorioHtml(d) {
  var linhasSecoes = d.secoes.map(function (s) {
    var cor = corBarra(Number(s.pct));
    return '<tr>' +
      '<td style="padding:7px 8px;font-size:12px;color:#1f2937;border-bottom:1px solid #eee;">' + esc(s.nome) + '</td>' +
      '<td style="padding:7px 8px;border-bottom:1px solid #eee;width:160px;">' +
        '<div style="background:#eaeaea;border-radius:4px;height:9px;width:100%;">' +
          '<div style="background:' + cor + ';height:9px;border-radius:4px;width:' + Number(s.pct) + '%;"></div>' +
        '</div></td>' +
      '<td style="padding:7px 8px;font-size:13px;font-weight:bold;color:' + COR.navy + ';text-align:right;border-bottom:1px solid #eee;width:42px;">' + esc(s.pct) + '</td>' +
    '</tr>';
  }).join('');

  var blocoRiscos;
  if (!d.riscos.length) {
    blocoRiscos = '<p style="font-size:13px;color:#065f46;background:#ecfdf5;padding:12px;border-radius:8px;margin:0;">' +
      'Nenhum item obrigat\u00f3rio (O) ou de Risco Regulat\u00f3rio (RR) foi marcado como n\u00e3o conforme ou parcial. Excelente!</p>';
  } else {
    blocoRiscos = d.riscos.map(function (r) {
      var corL = r.status === 'Parcial' ? COR.amarelo : COR.vermelho;
      return '<div style="background:#fff;border-left:3px solid ' + corL + ';border-radius:6px;padding:9px 11px;margin-bottom:7px;">' +
        '<div style="font-size:10px;font-weight:bold;color:' + corL + ';margin-bottom:3px;">' +
          esc(r.status) + (r.tipo ? ' \u00b7 ' + esc(r.tipo) : '') + ' \u00b7 Item ' + esc(r.id) + ' \u00b7 ' + esc(r.secao) +
        '</div>' +
        '<div style="font-size:12px;color:#1f2937;line-height:1.4;">' + esc(r.pergunta) + '</div>' +
      '</div>';
    }).join('');
  }

  var cartao = '' +
    '<div style="border:1px solid #e5e7eb;border-radius:12px;overflow:hidden;">' +
      '<div style="background:' + COR.navy + ';color:#fff;padding:18px 24px;">' +
        '<div style="font-size:11px;letter-spacing:2px;color:' + COR.dourado + ';text-transform:uppercase;">Dal Cero Consultoria \u00b7 Regulat\u00f3rio MAPA</div>' +
        '<div style="font-size:18px;font-weight:bold;margin-top:4px;">Autodiagn\u00f3stico Regulat\u00f3rio BPF / Autocontroles</div>' +
      '</div>' +
      '<div style="padding:20px 24px;">' +
        '<table style="width:100%;font-size:12px;color:#374151;margin-bottom:16px;"><tr>' +
          '<td><b>Empresa:</b> ' + esc(d.empresa) + '<br><b>Respons\u00e1vel:</b> ' + esc(d.nome) + '</td>' +
          '<td style="text-align:right;"><b>Registro MAPA:</b> ' + esc(d.registro || '\u2014') + '<br><b>Data:</b> ' + esc(d.data) + '</td>' +
        '</tr></table>' +

        '<div style="background:' + COR.navy + ';border-radius:10px;padding:18px;text-align:center;color:#fff;margin-bottom:20px;">' +
          '<div style="font-size:11px;letter-spacing:1.5px;opacity:.85;text-transform:uppercase;">Score de Conformidade BPF/Autocontroles</div>' +
          '<div style="font-size:44px;font-weight:bold;line-height:1.1;">' + esc(d.score) + '</div>' +
          '<div style="display:inline-block;margin-top:6px;padding:5px 16px;border-radius:16px;background:' + corClassificacao(d.classificacao) + ';color:#fff;font-weight:bold;font-size:13px;">' + esc(d.classificacao) + '</div>' +
        '</div>' +

        '<h3 style="color:' + COR.navy + ';font-size:15px;margin:0 0 8px;">Desempenho por \u00e1rea</h3>' +
        '<table style="width:100%;border-collapse:collapse;margin-bottom:20px;">' + linhasSecoes + '</table>' +

        '<h3 style="color:' + COR.navy + ';font-size:15px;margin:0 0 8px;">Riscos regulat\u00f3rios identificados</h3>' +
        '<div style="background:#fff7ed;border-left:4px solid ' + COR.dourado + ';border-radius:8px;padding:14px;">' + blocoRiscos + '</div>' +
      '</div>' +
    '</div>';

  return '<div style="max-width:640px;margin:0 auto;font-family:Arial,Helvetica,sans-serif;">' +
    cartao +
    montarRodapeHtml() +
    '<p style="font-size:10px;color:' + COR.cinza + ';margin:16px 6px 0;line-height:1.5;text-align:center;">Esta \u00e9 uma autoavalia\u00e7\u00e3o indicativa, preenchida pelo pr\u00f3prio estabelecimento. O c\u00e1lculo oficial do Risco Regulat\u00f3rio \u00e9 realizado pelo servi\u00e7o de inspe\u00e7\u00e3o competente. Baseado no Termo de Fiscaliza\u00e7\u00e3o do MAPA (TF-BPF/Autocontroles e M\u00f3dulo II de Medicamentos).</p>' +
    '<p style="font-size:11px;color:#9ca3af;margin:10px 6px 0;line-height:1.5;text-align:center;">Voc\u00ea est\u00e1 recebendo este e-mail porque preencheu o Autodiagn\u00f3stico Regulat\u00f3rio BPF no site da Dal Cero Consultoria.</p>' +
  '</div>';
}

// Rodap\u00e9 no estilo da calculadora: painel azul com grade de contatos e barra de endere\u00e7o
function montarRodapeHtml() {
  return '' +
  '<div style="background-color:' + COR.navy + ';background:linear-gradient(135deg,' + COR.navy + ' 0%,#003B91 60%,' + COR.azul + ' 100%);color:#fff;border-radius:14px;padding:28px 24px;margin-top:20px;">' +
    '<span style="display:inline-block;background:rgba(255,255,255,0.12);border:1px solid rgba(232,180,97,0.45);color:' + COR.dourado + ';padding:6px 14px;border-radius:999px;font-size:11px;font-weight:700;letter-spacing:1.6px;text-transform:uppercase;margin-bottom:14px;">Dal Cero Consultoria</span>' +
    '<div style="font-size:22px;font-weight:800;line-height:1.25;margin-bottom:12px;">Identificou pontos de aten\u00e7\u00e3o? N\u00f3s temos a solu\u00e7\u00e3o.</div>' +
    '<p style="font-size:14px;color:rgba(255,255,255,0.88);line-height:1.7;margin:0 0 22px;">A <strong style="color:#fff;">Dal Cero Consultoria</strong> \u00e9 especialista em regulat\u00f3rio MAPA para o setor de alimenta\u00e7\u00e3o animal \u2014 diagn\u00f3stico, adequa\u00e7\u00e3o de Boas Pr\u00e1ticas de Fabrica\u00e7\u00e3o, registros no SipeAgro, suporte em fiscaliza\u00e7\u00f5es e gest\u00e3o regulat\u00f3ria cont\u00ednua. <strong style="color:#fff;">Fale com nosso comercial e construa um plano personalizado.</strong></p>' +
    '<table role="presentation" width="100%" style="border-collapse:collapse;margin-bottom:20px;">' +
      '<tr><td colspan="2" style="padding:6px;">' +
        cardContato('Time Comercial',
          '<a style="color:#fff;text-decoration:none;" href="https://wa.me/5549999710329">(49) 99971-0329</a>&nbsp;&nbsp;\u00b7&nbsp;&nbsp;' +
          '<a style="color:#fff;text-decoration:none;" href="https://wa.me/5549991993297">(49) 99199-3297</a>') +
      '</td></tr>' +
      '<tr>' +
        '<td width="50%" style="padding:6px;">' + cardContato('E-mail Comercial', '<a style="color:#fff;text-decoration:none;" href="mailto:comercial@dalceroconsultoria.com.br">comercial@dalceroconsultoria.com.br</a>') + '</td>' +
        '<td width="50%" style="padding:6px;">' + cardContato('Site', '<a style="color:#fff;text-decoration:none;" href="https://dalceroacademy.com">dalceroacademy.com</a>') + '</td>' +
      '</tr>' +
      '<tr>' +
        '<td style="padding:6px;">' + cardContato('LinkedIn', '<a style="color:#fff;text-decoration:none;" href="https://www.linkedin.com/company/grupo-dal-cero/">/grupo-dal-cero</a>') + '</td>' +
        '<td style="padding:6px;">' + cardContato('Instagram', '<a style="color:#fff;text-decoration:none;" href="https://www.instagram.com/grupodalcero/">@grupodalcero</a>') + '</td>' +
      '</tr>' +
    '</table>' +
    '<div style="border-top:1px solid rgba(255,255,255,0.15);padding-top:16px;text-align:center;font-size:12.5px;color:rgba(255,255,255,0.72);line-height:1.6;">' +
      '<strong style="color:#fff;display:block;margin-bottom:4px;letter-spacing:0.06em;">Dal Cero Consultoria</strong>' +
      'Rua Prefeito Albino Cerutti Cella, 322, Sala 06 \u00b7 Centro \u00b7 Maravilha \u2014 SC' +
    '</div>' +
  '</div>';
}
function cardContato(label, valor) {
  return '<div style="background:rgba(255,255,255,0.06);padding:12px 14px;border-radius:8px;border:1px solid rgba(255,255,255,0.10);">' +
    '<div style="font-size:10px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:' + COR.dourado + ';margin-bottom:4px;">' + label + '</div>' +
    '<div style="font-size:14px;color:#fff;word-break:break-word;">' + valor + '</div>' +
  '</div>';
}

/* ---------- Corpos dos e-mails (relat\u00f3rio no corpo) ---------- */
function corpoEmailLead(d) {
  var primeiro = (d.nome || '').split(' ')[0] || 'Ol\u00e1';
  return '<div style="font-family:Arial,sans-serif;color:#1f2937;font-size:14px;line-height:1.6;">' +
    '<p style="max-width:640px;margin:0 auto 14px;">' + esc(primeiro) + ', tudo bem? Segue abaixo o resultado do seu <b>Autodiagn\u00f3stico Regulat\u00f3rio BPF/Autocontroles</b>.</p>' +
    montarRelatorioHtml(d) +
  '</div>';
}
function corpoEmailInterno(d) {
  return '<div style="font-family:Arial,sans-serif;color:#1f2937;font-size:14px;line-height:1.6;">' +
    '<div style="max-width:640px;margin:0 auto;">' +
      '<p><b>Novo autodiagn\u00f3stico preenchido.</b></p>' +
      '<table style="font-size:13px;margin-bottom:16px;">' +
        linha('Empresa', d.empresa) + linha('Respons\u00e1vel', d.nome) + linha('E-mail', d.email) +
        linha('Registro MAPA', d.registro || '\u2014') + linha('Data', d.data) +
        linha('Score', d.score + ' (' + d.classificacao + ')') +
        linha('Riscos O/RR', String(d.riscos.length)) + linha('Perfil', d.perfil) +
      '</table>' +
    '</div>' +
    montarRelatorioHtml(d) +
  '</div>';
}
function linha(rotulo, valor) {
  return '<tr><td style="padding:3px 10px 3px 0;color:#6b7280;vertical-align:top;"><b>' + esc(rotulo) + ':</b></td>' +
         '<td style="padding:3px 0;">' + esc(valor) + '</td></tr>';
}

/* ---------- Utilidades ---------- */
function esc(v) {
  return String(v == null ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function formatarData(iso) {
  try {
    var dt = iso ? new Date(iso) : new Date();
    return Utilities.formatDate(dt, Session.getScriptTimeZone() || 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
  } catch (e) { return ''; }
}

/* ---------- Teste manual (rode uma vez para autorizar e validar) ---------- */
function testarEnvio() {
  doPost({ parameter: {
    nome: 'Teste Dal Cero', empresa: 'F\u00e1brica Exemplo', email: Session.getActiveUser().getEmail(),
    registro: 'SC-00/0000', marketing: 'sim', data: new Date().toISOString(),
    score_geral: '72', classificacao: 'Risco M\u00e9dio',
    perfil: 'Aceita marketing: Sim | Fabrica medicamentos: N\u00e3o',
    secoes_json: JSON.stringify([{ nome: 'Documenta\u00e7\u00e3o e Registro', pct: 80 }, { nome: '\u00c1rea Externa', pct: 50 }]),
    riscos_json: JSON.stringify([{ id: '21', secao: 'Qualifica\u00e7\u00e3o de Fornecedores', status: 'N\u00e3o conforme', tipo: 'O+RR', pergunta: 'Lote vencido esquecido no estoque?' }]),
    respostas_json: JSON.stringify([
      { id: '1', secao: 'Documenta\u00e7\u00e3o e Registro', tipo: '\u2014', conformidade: 'Conforme', pergunta: 'Pergunta 1' },
      { id: '21', secao: 'Qualifica\u00e7\u00e3o de Fornecedores', tipo: 'O+RR', conformidade: 'N\u00e3o conforme', pergunta: 'Lote vencido esquecido no estoque?' }
    ])
  }});
}
