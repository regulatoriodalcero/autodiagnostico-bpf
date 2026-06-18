/**
 * ============================================================================
 *  AUTODIAGN\u00d3STICO REGULAT\u00d3RIO BPF \u2014 Dal Cero Consultoria
 *  Backend de e-mail (Google Apps Script Web App)
 *
 *  O app (p\u00e1gina) envia os dados via POST (FormData). Este script monta um
 *  relat\u00f3rio em PDF e envia por e-mail para a pessoa e para o time comercial.
 *
 *  COMO PUBLICAR (uma vez):
 *   1. Acesse https://script.google.com  \u2192 Novo projeto.
 *   2. Apague o conte\u00fado padr\u00e3o e cole TODO este arquivo.
 *   3. (Opcional) Ajuste EMAILS_INTERNOS abaixo.
 *   4. Implantar \u2192 Nova implanta\u00e7\u00e3o \u2192 Tipo: "App da Web".
 *        - Executar como: Eu (sua conta Google)
 *        - Quem pode acessar: Qualquer pessoa
 *   5. Autorize as permiss\u00f5es quando pedir.
 *   6. Copie a URL que termina em /exec e cole na constante
 *      APPS_SCRIPT_URL do index.html do autodiagn\u00f3stico.
 *
 *  Para testar sem o site: rode a fun\u00e7\u00e3o testarEnvio() uma vez.
 * ============================================================================
 */

// >>> Quem recebe a C\u00d3PIA interna de cada autodiagn\u00f3stico (al\u00e9m da pr\u00f3pria pessoa)
var EMAILS_INTERNOS = ['comercial@dalceroconsultoria.com.br'];
// Para incluir o regulat\u00f3rio, use:
// var EMAILS_INTERNOS = ['comercial@dalceroconsultoria.com.br', 'regulatorio@dalceroconsultoria.com.br'];

var REMETENTE_NOME = 'Dal Cero Consultoria';

// Paleta da marca
var COR = {
  navy: '#041F47', azul: '#0043D8', azulClaro: '#177fe5', dourado: '#E8B461',
  cinza: '#6b7280', verde: '#10b981', amarelo: '#f59e0b', laranja: '#f97316', vermelho: '#ef4444'
};

function doPost(e) {
  try {
    var p = (e && e.parameter) ? e.parameter : {};
    var dados = lerDados(p);
    var html = montarHtmlRelatorio(dados);
    var pdf = Utilities.newBlob(html, 'text/html', 'relatorio.html')
      .getAs('application/pdf')
      .setName('Autodiagnostico-BPF-' + sanitizar(dados.empresa) + '.pdf');

    // 1) E-mail para a pessoa
    if (dados.email) {
      MailApp.sendEmail({
        to: dados.email,
        name: REMETENTE_NOME,
        subject: 'Seu relat\u00f3rio de Autodiagn\u00f3stico Regulat\u00f3rio BPF \u2014 Dal Cero',
        htmlBody: corpoEmailLead(dados),
        attachments: [pdf]
      });
    }

    // 2) C\u00f3pia interna (notifica\u00e7\u00e3o de lead) para o time comercial
    if (EMAILS_INTERNOS && EMAILS_INTERNOS.length) {
      MailApp.sendEmail({
        to: EMAILS_INTERNOS.join(','),
        name: REMETENTE_NOME,
        subject: 'Novo autodiagn\u00f3stico: ' + dados.empresa + ' (score ' + dados.score + ')',
        htmlBody: corpoEmailInterno(dados),
        attachments: [pdf]
      });
    }

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, erro: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Permite abrir a URL no navegador s\u00f3 para confirmar que est\u00e1 publicada
function doGet() {
  return ContentService.createTextOutput('Autodiagn\u00f3stico BPF \u2014 endpoint de e-mail ativo.');
}

/* ---------- Leitura e normaliza\u00e7\u00e3o dos dados recebidos ---------- */
function lerDados(p) {
  var secoes = [], riscos = [];
  try { secoes = JSON.parse(p.secoes_json || '[]'); } catch (e) {}
  try { riscos = JSON.parse(p.riscos_json || '[]'); } catch (e) {}
  return {
    nome: p.nome || '',
    empresa: p.empresa || '(empresa n\u00e3o informada)',
    email: p.email || '',
    registro: p.registro || '',
    data: formatarData(p.data),
    score: p.score_geral || '0',
    classificacao: p.classificacao || '',
    perfil: p.perfil || '',
    secoes: secoes,
    riscos: riscos
  };
}

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

/* ---------- HTML do relat\u00f3rio (vira PDF) ---------- */
function montarHtmlRelatorio(d) {
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
    blocoRiscos = '<p style="font-size:13px;color:#065f46;background:#ecfdf5;padding:12px;border-radius:8px;">' +
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

  return '' +
  '<html><body style="font-family:Arial,Helvetica,sans-serif;margin:0;padding:0;color:#1f2937;">' +
    '<div style="background:' + COR.navy + ';color:#fff;padding:22px 28px;">' +
      '<div style="font-size:11px;letter-spacing:2px;color:' + COR.dourado + ';text-transform:uppercase;">Dal Cero Consultoria \u00b7 Regulat\u00f3rio MAPA</div>' +
      '<div style="font-size:20px;font-weight:bold;margin-top:4px;">Autodiagn\u00f3stico Regulat\u00f3rio BPF / Autocontroles</div>' +
    '</div>' +
    '<div style="padding:24px 28px;">' +
      '<table style="width:100%;font-size:12px;color:#374151;margin-bottom:18px;"><tr>' +
        '<td><b>Empresa:</b> ' + esc(d.empresa) + '<br><b>Respons\u00e1vel:</b> ' + esc(d.nome) + '</td>' +
        '<td style="text-align:right;"><b>Registro MAPA:</b> ' + esc(d.registro || '\u2014') + '<br><b>Data:</b> ' + esc(d.data) + '</td>' +
      '</tr></table>' +

      '<div style="background:linear-gradient(135deg,' + COR.navy + ',' + COR.azul + ');border-radius:10px;padding:20px;text-align:center;color:#fff;margin-bottom:22px;">' +
        '<div style="font-size:11px;letter-spacing:1.5px;opacity:.85;text-transform:uppercase;">Score de Conformidade BPF/Autocontroles</div>' +
        '<div style="font-size:46px;font-weight:bold;line-height:1.1;">' + esc(d.score) + '</div>' +
        '<div style="display:inline-block;margin-top:6px;padding:5px 16px;border-radius:16px;background:' + corClassificacao(d.classificacao) + ';font-weight:bold;font-size:13px;">' + esc(d.classificacao) + '</div>' +
      '</div>' +

      '<h3 style="color:' + COR.navy + ';font-size:15px;margin:0 0 8px;">Desempenho por \u00e1rea</h3>' +
      '<table style="width:100%;border-collapse:collapse;margin-bottom:22px;">' + linhasSecoes + '</table>' +

      '<h3 style="color:' + COR.navy + ';font-size:15px;margin:0 0 8px;">Riscos regulat\u00f3rios identificados</h3>' +
      '<div style="background:#fff7ed;border-left:4px solid ' + COR.dourado + ';border-radius:8px;padding:14px;">' + blocoRiscos + '</div>' +

      '<div style="margin-top:24px;background:' + COR.navy + ';color:#fff;border-radius:10px;padding:18px 20px;">' +
        '<div style="font-size:14px;font-weight:bold;margin-bottom:4px;">Quer transformar este diagn\u00f3stico em um plano de a\u00e7\u00e3o?</div>' +
        '<div style="font-size:12px;color:#dbe4f5;line-height:1.5;">A Dal Cero Consultoria \u00e9 especialista em regulat\u00f3rio MAPA para alimenta\u00e7\u00e3o animal \u2014 adequa\u00e7\u00e3o de BPF, registros no SipeAgro, suporte em fiscaliza\u00e7\u00f5es e gest\u00e3o cont\u00ednua.<br>' +
        'WhatsApp: (49) 99971-0329 \u00b7 (49) 99199-3297 \u00b7 comercial@dalceroconsultoria.com.br \u00b7 dalceroacademy.com</div>' +
      '</div>' +

      '<p style="font-size:10px;color:' + COR.cinza + ';margin-top:18px;line-height:1.5;">Esta \u00e9 uma autoavalia\u00e7\u00e3o indicativa, preenchida pelo pr\u00f3prio estabelecimento. O c\u00e1lculo oficial do Risco Regulat\u00f3rio \u00e9 realizado pelo servi\u00e7o de inspe\u00e7\u00e3o competente. Baseado no Termo de Fiscaliza\u00e7\u00e3o do MAPA (TF-BPF/Autocontroles e M\u00f3dulo II de Medicamentos).</p>' +
    '</div>' +
  '</body></html>';
}

/* ---------- Corpos dos e-mails ---------- */
function corpoEmailLead(d) {
  var primeiro = (d.nome || '').split(' ')[0] || 'Ol\u00e1';
  return '<div style="font-family:Arial,sans-serif;color:#1f2937;font-size:14px;line-height:1.6;">' +
    '<p>' + esc(primeiro) + ', tudo bem?</p>' +
    '<p>Segue em anexo o <b>relat\u00f3rio do seu Autodiagn\u00f3stico Regulat\u00f3rio BPF/Autocontroles</b> referente a <b>' + esc(d.empresa) + '</b>.</p>' +
    '<p>Seu score de conformidade foi <b>' + esc(d.score) + '</b> \u2014 classifica\u00e7\u00e3o <b>' + esc(d.classificacao) + '</b>. No PDF voc\u00ea encontra o desempenho por \u00e1rea e os pontos de aten\u00e7\u00e3o identificados.</p>' +
    '<p>Quer ajuda para tratar esses pontos? Fale com a gente:<br>WhatsApp (49) 99971-0329 \u00b7 (49) 99199-3297<br>comercial@dalceroconsultoria.com.br</p>' +
    '<p style="color:#6b7280;font-size:12px;">Dal Cero Consultoria \u00b7 Regulat\u00f3rio MAPA \u00b7 Alimenta\u00e7\u00e3o Animal</p>' +
  '</div>';
}
function corpoEmailInterno(d) {
  return '<div style="font-family:Arial,sans-serif;color:#1f2937;font-size:14px;line-height:1.6;">' +
    '<p><b>Novo autodiagn\u00f3stico preenchido.</b></p>' +
    '<table style="font-size:13px;">' +
      linha('Empresa', d.empresa) + linha('Respons\u00e1vel', d.nome) + linha('E-mail', d.email) +
      linha('Registro MAPA', d.registro || '\u2014') + linha('Data', d.data) +
      linha('Score', d.score + ' (' + d.classificacao + ')') +
      linha('Riscos O/RR', String(d.riscos.length)) + linha('Perfil', d.perfil) +
    '</table>' +
    '<p style="color:#6b7280;font-size:12px;">Relat\u00f3rio completo em anexo (PDF).</p>' +
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
function sanitizar(v) { return String(v || '').replace(/[^\w\-]+/g, '_').slice(0, 40); }
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
    registro: 'SC-00/0000', data: new Date().toISOString(),
    score_geral: '72', classificacao: 'Risco M\u00e9dio',
    perfil: 'Fabrica medicamentos: N\u00e3o | Controle de pragas: Terceirizado',
    secoes_json: JSON.stringify([{ nome: 'Documenta\u00e7\u00e3o e Registro', pct: 80 }, { nome: '\u00c1rea Externa', pct: 50 }, { nome: 'Treinamentos', pct: 100 }]),
    riscos_json: JSON.stringify([{ id: '21', secao: 'Qualifica\u00e7\u00e3o de Fornecedores', status: 'N\u00e3o conforme', tipo: 'O+RR', pergunta: 'Lote vencido esquecido no estoque?' }])
  }});
}
