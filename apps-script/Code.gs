/**
 * ============================================================================
 *  AUTODIAGNÓSTICO REGULATÓRIO BPF — Dal Cero Consultoria
 *  Backend de e-mail (Google Apps Script Web App)
 *
 *  O app (página) envia os dados via POST (FormData). Este script monta um
 *  relatório em PDF e envia por e-mail para a pessoa e para o time comercial.
 *
 *  COMO PUBLICAR (uma vez):
 *   1. Acesse https://script.google.com  → Novo projeto.
 *   2. Apague o conteúdo padrão e cole TODO este arquivo.
 *   3. (Opcional) Ajuste EMAILS_INTERNOS abaixo.
 *   4. Implantar → Nova implantação → Tipo: "App da Web".
 *        - Executar como: Eu (sua conta Google)
 *        - Quem pode acessar: Qualquer pessoa
 *   5. Autorize as permissões quando pedir.
 *   6. Copie a URL que termina em /exec e cole na constante
 *      APPS_SCRIPT_URL do index.html do autodiagnóstico.
 *
 *  Para testar sem o site: rode a função testarEnvio() uma vez.
 * ============================================================================
 */

// >>> Quem recebe a CÓPIA interna de cada autodiagnóstico (além da própria pessoa)
var EMAILS_INTERNOS = ['comercial@dalceroconsultoria.com.br'];
// Para incluir o regulatório, use:
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
        subject: 'Seu relatório de Autodiagnóstico Regulatório BPF — Dal Cero',
        htmlBody: corpoEmailLead(dados),
        attachments: [pdf]
      });
    }

    // 2) Cópia interna (notificação de lead) para o time comercial
    if (EMAILS_INTERNOS && EMAILS_INTERNOS.length) {
      MailApp.sendEmail({
        to: EMAILS_INTERNOS.join(','),
        name: REMETENTE_NOME,
        subject: 'Novo autodiagnóstico: ' + dados.empresa + ' (score ' + dados.score + ')',
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

// Permite abrir a URL no navegador só para confirmar que está publicada
function doGet() {
  return ContentService.createTextOutput('Autodiagnóstico BPF — endpoint de e-mail ativo.');
}

/* ---------- Leitura e normalização dos dados recebidos ---------- */
function lerDados(p) {
  var secoes = [], riscos = [];
  try { secoes = JSON.parse(p.secoes_json || '[]'); } catch (e) {}
  try { riscos = JSON.parse(p.riscos_json || '[]'); } catch (e) {}
  return {
    nome: p.nome || '',
    empresa: p.empresa || '(empresa não informada)',
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
  if (/médio|medio/i.test(nome)) return COR.amarelo;
  if (/altíssimo|altissimo/i.test(nome)) return COR.vermelho;
  if (/alto/i.test(nome)) return COR.laranja;
  return COR.azul;
}
function corBarra(pct) {
  if (pct >= 90) return COR.verde;
  if (pct >= 70) return COR.amarelo;
  if (pct >= 50) return COR.laranja;
  return COR.vermelho;
}

/* ---------- HTML do relatório (vira PDF) ---------- */
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
      'Nenhum item obrigatório (O) ou de Risco Regulatório (RR) foi marcado como não conforme ou parcial. Excelente!</p>';
  } else {
    blocoRiscos = d.riscos.map(function (r) {
      var corL = r.status === 'Parcial' ? COR.amarelo : COR.vermelho;
      return '<div style="background:#fff;border-left:3px solid ' + corL + ';border-radius:6px;padding:9px 11px;margin-bottom:7px;">' +
        '<div style="font-size:10px;font-weight:bold;color:' + corL + ';margin-bottom:3px;">' +
          esc(r.status) + (r.tipo ? ' · ' + esc(r.tipo) : '') + ' · Item ' + esc(r.id) + ' · ' + esc(r.secao) +
        '</div>' +
        '<div style="font-size:12px;color:#1f2937;line-height:1.4;">' + esc(r.pergunta) + '</div>' +
      '</div>';
    }).join('');
  }

  return '' +
  '<html><body style="font-family:Arial,Helvetica,sans-serif;margin:0;padding:0;color:#1f2937;">' +
    '<div style="background:' + COR.navy + ';color:#fff;padding:22px 28px;">' +
      '<div style="font-size:11px;letter-spacing:2px;color:' + COR.dourado + ';text-transform:uppercase;">Dal Cero Consultoria · Regulatório MAPA</div>' +
      '<div style="font-size:20px;font-weight:bold;margin-top:4px;">Autodiagnóstico Regulatório BPF / Autocontroles</div>' +
    '</div>' +
    '<div style="padding:24px 28px;">' +
      '<table style="width:100%;font-size:12px;color:#374151;margin-bottom:18px;"><tr>' +
        '<td><b>Empresa:</b> ' + esc(d.empresa) + '<br><b>Responsável:</b> ' + esc(d.nome) + '</td>' +
        '<td style="text-align:right;"><b>Registro MAPA:</b> ' + esc(d.registro || '—') + '<br><b>Data:</b> ' + esc(d.data) + '</td>' +
      '</tr></table>' +

      '<div style="background:linear-gradient(135deg,' + COR.navy + ',' + COR.azul + ');border-radius:10px;padding:20px;text-align:center;color:#fff;margin-bottom:22px;">' +
        '<div style="font-size:11px;letter-spacing:1.5px;opacity:.85;text-transform:uppercase;">Score de Conformidade BPF/Autocontroles</div>' +
        '<div style="font-size:46px;font-weight:bold;line-height:1.1;">' + esc(d.score) + '</div>' +
        '<div style="display:inline-block;margin-top:6px;padding:5px 16px;border-radius:16px;background:' + corClassificacao(d.classificacao) + ';font-weight:bold;font-size:13px;">' + esc(d.classificacao) + '</div>' +
      '</div>' +

      '<h3 style="color:' + COR.navy + ';font-size:15px;margin:0 0 8px;">Desempenho por área</h3>' +
      '<table style="width:100%;border-collapse:collapse;margin-bottom:22px;">' + linhasSecoes + '</table>' +

      '<h3 style="color:' + COR.navy + ';font-size:15px;margin:0 0 8px;">Riscos regulatórios identificados</h3>' +
      '<div style="background:#fff7ed;border-left:4px solid ' + COR.dourado + ';border-radius:8px;padding:14px;">' + blocoRiscos + '</div>' +

      '<div style="margin-top:24px;background:' + COR.navy + ';color:#fff;border-radius:10px;padding:18px 20px;">' +
        '<div style="font-size:14px;font-weight:bold;margin-bottom:4px;">Quer transformar este diagnóstico em um plano de ação?</div>' +
        '<div style="font-size:12px;color:#dbe4f5;line-height:1.5;">A Dal Cero Consultoria é especialista em regulatório MAPA para alimentação animal — adequação de BPF, registros no SipeAgro, suporte em fiscalizações e gestão contínua.<br>' +
        'WhatsApp: (49) 99971-0329 · (49) 99199-3297 · comercial@dalceroconsultoria.com.br · dalceroacademy.com</div>' +
      '</div>' +

      '<p style="font-size:10px;color:' + COR.cinza + ';margin-top:18px;line-height:1.5;">Esta é uma autoavaliação indicativa, preenchida pelo próprio estabelecimento. O cálculo oficial do Risco Regulatório é realizado pelo serviço de inspeção competente. Baseado no Termo de Fiscalização do MAPA (TF-BPF/Autocontroles e Módulo II de Medicamentos).</p>' +
    '</div>' +
  '</body></html>';
}

/* ---------- Corpos dos e-mails ---------- */
function corpoEmailLead(d) {
  var primeiro = (d.nome || '').split(' ')[0] || 'Olá';
  return '<div style="font-family:Arial,sans-serif;color:#1f2937;font-size:14px;line-height:1.6;">' +
    '<p>' + esc(primeiro) + ', tudo bem?</p>' +
    '<p>Segue em anexo o <b>relatório do seu Autodiagnóstico Regulatório BPF/Autocontroles</b> referente a <b>' + esc(d.empresa) + '</b>.</p>' +
    '<p>Seu score de conformidade foi <b>' + esc(d.score) + '</b> — classificação <b>' + esc(d.classificacao) + '</b>. No PDF você encontra o desempenho por área e os pontos de atenção identificados.</p>' +
    '<p>Quer ajuda para tratar esses pontos? Fale com a gente:<br>WhatsApp (49) 99971-0329 · (49) 99199-3297<br>comercial@dalceroconsultoria.com.br</p>' +
    '<p style="color:#6b7280;font-size:12px;">Dal Cero Consultoria · Regulatório MAPA · Alimentação Animal</p>' +
  '</div>';
}
function corpoEmailInterno(d) {
  return '<div style="font-family:Arial,sans-serif;color:#1f2937;font-size:14px;line-height:1.6;">' +
    '<p><b>Novo autodiagnóstico preenchido.</b></p>' +
    '<table style="font-size:13px;">' +
      linha('Empresa', d.empresa) + linha('Responsável', d.nome) + linha('E-mail', d.email) +
      linha('Registro MAPA', d.registro || '—') + linha('Data', d.data) +
      linha('Score', d.score + ' (' + d.classificacao + ')') +
      linha('Riscos O/RR', String(d.riscos.length)) + linha('Perfil', d.perfil) +
    '</table>' +
    '<p style="color:#6b7280;font-size:12px;">Relatório completo em anexo (PDF).</p>' +
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
    nome: 'Teste Dal Cero', empresa: 'Fábrica Exemplo', email: Session.getActiveUser().getEmail(),
    registro: 'SC-00/0000', data: new Date().toISOString(),
    score_geral: '72', classificacao: 'Risco Médio',
    perfil: 'Fabrica medicamentos: Não | Controle de pragas: Terceirizado',
    secoes_json: JSON.stringify([{ nome: 'Documentação e Registro', pct: 80 }, { nome: 'Área Externa', pct: 50 }, { nome: 'Treinamentos', pct: 100 }]),
    riscos_json: JSON.stringify([{ id: '21', secao: 'Qualificação de Fornecedores', status: 'Não conforme', tipo: 'O+RR', pergunta: 'Lote vencido esquecido no estoque?' }])
  }});
}
