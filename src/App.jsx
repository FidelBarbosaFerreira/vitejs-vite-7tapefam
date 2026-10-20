import { useState, useEffect } from 'react';
import { supabase } from './supabase';
import * as XLSX from 'xlsx';
import {
  MapPin,
  BarChart3,
  CalendarDays,
  Download,
  ShieldAlert,
  Check,
  Plus,
  Trash2,
  RefreshCw,
  Clock,
  Sun,
  Moon,
  ExternalLink,
  Upload,
  Truck,
  Search,
  History,
  DollarSign,
  AlertTriangle,
} from 'lucide-react';

const LINK_POWER_BI = 'https://app.powerbi.com';

// ============================================================================
// --- FUNÇÃO AUXILIAR DE TRATAMENTO DE DATAS INTERNAS DO EXCEL ---
// ============================================================================
function formatarDataVersaoBancoNoTopo(raw, dataPadrao) {
  if (!raw || String(raw).trim() === '' || String(raw).trim() === '-') return null;
  if (typeof raw === 'number') {
    const UTC_DAYS = Math.floor(raw - 25569);
    const UTC_VALUE = UTC_DAYS * 86400;
    return new Date(UTC_VALUE * 1000).toISOString().split('T')[0];
  }
  const textoData = String(raw).trim();
  if (textoData.includes('/')) {
    const partes = textoData.split('/');
    if (partes.length === 3) {
      const dia = partes[0].padStart(2, '0');
      const mes = partes[1].padStart(2, '0');
      const ano = partes[2].length === 2 ? `20${partes[2]}` : partes[2];
      return `${ano}-${mes}-${dia}`;
    }
  }
  return textoData;
}

// ============================================================================
// --- CONVERSOR AUXILIAR DE TEXTO PARA NÚMERO DE MÊS LOGÍSTICO ---
// ============================================================================
function extrairNumeroMesPorExtenso(texto) {
  if (!texto) return null;
  const t = String(texto).toUpperCase().trim();
  if (t.includes('JAN')) return 1;
  if (t.includes('FEV')) return 2;
  if (t.includes('MAR')) return 3;
  if (t.includes('ABR')) return 4;
  if (t.includes('MAI')) return 5;
  if (t.includes('JUN')) return 6;
  if (t.includes('JUL')) return 7;
  if (t.includes('AGO')) return 8;
  if (t.includes('SET')) return 9;
  if (t.includes('OUT')) return 10;
  if (t.includes('NOV')) return 11;
  if (t.includes('DEZ')) return 12;
  return null;
}

// ============================================================================
// --- FORMATADOR DE MOEDA BRL (R$) ---
// ============================================================================
function formatarMoeda(valor) {
  const num = Number(valor) || 0;
  return num.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

// ============================================================================
// --- COMPONENTE PRINCIPAL ---
// ============================================================================
function App() {
  const mesesDoAno = [
    { num: 1, nome: 'Janeiro', sigla: 'jan' },
    { num: 2, nome: 'Fevereiro', sigla: 'fev' },
    { num: 3, nome: 'Março', sigla: 'mar' },
    { num: 4, nome: 'Abril', sigla: 'abr' },
    { num: 5, nome: 'Maio', sigla: 'mai' },
    { num: 6, nome: 'Junho', sigla: 'jun' },
    { num: 7, nome: 'Julho', sigla: 'jul' },
    { num: 8, nome: 'Agosto', sigla: 'ago' },
    { num: 9, nome: 'Setembro', sigla: 'set' },
    { num: 10, nome: 'Outubro', sigla: 'out' },
    { num: 11, nome: 'Novembro', sigla: 'nov' },
    { num: 12, nome: 'Dezembro', sigla: 'dez' },
  ];

  // --- ESTADOS DE AUTENTICAÇÃO ---
  const [autenticado, setAutenticado] = useState(false);
  const [usuarioInput, setUsuarioInput] = useState('');
  const [senhaInput, setSenhaInput] = useState('');
  const [erroLogin, setErroLogin] = useState('');

  // --- ESTADOS GERAIS ---
  const [activeTab, setActiveTab] = useState('frota');
  const [filialAtual, setFilialAtual] = useState('Todas');
  const [subAbaFrota, setSubAbaFrota] = useState('preventivas');
  const [carregando, setCarregando] = useState(false);
  const [rotacionar, setRotacionar] = useState(false);
  const [isDarkMode, setIsDarkMode] = useState(true);

  // --- FILTROS DE BUSCA ---
  const [termoBusca, setTermoBusca] = useState('');
  const [termoBuscaTaco, setTermoBuscaTaco] = useState('');
  const [termoBuscaGerencial, setTermoBuscaGerencial] = useState('');
  const [termoBuscaGerencialTaco, setTermoBuscaGerencialTaco] = useState('');
  const [termoBuscaFinanceiro, setTermoBuscaFinanceiro] = useState('');

  // --- SELETORES DE DATA ---
  const [mesSelecionadoHistorico, setMesSelecionadoHistorico] = useState(
    new Date().getMonth() + 1
  );
  const [mesSelecionadoTacoGerencial, setMesSelecionadoTacoGerencial] =
    useState(new Date().getMonth() + 1);
  const [mesSelecionadoFinanceiro, setMesSelecionadoFinanceiro] = useState(10);

  // --- INTERAÇÃO DE FILTRO DOS CARDS ---
  const [statusFiltroCard, setStatusFiltroCard] = useState('TODOS');
  const [statusFiltroCardTaco, setStatusFiltroCardTaco] = useState('TODOS');
  const [contaFiltroFinanceiro, setContaFiltroFinanceiro] = useState('TODAS');
  const [nfFiltroFinanceiro, setNfFiltroFinanceiro] = useState('TODAS');
  const [cdvFiltroFinanceiro, setCdvFiltroFinanceiro] = useState('TODOS');

  // --- FROTA OFICIAL POR CDV (ABA POSIÇÃO NACIONAL - SUL2 = 485 PLACAS) ---
  const [frotaPorCdv, setFrotaPorCdv] = useState({
    'CDV BELTRAO': 24,
    'CDV CAMPO MOURAO': 39,
    'CDV CASCAVEL': 60,
    'CDV GASPAR': 101,
    'CDV LONDRINA': 69,
    'CDV MARINGA': 2,
    'CDV MEGA CURITIBA': 135,
    'CDV PONTA GROSSA': 55,
  });

  // --- INTERRUPTORES DE SUBSTITUIÇÃO ---
  const [substituirFrota, setSubstituirFrota] = useState(false);
  const [substituirTaco, setSubstituirTaco] = useState(false);
  const [substituirFinanceiro, setSubstituirFinanceiro] = useState(true);

  // --- DADOS REAIS DO BANCO ---
  const [veiculos, setVeiculos] = useState([]);
  const [preventivas, setPreventivas] = useState([]);
  const [tacografos, setTacografos] = useState([]);
  const [finRateio, setFinRateio] = useState([]);
  const [finBase, setFinBase] = useState([]);
  const [finCta, setFinCta] = useState([]);

  useEffect(() => {
    if (autenticado) {
      fetchDados();
    }
  }, [autenticado]);

  async function fetchDados() {
    setCarregando(true);
    setRotacionar(true);
    try {
      const { data: vecData } = await supabase
        .from('veiculos')
        .select('*')
        .order('placa', { ascending: true });
      setVeiculos(vecData || []);

      const { data: prevData } = await supabase
        .from('preventivas')
        .select('*')
        .order('placa', { ascending: true });
      setPreventivas(prevData || []);

      const { data: tacoData } = await supabase
        .from('tacografos')
        .select('*')
        .order('placa', { ascending: true });
      setTacografos(tacoData || []);

      const { data: rateioData } = await supabase
        .from('financeiro_rateio')
        .select('*')
        .order('id', { ascending: true });
      setFinRateio(rateioData || []);
      if (rateioData && rateioData.length > 0 && rateioData[0].mes) {
        setMesSelecionadoFinanceiro(Number(rateioData[0].mes));
      }

      const { data: ctaData } = await supabase
        .from('financeiro_cta')
        .select('*')
        .order('valor', { ascending: false });
      setFinCta(ctaData || []);

      // Busca até 10000 registros da base financeira em lotes
      let todasBase = [];
      let de = 0;
      const passo = 1000;
      while (true) {
        const { data: lote, error: errBase } = await supabase
          .from('financeiro_base')
          .select('*')
          .range(de, de + passo - 1)
          .order('id', { ascending: false });
        if (errBase || !lote || lote.length === 0) break;
        todasBase = todasBase.concat(lote);
        if (lote.length < passo) break;
        de += passo;
      }
      setFinBase(todasBase);
    } catch (error) {
      console.error('Erro na sincronização:', error.message);
    } finally {
      setCarregando(false);
      setTimeout(() => setRotacionar(false), 600);
    }
  }

  const lidarComLogin = (e) => {
    e.preventDefault();
    if (usuarioInput === 'SUL2CO' && senhaInput === 'frota2026') {
      setAutenticado(true);
      setErroLogin('');
    } else {
      setErroLogin('Usuário ou senha inválidos operacionalmente!');
    }
  };

  // --- BOTÕES MASTER DE EXCLUSÃO ---
  const limparTodaTabelaPreventivas = async () => {
    if (
      confirm(
        '⚠️ CUIDADO: Você vai apagar TODOS os registros da tabela de preventivas no Supabase. Tem certeza absoluta que quer fazer isso?'
      )
    ) {
      setCarregando(true);
      try {
        const { error } = await supabase
          .from('preventivas')
          .delete()
          .neq('placa', 'REMOVER_TUDO_MASTER');
        if (error) throw error;
        alert('Tabela de preventivas completamente limpa!');
        fetchDados();
      } catch (err) {
        alert('Erro ao limpar banco: ' + err.message);
      } finally {
        setCarregando(false);
      }
    }
  };

  const limparTodaTabelaTacografos = async () => {
    if (
      confirm(
        '⚠️ CUIDADO: Você vai apagar TODOS os registros da tabela de tacógrafos no Supabase. Tem certeza absoluta que quer fazer isso?'
      )
    ) {
      setCarregando(true);
      try {
        const { error } = await supabase
          .from('tacografos')
          .delete()
          .neq('placa', 'REMOVER_TUDO_MASTER');
        if (error) throw error;
        alert('Tabela de tacógrafos completamente limpa!');
        fetchDados();
      } catch (err) {
        alert('Erro ao limpar banco: ' + err.message);
      } finally {
        setCarregando(false);
      }
    }
  };

  const limparTodaTabelaFinanceiro = async () => {
    if (
      confirm(
        '⚠️ CUIDADO: Você vai apagar TODOS os registros Financeiros (Rateio, Base de OS e Alertas CTA) no Supabase. Tem certeza absoluta?'
      )
    ) {
      setCarregando(true);
      try {
        await supabase.from('financeiro_rateio').delete().neq('conta', 'REMOVER_TUDO');
        await supabase.from('financeiro_cta').delete().neq('placa', 'REMOVER_TUDO');
        await supabase.from('financeiro_base').delete().neq('os', 'REMOVER_TUDO');
        alert('Tabelas financeiras completamente limpas!');
        fetchDados();
      } catch (err) {
        alert('Erro ao limpar banco financeiro: ' + err.message);
      } finally {
        setCarregando(false);
      }
    }
  };

  // --- IMPORTAÇÃO 1: PREVENTIVAS ---
  const lidarComImportacaoPreventivas = (e) => {
    const arquivo = e.target.files[0];
    if (!arquivo) return;
    if (
      substituirFrota &&
      !confirm(
        '⚠️ ATENÇÃO MÁXIMA: Você marcou para substituir a malha ativa. Isso limpará todos os registros existentes. Deseja prosseguir?'
      )
    ) {
      e.target.value = '';
      return;
    }
    setCarregando(true);
    const leitor = new FileReader();
    leitor.onload = async (evento) => {
      try {
        const dadosBinarios = evento.target.result;
        const livro = XLSX.read(dadosBinarios, { type: 'binary' });
        const nomeFolhaAlvo = 'PREVENTIVAS SUL2';
        if (!livro.SheetNames.includes(nomeFolhaAlvo)) {
          throw new Error(
            `Aba obrigatória "${nomeFolhaAlvo}" não foi encontrada!`
          );
        }
        const matrizDados = XLSX.utils.sheet_to_json(
          livro.Sheets[nomeFolhaAlvo],
          { header: 1 }
        );
        const mapaPreventivasUnicas = new Map();
        const hojeStr = new Date().toISOString().split('T')[0];

        for (let i = 1; i < matrizDados.length; i++) {
          const linha = matrizDados[i];
          if (!linha || !linha[0]) continue;

          const placaRaw = String(linha[0]).toUpperCase().trim();
          const filialRaw = String(linha[1] || 'GERAL')
            .toUpperCase()
            .trim();
          const cdvRaw = String(linha[2] || '')
            .toUpperCase()
            .trim();
          const oficinaRaw = String(linha[4] || '').trim();
          const kmRaw = String(linha[7] || '').trim();
          const statusPlanilha = String(linha[10] || '')
            .toUpperCase()
            .trim();
          const obsPlanilha = String(linha[11] || '').trim();
          const obs2Planilha = String(linha[12] || '').trim();
          const mesRefRaw = linha[14]
            ? String(linha[14]).toUpperCase().trim()
            : null;

          const statusFinalBanco =
            statusPlanilha === 'EM DIA'
              ? 'Concluído'
              : statusPlanilha || 'PENDENTE';
          const numMesMapeado = extrairNumeroMesPorExtenso(mesRefRaw);

          const chaveMapa = `${placaRaw}-${numMesMapeado || 'SEM-MES'}-${i}`;
          mapaPreventivasUnicas.set(chaveMapa, {
            placa: placaRaw,
            regional: filialRaw,
            cdv: cdvRaw,
            km: kmRaw || null,
            oficina: oficinaRaw || null,
            status: statusFinalBanco,
            data_conclusao: statusPlanilha === 'EM DIA' ? hojeStr : null,
            observacao: obsPlanilha || null,
            observacao2: obs2Planilha || null,
            mes_referencia: mesRefRaw || null,
            ano_referencia: '2026',
          });
        }
        if (substituirFrota) {
          await supabase
            .from('preventivas')
            .delete()
            .neq('placa', 'REMOVER_TUDO_MASTER');
        }
        await supabase
          .from('preventivas')
          .insert(Array.from(mapaPreventivasUnicas.values()));
        alert('Malha de Preventivas por Mês sincronizada!');
        fetchDados();
      } catch (err) {
        alert(err.message);
      } finally {
        setCarregando(false);
        e.target.value = '';
      }
    };
    leitor.readAsBinaryString(arquivo);
  };

  // --- IMPORTAÇÃO 2: TACOGRAFOS ---
  const lidarComImportacaoTacografos = (e) => {
    const arquivo = e.target.files[0];
    if (!arquivo) return;
    if (
      substituirTaco &&
      !confirm(
        '⚠️ ATENÇÃO MÁXIMA: Você marcou para substituir a malha ativa. Deseja prosseguir?'
      )
    ) {
      e.target.value = '';
      return;
    }
    setCarregando(true);
    const leitor = new FileReader();
    leitor.onload = async (evento) => {
      try {
        const dadosBinarios = evento.target.result;
        const livro = XLSX.read(dadosBinarios, { type: 'binary' });
        const folha = livro.Sheets[livro.SheetNames[0]];
        const matrizDados = XLSX.utils.sheet_to_json(folha, { header: 1 });
        const mapaTacoUnicos = new Map();

        for (let i = 1; i < matrizDados.length; i++) {
          const linha = matrizDados[i];
          if (!linha || !linha[0]) continue;

          const placaRaw = String(linha[0]).toUpperCase().trim();
          const modeloRaw = String(linha[1] || 'CAMINHÃO').trim();
          const cdvRaw = String(linha[2] || '')
            .toUpperCase()
            .trim();
          const filialRaw = String(linha[3] || 'GERAL')
            .toUpperCase()
            .trim();
          const validadeRaw = linha[4]
            ? formatarDataVersaoBancoNoTopo(linha[4], null)
            : null;
          const statusPlanilha = String(linha[5] || '')
            .toUpperCase()
            .trim();
          const mesRefRaw = linha[11]
            ? String(linha[11]).toUpperCase().trim()
            : null;

          const statusFinalBanco =
            statusPlanilha === 'EM DIA'
              ? 'Concluído'
              : statusPlanilha || 'PENDENTE';
          const numMesMapeado = extrairNumeroMesPorExtenso(mesRefRaw);

          const chaveMapa = `${placaRaw}-${numMesMapeado || 'SEM-MES'}-${i}`;
          mapaTacoUnicos.set(chaveMapa, {
            placa: placaRaw,
            modelo: modeloRaw,
            cdv: cdvRaw,
            regional: filialRaw,
            validade: validadeRaw,
            status: statusFinalBanco,
            mes_referencia: mesRefRaw || null,
            ano_referencia: '2026',
          });
        }
        if (substituirTaco) {
          await supabase
            .from('tacografos')
            .delete()
            .neq('placa', 'REMOVER_TUDO_MASTER');
        }
        await supabase
          .from('tacografos')
          .insert(Array.from(mapaTacoUnicos.values()));
        alert('Painel de Tacógrafos atualizado com sucesso!');
        fetchDados();
      } catch (err) {
        alert(err.message);
      } finally {
        setCarregando(false);
        e.target.value = '';
      }
    };
    leitor.readAsBinaryString(arquivo);
  };

  // --- IMPORTAÇÃO 3: FINANCEIRO (LÊ DIVISÃO RATEIO + CTA + BASE (SUL II) + POSIÇÃO NACIONAL) ---
  const lidarComImportacaoFinanceiro = (e) => {
    const arquivo = e.target.files[0];
    if (!arquivo) return;
    if (
      substituirFinanceiro &&
      !confirm(
        '⚠️ Deseja atualizar e substituir a base financeira (DIVISÃO RATEIO + BASE SUL II + CTA)?'
      )
    ) {
      e.target.value = '';
      return;
    }
    setCarregando(true);
    const leitor = new FileReader();
    leitor.onload = async (evento) => {
      try {
        const dadosBinarios = evento.target.result;
        const livro = XLSX.read(dadosBinarios, { type: 'binary' });

        // 1. Validar abas necessárias
        if (!livro.SheetNames.includes('DIVISÃO RATEIO')) {
          throw new Error('A aba "DIVISÃO RATEIO" não foi encontrada na planilha!');
        }
        if (!livro.SheetNames.includes('BASE (SUL II)')) {
          throw new Error('A aba "BASE (SUL II)" não foi encontrada na planilha!');
        }

        // ====================================================================
        // PARTE 0: ATUALIZAR CONTAGEM DE PLACAS DA ABA "Posição Nacional" (SE EXISTIR)
        // ====================================================================
        if (livro.SheetNames.includes('Posição Nacional')) {
          const matrizPN = XLSX.utils.sheet_to_json(
            livro.Sheets['Posição Nacional'],
            { header: 1 }
          );
          const novoMapaCdv = {};
          for (let r = 1; r < matrizPN.length; r++) {
            const row = matrizPN[r];
            if (!row || !row[0]) continue;
            const areaPN = String(row[3] || '').toUpperCase().trim();
            const cdvPN = String(row[2] || '').toUpperCase().trim();
            if (areaPN === 'SUL2' && cdvPN) {
              novoMapaCdv[cdvPN] = (novoMapaCdv[cdvPN] || 0) + 1;
            }
          }
          if (Object.keys(novoMapaCdv).length > 0) {
            setFrotaPorCdv(novoMapaCdv);
          }
        }

        // ====================================================================
        // PARTE A: EXTRAIR ABA "DIVISÃO RATEIO" DIRETAMENTE
        // ====================================================================
        const matrizRateio = XLSX.utils.sheet_to_json(
          livro.Sheets['DIVISÃO RATEIO'],
          { header: 1 }
        );

        let mesRateioPlanilha = 10;
        let areaRateioPlanilha = 'SUL2';
        let linhaCabecalhoRateio = -1;

        for (let r = 0; r < matrizRateio.length; r++) {
          const row = matrizRateio[r];
          if (!row) continue;
          const colB = String(row[1] || '').toUpperCase().trim();
          if (colB === 'MÊS' && row[2]) {
            mesRateioPlanilha = Number(row[2]) || 10;
          }
          if (colB === 'AREA' && row[2]) {
            areaRateioPlanilha = String(row[2]).toUpperCase().trim();
          }
          if (String(row[2] || '').toUpperCase().trim() === 'AOP') {
            linhaCabecalhoRateio = r;
          }
        }

        const listaRateioParaInserir = [];
        if (linhaCabecalhoRateio !== -1) {
          const cabecalho = matrizRateio[linhaCabecalhoRateio] || [];
          const nf1Label = String(cabecalho[5] || 'NF 1/4').trim();
          const nf2Label = String(cabecalho[6] || 'NF 2/4').trim();
          const nf3Label = String(cabecalho[7] || 'NF 3/4').trim();
          const nf4Label = String(cabecalho[8] || 'NF 4/4').trim();

          for (let r = linhaCabecalhoRateio + 1; r < matrizRateio.length; r++) {
            const row = matrizRateio[r];
            if (!row || !row[1]) continue;
            const nomeConta = String(row[1]).trim();
            if (!nomeConta) continue;

            const aopVal = Number(row[2]) || 0;
            const provVal = Number(row[3]) || 0;
            const pendVal = Number(row[4]) || 0;
            const nf1Val = Number(row[5]) || 0;
            const nf2Val = Number(row[6]) || 0;
            const nf3Val = Number(row[7]) || 0;
            const nf4Val = Number(row[8]) || 0;
            const saldoCalculado =
              row[9] !== undefined && row[9] !== null && !isNaN(Number(row[9]))
                ? Number(row[9])
                : aopVal + provVal - (pendVal + nf1Val + nf2Val + nf3Val + nf4Val);

            listaRateioParaInserir.push({
              mes: mesRateioPlanilha,
              area: areaRateioPlanilha,
              conta: nomeConta.toUpperCase(),
              aop: aopVal,
              prov: provVal,
              pend: pendVal,
              nf1: nf1Val,
              nf2: nf2Val,
              nf3: nf3Val,
              nf4: nf4Val,
              nf1_label: nf1Label,
              nf2_label: nf2Label,
              nf3_label: nf3Label,
              nf4_label: nf4Label,
              saldo: saldoCalculado,
            });

            if (nomeConta.toUpperCase().includes('TOTAL GERAL')) break;
          }
        }

        // ====================================================================
        // PARTE B: EXTRAIR ABA "CTA" (CALL TO ACTION - ALERTAS CRÍTICOS)
        // ====================================================================
        const listaCtaParaInserir = [];
        if (livro.SheetNames.includes('CTA')) {
          const matrizCta = XLSX.utils.sheet_to_json(livro.Sheets['CTA'], {
            header: 1,
          });
          for (let r = 0; r < matrizCta.length; r++) {
            const row = matrizCta[r];
            if (!row || !row[0]) continue;
            const col0 = String(row[0]).toUpperCase().trim();
            if (
              col0 === 'PLACA' ||
              col0.includes('CALL TO ACTION') ||
              col0 === ''
            )
              continue;

            listaCtaParaInserir.push({
              placa: col0,
              servico: String(row[1] || '').trim(),
              oficina: String(row[2] || '').trim(),
              status: String(row[3] || '').toUpperCase().trim(),
              os: String(row[4] || '').trim(),
              valor: Number(row[5]) || 0,
              acao: String(row[6] || '').trim(),
            });
          }
        }

        // ====================================================================
        // PARTE C: EXTRAIR ABA "BASE (SUL II)" (ORDENS DE SERVIÇO)
        // ====================================================================
        const matrizBase = XLSX.utils.sheet_to_json(
          livro.Sheets['BASE (SUL II)'],
          { header: 1 }
        );
        let linhaHeaderBase = 10; // Padrão linha 11 no Excel (índice 10)
        for (let r = 0; r < Math.min(20, matrizBase.length); r++) {
          const row = matrizBase[r];
          if (row && String(row[1] || '').toUpperCase().trim() === 'OS') {
            linhaHeaderBase = r;
            break;
          }
        }

        const listaBaseParaInserir = [];
        for (let r = linhaHeaderBase + 1; r < matrizBase.length; r++) {
          const row = matrizBase[r];
          if (!row) continue;

          const osRaw = row[1] !== undefined && row[1] !== null ? String(row[1]).trim() : '';
          const placaRaw = row[3] !== undefined && row[3] !== null ? String(row[3]).toUpperCase().trim() : '';
          const valorRaw = Number(row[5]) || 0;
          const mesRaw = Number(row[13]) || null;
          const nfRaw = row[15] !== undefined && row[15] !== null ? String(row[15]).trim() : '';

          if (!osRaw && !placaRaw && valorRaw === 0) continue;

          listaBaseParaInserir.push({
            os: osRaw || '-',
            data_os: formatarDataVersaoBancoNoTopo(row[2], null),
            placa: placaRaw || '-',
            estabelecimento: row[4] ? String(row[4]).trim() : '-',
            valor_total: valorRaw,
            ccusto: row[6] ? String(row[6]).trim() : null,
            cdv: row[7] ? String(row[7]).toUpperCase().trim() : 'GERAL',
            area: row[8] ? String(row[8]).toUpperCase().trim() : 'SUL2',
            conta_tl: row[9] ? String(row[9]).trim() : null,
            conta_correcao: row[10] ? String(row[10]).toUpperCase().trim() : 'OUTROS',
            status: row[11] ? String(row[11]).trim() : 'Pendente',
            data_faturamento: formatarDataVersaoBancoNoTopo(row[12], null),
            mes: mesRaw,
            titulo: row[14] && !String(row[14]).includes('#') ? String(row[14]).trim() : null,
            nf: nfRaw,
          });
        }

        // ====================================================================
        // SINCRONIZAÇÃO COM O SUPABASE
        // ====================================================================
        if (substituirFinanceiro) {
          await supabase.from('financeiro_rateio').delete().neq('conta', 'REMOVER_TUDO');
          await supabase.from('financeiro_cta').delete().neq('placa', 'REMOVER_TUDO');
          await supabase.from('financeiro_base').delete().neq('os', 'REMOVER_TUDO');
        }

        if (listaRateioParaInserir.length > 0) {
          const { error: errRat } = await supabase
            .from('financeiro_rateio')
            .insert(listaRateioParaInserir);
          if (errRat) throw errRat;
        }

        if (listaCtaParaInserir.length > 0) {
          const { error: errCta } = await supabase
            .from('financeiro_cta')
            .insert(listaCtaParaInserir);
          if (errCta) throw errCta;
        }

        // Inserir a BASE (SUL II) em lotes de 500 registros para alta velocidade
        const tamanhoLote = 500;
        for (let i = 0; i < listaBaseParaInserir.length; i += tamanhoLote) {
          const lote = listaBaseParaInserir.slice(i, i + tamanhoLote);
          const { error: errB } = await supabase.from('financeiro_base').insert(lote);
          if (errB) throw errB;
        }

        setMesSelecionadoFinanceiro(mesRateioPlanilha);
        alert(
          `✅ Sincronização Financeira Concluída!\n- Divisão Rateio (Mês ${mesRateioPlanilha}): ${listaRateioParaInserir.length} linhas\n- Alertas CTA: ${listaCtaParaInserir.length} veículos críticos\n- Base SUL II: ${listaBaseParaInserir.length} registros importados`
        );
        fetchDados();
      } catch (err) {
        alert('Erro na importação financeira: ' + err.message);
      } finally {
        setCarregando(false);
        e.target.value = '';
      }
    };
    leitor.readAsBinaryString(arquivo);
  };

  const concluirPreventiva = async (id) => {
    await supabase
      .from('preventivas')
      .update({
        status: 'Concluído',
        data_conclusao: new Date().toISOString().split('T')[0],
      })
      .eq('id', id);
    fetchDados();
  };

  const concluirTacografo = async (id) => {
    await supabase
      .from('tacografos')
      .update({ status: 'Concluído' })
      .eq('id', id);
    fetchDados();
  };

  // --- FILTRAGEM DE INFRAESTRUTURA LOGÍSTICA ---
  const prevsFiltradasFila = preventivas.filter(
    (p) => filialAtual === 'Todas' || p.regional === filialAtual
  );
  const tacosFiltradosFila = tacografos.filter(
    (t) => filialAtual === 'Todas' || t.regional === filialAtual
  );

  const listagemPrevsPendentes = prevsFiltradasFila.filter(
    (p) => p.status !== 'Concluído'
  );
  const listagemTacosPendentes = tacosFiltradosFila.filter(
    (t) => t.status !== 'Concluído'
  );

  // ============================================================================
  // --- ABA GERENCIAL 1: REGRAS DE PREVENTIVAS ---
  // ============================================================================
  const preventivasDoMesHistorico = prevsFiltradasFila.filter((p) => {
    if (!p.mes_referencia) return false;
    return (
      extrairNumeroMesPorExtenso(p.mes_referencia) ===
      Number(mesSelecionadoHistorico)
    );
  });

  const preventivasSemDataFilaGeral = prevsFiltradasFila.filter(
    (p) => !p.mes_referencia || String(p.mes_referencia).trim() === ''
  );

  const contagemPorMesPrev = mesesDoAno.map((m) => {
    const qtd = prevsFiltradasFila.filter(
      (p) => extrairNumeroMesPorExtenso(p.mes_referencia) === m.num
    ).length;
    return { ...m, qtd };
  });

  const maxQtdMesesPrev = Math.max(
    ...contagemPorMesPrev.map((m) => m.qtd),
    preventivasSemDataFilaGeral.length,
    1
  );

  const totalGeralMes = preventivasDoMesHistorico.length || 1;
  const qtdEmDiaComData = preventivasDoMesHistorico.filter(
    (p) => p.status === 'Concluído' || p.status === 'EM DIA'
  ).length;
  const qtdExecucaoComData = preventivasDoMesHistorico.filter(
    (p) => p.status === 'EM EXECUÇÃO' || p.status === 'OFICINA'
  ).length;
  const qtdAVencerComData = preventivasDoMesHistorico.filter(
    (p) => p.status === 'A VENCER'
  ).length;
  const qtdVencidasComData = preventivasDoMesHistorico.filter(
    (p) =>
      p.status === 'VENCIDA' ||
      p.status === 'VENCIDAS' ||
      p.status === 'VENCIDO'
  ).length;

  const pctEmDia = ((qtdEmDiaComData / totalGeralMes) * 100).toFixed(1);
  const pctExecucao = ((qtdExecucaoComData / totalGeralMes) * 100).toFixed(1);
  const pctAVencer = ((qtdAVencerComData / totalGeralMes) * 100).toFixed(1);
  const pctVencidas = ((qtdVencidasComData / totalGeralMes) * 100).toFixed(1);

  const listagemFiltradaCardsChefe = (
    statusFiltroCard === 'SEM DATA'
      ? preventivasSemDataFilaGeral
      : preventivasDoMesHistorico
  ).filter((p) => {
    let passaCard = true;
    if (statusFiltroCard === 'EM DIA')
      passaCard = p.status === 'Concluído' || p.status === 'EM DIA';
    else if (statusFiltroCard === 'EM EXECUÇÃO')
      passaCard = p.status === 'EM EXECUÇÃO' || p.status === 'OFICINA';
    else if (statusFiltroCard === 'A VENCER')
      passaCard = p.status === 'A VENCER';
    else if (statusFiltroCard === 'VENCIDAS')
      passaCard =
        p.status === 'VENCIDA' ||
        p.status === 'VENCIDAS' ||
        p.status === 'VENCIDO';

    let passaBusca =
      !termoBuscaGerencial ||
      p.placa.includes(termoBuscaGerencial.toUpperCase()) ||
      p.cdv.includes(termoBuscaGerencial.toUpperCase());
    return passaCard && passaBusca;
  });

  // ============================================================================
  // --- ABA GERENCIAL 2: REGRAS DE TACOGRAFOS ---
  // ============================================================================
  const tacosDoMesHistorico = tacosFiltradosFila.filter((t) => {
    if (!t.mes_referencia) return false;
    return (
      extrairNumeroMesPorExtenso(t.mes_referencia) ===
      Number(mesSelecionadoTacoGerencial)
    );
  });

  const tacosAnosFuturos = tacosFiltradosFila.filter((t) => {
    if (!t.mes_referencia) return false;
    const ref = String(t.mes_referencia).trim();
    return /^\d{4}$/.test(ref) && Number(ref) > 2026;
  });

  const tacosSemDataFilaGeral = tacosFiltradosFila.filter(
    (t) => !t.mes_referencia || String(t.mes_referencia).trim() === ''
  );

  const contagemPorMesTaco = mesesDoAno.map((m) => {
    const qtd = tacosFiltradosFila.filter(
      (t) => extrairNumeroMesPorExtenso(t.mes_referencia) === m.num
    ).length;
    return { ...m, qtd };
  });

  const maxQtdMesesTaco = Math.max(
    ...contagemPorMesTaco.map((m) => m.qtd),
    tacosAnosFuturos.length,
    tacosSemDataFilaGeral.length,
    1
  );

  const totalTacosGeralMes = tacosDoMesHistorico.length || 1;
  const qtdTacoEmDiaComData = tacosDoMesHistorico.filter(
    (t) => t.status === 'Concluído' || t.status === 'EM DIA'
  ).length;
  const qtdTacoExecucaoComData = tacosDoMesHistorico.filter(
    (t) => t.status === 'EM EXECUÇÃO' || t.status === 'OFICINA'
  ).length;
  const qtdTacoAVencerComData = tacosDoMesHistorico.filter(
    (t) => t.status === 'A VENCER'
  ).length;
  const qtdTacoVencidasComData = tacosDoMesHistorico.filter(
    (t) =>
      t.status === 'VENCIDA' ||
      t.status === 'VENCIDAS' ||
      t.status === 'VENCIDO'
  ).length;

  const pctTacoEmDia = (
    (qtdTacoEmDiaComData / totalTacosGeralMes) *
    100
  ).toFixed(1);
  const pctTacoExecucao = (
    (qtdTacoExecucaoComData / totalTacosGeralMes) *
    100
  ).toFixed(1);
  const pctTacoAVencer = (
    (qtdTacoAVencerComData / totalTacosGeralMes) *
    100
  ).toFixed(1);
  const pctTacoVencidas = (
    (qtdTacoVencidasComData / totalTacosGeralMes) *
    100
  ).toFixed(1);

  let baseTacosLista = tacosDoMesHistorico;
  if (statusFiltroCardTaco === 'SEM DATA') {
    baseTacosLista = tacosSemDataFilaGeral;
  } else if (statusFiltroCardTaco === 'FUTUROS') {
    baseTacosLista = tacosAnosFuturos;
  }

  const listagemFiltradaTacosChefe = baseTacosLista.filter((t) => {
    let passaCard = true;
    if (statusFiltroCardTaco === 'EM DIA')
      passaCard = t.status === 'Concluído' || t.status === 'EM DIA';
    else if (statusFiltroCardTaco === 'EM EXECUÇÃO')
      passaCard = t.status === 'EM EXECUÇÃO' || t.status === 'OFICINA';
    else if (statusFiltroCardTaco === 'A VENCER')
      passaCard = t.status === 'A VENCER';
    else if (statusFiltroCardTaco === 'VENCIDAS')
      passaCard =
        t.status === 'VENCIDA' ||
        t.status === 'VENCIDAS' ||
        t.status === 'VENCIDO';

    let passaBusca =
      !termoBuscaGerencialTaco ||
      t.placa.includes(termoBuscaGerencialTaco.toUpperCase()) ||
      t.cdv.includes(termoBuscaGerencialTaco.toUpperCase());
    return passaCard && passaBusca;
  });

  // ============================================================================
  // --- ABA GERENCIAL 3: REGRAS DE STATUS FINANCEIRO (DIVISÃO RATEIO + BASE) ---
  // ============================================================================
  const osReaisBase = finBase.filter(
    (item) =>
      item.nf !== 'AOP' &&
      item.nf !== 'PROV' &&
      item.os !== 'AOP' &&
      item.os !== 'PROV' &&
      item.os !== '-'
  );

  const mesOficialRateio =
    finRateio.length > 0 ? Number(finRateio[0].mes) : null;
  const usandoRateioDireto =
    mesOficialRateio === Number(mesSelecionadoFinanceiro);

  const contasPadrao = ['FRETE', 'LAVAGEM', 'MANUTENÇÃO', 'PNEUS', 'SINISTRO'];

  let linhasMatrizRateio = [];
  let labelNf1 = finRateio[0]?.nf1_label || 'NF 1/4';
  let labelNf2 = finRateio[0]?.nf2_label || 'NF 2/4';
  let labelNf3 = finRateio[0]?.nf3_label || 'NF 3/4';
  let labelNf4 = finRateio[0]?.nf4_label || 'NF 4/4';

  if (usandoRateioDireto && finRateio.length > 0) {
    linhasMatrizRateio = finRateio;
  } else {
    const registrosMes = finBase.filter(
      (b) => Number(b.mes) === Number(mesSelecionadoFinanceiro)
    );
    const nfsDoMes = Array.from(
      new Set(
        registrosMes
          .map((b) => b.nf)
          .filter((nf) => nf && nf !== 'AOP' && nf !== 'PROV' && nf !== 'PEND')
      )
    ).sort();

    labelNf1 = nfsDoMes[0] || '1/4 NF';
    labelNf2 = nfsDoMes[1] || '2/4 NF';
    labelNf3 = nfsDoMes[2] || '3/4 NF';
    labelNf4 = nfsDoMes[3] || '4/4 NF';

    let somaAop = 0,
      somaProv = 0,
      somaPend = 0,
      somaNf1 = 0,
      somaNf2 = 0,
      somaNf3 = 0,
      somaNf4 = 0;

    linhasMatrizRateio = contasPadrao.map((conta) => {
      const regConta = registrosMes.filter((b) => b.conta_correcao === conta);
      const regAopComData = regConta.filter(
        (b) => b.nf === 'AOP' && b.data_faturamento
      );
      const regAop =
        regAopComData.length > 0
          ? regAopComData
          : regConta.filter((b) => b.nf === 'AOP');

      const aop = regAop.reduce((acc, r) => acc + (Number(r.valor_total) || 0), 0);
      const prov = regConta
        .filter((b) => b.nf === 'PROV')
        .reduce((acc, r) => acc + (Number(r.valor_total) || 0), 0);
      const pend = regConta
        .filter((b) => b.nf === 'PEND')
        .reduce((acc, r) => acc + (Number(r.valor_total) || 0), 0);
      const nf1 = nfsDoMes[0]
        ? regConta
            .filter((b) => b.nf === nfsDoMes[0])
            .reduce((acc, r) => acc + (Number(r.valor_total) || 0), 0)
        : 0;
      const nf2 = nfsDoMes[1]
        ? regConta
            .filter((b) => b.nf === nfsDoMes[1])
            .reduce((acc, r) => acc + (Number(r.valor_total) || 0), 0)
        : 0;
      const nf3 = nfsDoMes[2]
        ? regConta
            .filter((b) => b.nf === nfsDoMes[2])
            .reduce((acc, r) => acc + (Number(r.valor_total) || 0), 0)
        : 0;
      const nf4 = nfsDoMes[3]
        ? regConta
            .filter((b) => b.nf === nfsDoMes[3])
            .reduce((acc, r) => acc + (Number(r.valor_total) || 0), 0)
        : 0;

      const saldo = aop + prov - (pend + nf1 + nf2 + nf3 + nf4);

      somaAop += aop;
      somaProv += prov;
      somaPend += pend;
      somaNf1 += nf1;
      somaNf2 += nf2;
      somaNf3 += nf3;
      somaNf4 += nf4;

      return { conta, aop, prov, pend, nf1, nf2, nf3, nf4, saldo };
    });

    linhasMatrizRateio.push({
      conta: 'TOTAL GERAL',
      aop: somaAop,
      prov: somaProv,
      pend: somaPend,
      nf1: somaNf1,
      nf2: somaNf2,
      nf3: somaNf3,
      nf4: somaNf4,
      saldo:
        somaAop +
        somaProv -
        (somaPend + somaNf1 + somaNf2 + somaNf3 + somaNf4),
    });
  }

  const linhaTotalGeralRateio = linhasMatrizRateio.find((l) =>
    String(l.conta).toUpperCase().includes('TOTAL GERAL')
  ) || {
    aop: 0,
    prov: 0,
    pend: 0,
    nf1: 0,
    nf2: 0,
    nf3: 0,
    nf4: 0,
    saldo: 0,
  };

  const totalOrcamentoMes =
    (Number(linhaTotalGeralRateio.aop) || 0) +
    (Number(linhaTotalGeralRateio.prov) || 0);
  const totalFaturadoNFsMes =
    (Number(linhaTotalGeralRateio.nf1) || 0) +
    (Number(linhaTotalGeralRateio.nf2) || 0) +
    (Number(linhaTotalGeralRateio.nf3) || 0) +
    (Number(linhaTotalGeralRateio.nf4) || 0);
  const totalPendenteMes = Number(linhaTotalGeralRateio.pend) || 0;
  const saldoFinalMes = Number(linhaTotalGeralRateio.saldo) || 0;

  // Evolução Anual de Gastos Reais (Jan a Dez) para o Gráfico Financeiro
  const gastosPorMesAnual = mesesDoAno.map((m) => {
    const osMes = osReaisBase.filter((b) => Number(b.mes) === m.num);
    const valorGasto = osMes.reduce(
      (acc, item) => acc + (Number(item.valor_total) || 0),
      0
    );
    return { ...m, valorGasto, qtdOs: osMes.length };
  });

  const maxValorMesFinanceiro = Math.max(
    ...gastosPorMesAnual.map((m) => m.valorGasto),
    1
  );

  // ============================================================================
  // --- CÁLCULO DO RATEIO POR CDV (POSIÇÃO NACIONAL vs GASTO REAL BASE SUL II) ---
  // ============================================================================
  const totalPlacasSul2 =
    Object.values(frotaPorCdv).reduce((acc, qtd) => acc + qtd, 0) || 485;

  const osDoMesSelecionado = osReaisBase.filter(
    (b) => Number(b.mes) === Number(mesSelecionadoFinanceiro)
  );

  const totalGastoGeralMes =
    osDoMesSelecionado.reduce(
      (acc, b) => acc + (Number(b.valor_total) || 0),
      0
    ) || 1;

  const todosCdvsMes = Array.from(
    new Set([
      ...Object.keys(frotaPorCdv),
      ...osDoMesSelecionado.map((b) =>
        String(b.cdv || 'GERAL')
          .toUpperCase()
          .trim()
      ),
    ])
  );

  const analisePorCdv = todosCdvsMes
    .map((cdvNome) => {
      const qtdPlacas = frotaPorCdv[cdvNome] || 0;
      const pctFrota = (qtdPlacas / totalPlacasSul2) * 100;
      const orcamentoCdv = totalOrcamentoMes * (qtdPlacas / totalPlacasSul2);

      const osDesteCdv = osDoMesSelecionado.filter(
        (b) =>
          String(b.cdv || 'GERAL')
            .toUpperCase()
            .trim() === cdvNome
      );
      const gastoReal = osDesteCdv.reduce(
        (acc, b) => acc + (Number(b.valor_total) || 0),
        0
      );

      // % consumido do próprio orçamento (Ex: 109.22%) e quanto passou de 100% (Ex: +9.22%)
      const pctDoOrcamento =
        orcamentoCdv > 0 ? (gastoReal / orcamentoCdv) * 100 : 0;
      const pctEstouro = orcamentoCdv > 0 ? pctDoOrcamento - 100 : 0;
      const saldoCdv = orcamentoCdv - gastoReal;

      // % que este CDV representa no gasto total de todos os CDVs (Quem gasta mais)
      const shareGastoTotal = (gastoReal / totalGastoGeralMes) * 100;

      return {
        cdv: cdvNome,
        placas: qtdPlacas,
        pctFrota,
        orcamentoCdv,
        gastoReal,
        pctDoOrcamento,
        pctEstouro,
        saldoCdv,
        shareGastoTotal,
        qtdOs: osDesteCdv.length,
      };
    })
    .sort((a, b) => b.gastoReal - a.gastoReal);

  // Listagem de OSs da BASE (SUL II) filtrada pelo Mês, Conta, CDV clicado e Busca
  const listagemOsFinanceiroFiltrada = osReaisBase
    .filter((b) => Number(b.mes) === Number(mesSelecionadoFinanceiro))
    .filter(
      (b) =>
        contaFiltroFinanceiro === 'TODAS' ||
        String(b.conta_correcao).toUpperCase() === contaFiltroFinanceiro
    )
    .filter(
      (b) =>
        cdvFiltroFinanceiro === 'TODOS' ||
        String(b.cdv || 'GERAL')
          .toUpperCase()
          .trim() === cdvFiltroFinanceiro
    )
    .filter(
      (b) =>
        nfFiltroFinanceiro === 'TODAS' ||
        (nfFiltroFinanceiro === 'PEND' ? b.nf === 'PEND' : b.nf !== 'PEND')
    )
    .filter((b) => {
      if (!termoBuscaFinanceiro) return true;
      const t = termoBuscaFinanceiro.toUpperCase();
      return (
        String(b.placa || '').toUpperCase().includes(t) ||
        String(b.os || '').toUpperCase().includes(t) ||
        String(b.cdv || '').toUpperCase().includes(t) ||
        String(b.estabelecimento || '').toUpperCase().includes(t) ||
        String(b.nf || '').toUpperCase().includes(t)
      );
    });

  if (!autenticado) {
    return (
      <div className="min-h-screen bg-[#0B1329] flex flex-col items-center justify-center font-sans antialiased text-gray-100 p-4">
        <div className="w-full max-w-md bg-[#1C2541]/90 border border-gray-800 rounded-3xl p-8 shadow-2xl relative z-10">
          <div className="flex flex-col items-center mb-8">
            <div className="w-16 h-16 bg-gradient-to-br from-[#0000FF] to-[#FFA500] rounded-2xl flex items-center justify-center text-white font-black text-2xl mb-4">
              PF
            </div>
            <h1 className="text-2xl font-black text-white uppercase tracking-tight">
              Painel Frota PepsiCo
            </h1>
            <p className="text-xs text-gray-400 mt-1">
              Acesso à Malha Logística Assinado
            </p>
          </div>
          <form onSubmit={lidarComLogin} className="space-y-5">
            <input
              type="text"
              placeholder="Usuário (Ex: SUL2CO)"
              value={usuarioInput}
              onChange={(e) => setUsuarioInput(e.target.value)}
              className="w-full bg-[#0B1329] border border-gray-800 rounded-xl px-4 py-3 text-sm font-bold text-white focus:outline-none"
              required
            />
            <input
              type="password"
              placeholder="Senha"
              value={senhaInput}
              onChange={(e) => setSenhaInput(e.target.value)}
              className="w-full bg-[#0B1329] border border-gray-800 rounded-xl px-4 py-3 text-sm font-bold text-white focus:outline-none"
              required
            />
            {erroLogin && (
              <p className="text-xs font-bold text-red-400">{erroLogin}</p>
            )}
            <button
              type="submit"
              className="w-full bg-gradient-to-r from-orange-500 to-amber-500 text-white font-black text-xs uppercase py-3.5 rounded-xl shadow-md"
            >
              Autenticar Sistema
            </button>
          </form>
        </div>
        <div className="mt-8 text-center">
          <p className="text-[11px] font-bold text-gray-500 uppercase tracking-widest">
            Operado por{' '}
            <span className="text-orange-400 font-black">Opentech</span>
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className="flex h-screen font-sans bg-[#0B1329] text-gray-100 relative overflow-hidden">
      <aside className="w-64 border-r border-gray-800 flex flex-col bg-[#1C2541]/90">
        <div className="h-16 flex items-center px-6 border-b border-gray-800">
          <div className="w-8 h-8 bg-gradient-to-br from-[#0000FF] to-[#FFA500] rounded flex items-center justify-center text-white font-black text-sm">
            PF
          </div>
          <span className="font-bold text-lg ml-3 tracking-tight text-white">
            Painel Frota
          </span>
        </div>
        <div className="p-4 border-b border-gray-800">
          <label className="text-[10px] font-bold text-gray-400 uppercase tracking-wider">
            Unidade / Filial
          </label>
          <select
            className="w-full bg-[#0B1329] border border-gray-700 text-xs font-bold p-2 mt-1 rounded-lg uppercase text-white outline-none"
            value={filialAtual}
            onChange={(e) => setFilialAtual(e.target.value)}
          >
            <option value="Todas">Visão Geral (Todas)</option>
            {Array.from(new Set(preventivas.map((p) => p.regional)))
              .filter(Boolean)
              .map((f) => (
                <option key={f} value={f}>
                  {f}
                </option>
              ))}
          </select>
        </div>
        <nav className="p-3 space-y-1 flex-1">
          <button
            onClick={() => setActiveTab('frota')}
            className={`w-full flex items-center gap-3 p-3 rounded-xl text-xs font-bold ${
              activeTab === 'frota'
                ? 'bg-blue-50/10 text-[#0000FF]'
                : 'text-gray-400'
            }`}
          >
            <CalendarDays size={18} />
            <span>Fila de Oficina</span>
          </button>
          <button
            onClick={() => setActiveTab('historico_mensal')}
            className={`w-full flex items-center gap-3 p-3 rounded-xl text-xs font-bold ${
              activeTab === 'historico_mensal'
                ? 'bg-emerald-50/10 text-emerald-400'
                : 'text-gray-400'
            }`}
          >
            <History size={18} />
            <span>Status Preventivas</span>
          </button>
          <button
            onClick={() => setActiveTab('status_tacografos')}
            className={`w-full flex items-center gap-3 p-3 rounded-xl text-xs font-bold ${
              activeTab === 'status_tacografos'
                ? 'bg-amber-50/10 text-amber-400'
                : 'text-gray-400'
            }`}
          >
            <Clock size={18} />
            <span>Status Tacógrafos</span>
          </button>

          {/* NOVA ABA: STATUS FINANCEIRO */}
          <button
            onClick={() => setActiveTab('status_financeiro')}
            className={`w-full flex items-center justify-between p-3 rounded-xl text-xs font-bold transition-all ${
              activeTab === 'status_financeiro'
                ? 'bg-rose-500/15 text-rose-400 border border-rose-500/30'
                : 'text-gray-400 hover:bg-gray-800/40'
            }`}
          >
            <div className="flex items-center gap-3">
              <DollarSign size={18} className="text-rose-400" />
              <span>Status Financeiro</span>
            </div>
            {finCta.length > 0 && (
              <span className="bg-red-600 text-white text-[9px] font-black px-1.5 py-0.5 rounded-full animate-pulse">
                {finCta.length} CTA
              </span>
            )}
          </button>

          <a
            href={LINK_POWER_BI}
            target="_blank"
            rel="noopener noreferrer"
            className="w-full flex items-center gap-3 p-3 rounded-xl text-gray-400 hover:bg-emerald-500/10 hover:text-emerald-400 transition-all group border border-transparent hover:border-emerald-500/20"
          >
            <BarChart3
              size={18}
              className="text-emerald-500 group-hover:animate-pulse"
            />
            <span className="flex items-center gap-1 font-bold text-xs">
              {' '}
              BI de Multas <ExternalLink
                size={12}
                className="opacity-60"
              />{' '}
            </span>
          </a>
        </nav>
        <div className="p-4 border-t border-gray-800 text-center">
          <p className="text-[9px] font-bold text-gray-500 uppercase tracking-wider">
            Opentech @ 2026
          </p>
        </div>
      </aside>

      <main className="flex-1 overflow-y-auto p-8 relative">
        <div className="fixed top-4 right-4 flex items-center gap-2 z-50">
          <button
            onClick={fetchDados}
            className="border border-gray-700 bg-[#1C2541] p-2 rounded-full text-gray-200"
          >
            <RefreshCw
              size={14}
              className={rotacionar ? 'animate-spin text-[#FFA500]' : ''}
            />
          </button>
        </div>

        {activeTab === 'frota' && (
          <div className="max-w-6xl mx-auto space-y-6 animate-fadeIn">
            <div className="flex justify-between items-center border-b border-gray-800 pb-4">
              <h2 className="text-xl font-bold">
                Painel Operacional Logístico
              </h2>
              <div className="flex p-1 rounded-xl border border-gray-800 bg-[#111A33]">
                <button
                  onClick={() => setSubAbaFrota('preventivas')}
                  className={`px-3 py-1.5 rounded-lg text-xs font-bold ${
                    subAbaFrota === 'preventivas'
                      ? 'bg-orange-500/10 text-[#FFA500]'
                      : 'text-gray-400'
                  }`}
                >
                  🛠️ Preventivas ({listagemPrevsPendentes.length})
                </button>
                <button
                  onClick={() => setSubAbaFrota('tacografos')}
                  className={`px-3 py-1.5 rounded-lg text-xs font-bold ${
                    subAbaFrota === 'tacografos'
                      ? 'bg-blue-500/10 text-[#0000FF]'
                      : 'text-gray-400'
                  }`}
                >
                  ⏱️ Tacógrafos ({listagemTacosPendentes.length})
                </button>
              </div>
            </div>

            {subAbaFrota === 'preventivas' && (
              <div className="space-y-6">
                <div className="p-5 bg-[#1C2541] border border-gray-800 rounded-2xl flex justify-between items-center gap-4">
                  <div className="flex-1">
                    <h3 className="text-sm font-black uppercase text-orange-400">
                      🛠️ Importar Excel de preventivas
                    </h3>
                    <p className="text-[11px] opacity-70 mt-1">
                      Selecione a opção abaixo para evitar placas duplicadas.
                    </p>
                    <label className="flex items-center gap-2 mt-2 cursor-pointer text-xs text-gray-400">
                      <input
                        type="checkbox"
                        checked={substituirFrota}
                        onChange={(e) => setSubstituirFrota(e.target.checked)}
                        className="accent-orange-500"
                      />
                      Substituir arquivo atual
                    </label>
                  </div>
                  <div className="flex flex-col sm:flex-row gap-2 shrink-0">
                    <button
                      onClick={limparTodaTabelaPreventivas}
                      className="bg-red-600/10 text-red-500 hover:bg-red-600 hover:text-white border border-red-500/20 text-xs font-black px-4 py-2.5 rounded-xl transition-all flex items-center gap-1.5"
                    >
                      <Trash2 size={14} /> Limpar Banco
                    </button>
                    <label className="bg-orange-500 hover:bg-orange-600 text-white font-black text-xs px-4 py-2.5 rounded-xl cursor-pointer flex items-center gap-1.5 transition-all">
                      <Upload size={14} />
                      Carregar Excel
                      <input
                        type="file"
                        accept=".xlsx, .xls"
                        onChange={lidarComImportacaoPreventivas}
                        className="hidden"
                      />
                    </label>
                  </div>
                </div>
                <div className="p-4 bg-[#111A33] border border-gray-800 rounded-xl flex items-center gap-3">
                  <Search size={18} className="text-gray-400" />
                  <input
                    type="text"
                    value={termoBusca}
                    onChange={(e) => setTermoBusca(e.target.value)}
                    placeholder="Pesquisar preventiva ativa..."
                    className="bg-transparent w-full outline-none text-xs text-white"
                  />
                </div>
                <div className="rounded-2xl border border-gray-800 bg-[#1C2541] overflow-hidden">
                  <div className="max-h-[500px] overflow-y-auto">
                    <table className="w-full text-left text-xs table-auto">
                      <thead className="sticky top-0 bg-[#111A33] font-bold text-gray-400 uppercase text-[10px] border-b border-gray-800">
                        <tr>
                          <th className="p-4">Placa</th>
                          <th className="p-4">Filial</th>
                          <th className="p-4">CDV</th>
                          <th className="p-4">KM</th>
                          <th className="p-4 text-orange-400">Mês Meta</th>
                          <th className="p-4">Oficina</th>
                          <th className="p-4">Obs 1</th>
                          <th className="p-4">Obs 2</th>
                          <th className="p-4 text-right">Ação</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gray-800">
                        {listagemPrevsPendentes
                          .filter(
                            (p) =>
                              !termoBusca ||
                              p.placa.includes(termoBusca.toUpperCase()) ||
                              p.cdv.includes(termoBusca.toUpperCase())
                          )
                          .map((p) => (
                            <tr key={p.id} className="hover:bg-gray-800/5">
                              <td className="p-4 font-mono font-bold">
                                {p.placa}
                              </td>
                              <td className="p-4 text-blue-400 uppercase">
                                {p.regional}
                              </td>
                              <td className="p-4 uppercase">{p.cdv}</td>
                              <td className="p-4">{p.km || '--'}</td>
                              <td className="p-4 text-orange-400 font-black uppercase">
                                {p.mes_referencia || 'SEM MÊS'}
                              </td>
                              <td className="p-4">{p.oficina || '--'}</td>
                              <td className="p-4 text-gray-400 italic">
                                {p.observacao || '--'}
                              </td>
                              <td className="p-4 text-gray-400 italic">
                                {p.observacao2 || '--'}
                              </td>
                              <td className="p-4 text-right">
                                <button
                                  onClick={() => concluirPreventiva(p.id)}
                                  className="bg-emerald-600/10 text-emerald-500 text-[10px] px-2.5 py-1 rounded font-bold"
                                >
                                  Baixa
                                </button>
                              </td>
                            </tr>
                          ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}

            {subAbaFrota === 'tacografos' && (
              <div className="space-y-6">
                <div className="p-5 bg-[#1C2541] border border-gray-800 rounded-2xl flex justify-between items-center gap-4">
                  <div className="flex-1">
                    <h3 className="text-sm font-black uppercase text-blue-400">
                      ⏱️ Importar excel de Tacógrafos
                    </h3>
                    <p className="text-[11px] opacity-70 mt-1">
                      Selecione a opção abaixo para evitar placas duplicadas.
                    </p>
                    <label className="flex items-center gap-2 mt-2 cursor-pointer text-xs text-gray-400">
                      <input
                        type="checkbox"
                        checked={substituirTaco}
                        onChange={(e) => setSubstituirTaco(e.target.checked)}
                        className="accent-blue-500"
                      />
                      Substituir arquivo atual
                    </label>
                  </div>
                  <div className="flex flex-col sm:flex-row gap-2 shrink-0">
                    <button
                      onClick={limparTodaTabelaTacografos}
                      className="bg-red-600/10 text-red-500 hover:bg-red-600 hover:text-white border border-red-500/20 text-xs font-black px-4 py-2.5 rounded-xl transition-all flex items-center gap-1.5"
                    >
                      <Trash2 size={14} /> Limpar Banco
                    </button>
                    <label className="bg-blue-600 text-white font-black text-xs px-4 py-2.5 rounded-xl cursor-pointer flex items-center gap-1.5 transition-all">
                      <Upload size={14} />
                      Carregar Excel
                      <input
                        type="file"
                        accept=".xlsx, .xls"
                        onChange={lidarComImportacaoTacografos}
                        className="hidden"
                      />
                    </label>
                  </div>
                </div>
                <div className="p-4 bg-[#111A33] border border-gray-800 rounded-xl flex items-center gap-3">
                  <Search size={18} className="text-gray-400" />
                  <input
                    type="text"
                    value={termoBuscaTaco}
                    onChange={(e) => setTermoBuscaTaco(e.target.value)}
                    placeholder="Pesquisar laudo ativo..."
                    className="bg-transparent w-full outline-none text-xs text-white"
                  />
                </div>
                <div className="rounded-2xl border border-gray-800 bg-[#1C2541] overflow-hidden">
                  <div className="max-h-[500px] overflow-y-auto">
                    <table className="w-full text-left text-xs table-auto">
                      <thead className="sticky top-0 bg-[#111A33] font-bold text-gray-400 uppercase text-[10px] border-b border-gray-800">
                        <tr>
                          <th className="p-4">Placa</th>
                          <th className="p-4">Modelo</th>
                          <th className="p-4">CDV Logístico</th>
                          <th className="p-4 font-bold text-blue-400">
                            Filial / Sede
                          </th>
                          <th className="p-4">Mês / Ano Meta</th>
                          <th className="p-4">Validade Legal</th>
                          <th className="p-4 text-right">Ação</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gray-800">
                        {listagemTacosPendentes
                          .filter(
                            (t) =>
                              !termoBuscaTaco ||
                              t.placa.includes(termoBuscaTaco.toUpperCase()) ||
                              t.cdv.includes(termoBuscaTaco.toUpperCase())
                          )
                          .map((t) => (
                            <tr key={t.id} className="hover:bg-gray-800/5">
                              <td className="p-4 font-mono font-bold">
                                {t.placa}
                              </td>
                              <td className="p-4 text-gray-400 uppercase">
                                {t.modelo || 'CAMINHÃO'}
                              </td>
                              <td className="p-4 uppercase">{t.cdv}</td>
                              <td className="p-4 uppercase text-blue-400">
                                {t.regional}
                              </td>
                              <td className="p-4 text-blue-400 font-black uppercase">
                                {t.mes_referencia || 'SEM MÊS'}
                              </td>
                              <td className="p-4 font-bold">
                                {t.validade
                                  ? new Date(
                                      t.validade + 'T00:00:00'
                                    ).toLocaleDateString('pt-BR')
                                  : '--'}
                              </td>
                              <td className="p-4 text-right">
                                <button
                                  onClick={() => concluirTacografo(t.id)}
                                  className="bg-emerald-600/10 text-emerald-500 text-[10px] px-2.5 py-1 rounded font-bold"
                                >
                                  Regularizar
                                </button>
                              </td>
                            </tr>
                          ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}
          </div>
        )}

        {/* --- TELA 2: ABA STATUS PREVENTIVAS --- */}
        {activeTab === 'historico_mensal' && (
          <div className="max-w-6xl mx-auto animate-fadeIn space-y-6">
            <div className="flex justify-between items-center">
              <div>
                <h2 className="text-2xl font-bold tracking-tight text-white">
                  Cronograma de Preventivas
                </h2>
                <p className="text-gray-400 text-xs mt-0.5">
                  Indicadores operacionais e controle por mês.
                </p>
              </div>
              <div className="flex items-center gap-2 p-2 rounded-xl border border-gray-800 bg-[#1C2541] text-xs font-bold text-white">
                <Clock size={14} className="text-[#FFA500]" />
                <span className="text-gray-400 text-[11px]">Mês:</span>
                <select
                  value={mesSelecionadoHistorico}
                  onChange={(e) => {
                    setMesSelecionadoHistorico(e.target.value);
                    setStatusFiltroCard('TODOS');
                  }}
                  className="bg-transparent outline-none cursor-pointer font-black text-orange-400 uppercase"
                >
                  {mesesDoAno.map((m) => (
                    <option
                      key={m.num}
                      value={m.num}
                      className="bg-[#1C2541] text-white"
                    >
                      {m.nome}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            {/* CARDS MESTRES DE PREVENTIVAS */}
            <div className="grid grid-cols-1 sm:grid-cols-3 lg:grid-cols-5 gap-4">
              <div
                onClick={() =>
                  setStatusFiltroCard(
                    statusFiltroCard === 'EM DIA' ? 'TODOS' : 'EM DIA'
                  )
                }
                className={`p-4 rounded-xl shadow-sm border border-gray-800 cursor-pointer transition-all ${
                  statusFiltroCard === 'EM DIA' ? 'ring-4 ring-green-400' : ''
                } bg-gradient-to-br from-green-500 to-emerald-600 text-white`}
              >
                <div className="flex justify-between items-start">
                  <span className="text-[11px] font-black">EM DIA</span>
                  <span className="text-[10px] bg-white/20 px-1.5 py-0.5 rounded font-bold">
                    {pctEmDia}%
                  </span>
                </div>
                <h3 className="text-3xl font-black mt-3">
                  {String(qtdEmDiaComData).padStart(2, '0')}{' '}
                  <span className="text-[10px] font-normal opacity-80">
                    Placas
                  </span>
                </h3>
              </div>
              <div
                onClick={() =>
                  setStatusFiltroCard(
                    statusFiltroCard === 'EM EXECUÇÃO' ? 'TODOS' : 'EM EXECUÇÃO'
                  )
                }
                className={`p-4 rounded-xl shadow-sm border border-gray-800 cursor-pointer transition-all ${
                  statusFiltroCard === 'EM EXECUÇÃO'
                    ? 'ring-4 ring-orange-400'
                    : ''
                } bg-gradient-to-br from-orange-500 to-amber-600 text-white`}
              >
                <div className="flex justify-between items-start">
                  <span className="text-[11px] font-black">EM EXECUÇÃO</span>
                  <span className="text-[10px] bg-white/20 px-1.5 py-0.5 rounded font-bold">
                    {pctExecucao}%
                  </span>
                </div>
                <h3 className="text-3xl font-black mt-3">
                  {String(qtdExecucaoComData).padStart(2, '0')}{' '}
                  <span className="text-[10px] font-normal opacity-80">
                    Placas
                  </span>
                </h3>
              </div>
              <div
                onClick={() =>
                  setStatusFiltroCard(
                    statusFiltroCard === 'A VENCER' ? 'TODOS' : 'A VENCER'
                  )
                }
                className={`p-4 rounded-xl shadow-sm border border-gray-800 cursor-pointer transition-all ${
                  statusFiltroCard === 'A VENCER'
                    ? 'ring-4 ring-yellow-400'
                    : ''
                } bg-gradient-to-br from-yellow-400 to-yellow-500 text-gray-900`}
              >
                <div className="flex justify-between items-start">
                  <span className="text-[11px] font-black">A VENCER</span>
                  <span className="text-[10px] bg-black/10 px-1.5 py-0.5 rounded font-bold">
                    {pctAVencer}%
                  </span>
                </div>
                <h3 className="text-3xl font-black mt-3">
                  {String(qtdAVencerComData).padStart(2, '0')}{' '}
                  <span className="text-[10px] font-normal opacity-70">
                    Placas
                  </span>
                </h3>
              </div>
              <div
                onClick={() =>
                  setStatusFiltroCard(
                    statusFiltroCard === 'VENCIDAS' ? 'TODOS' : 'VENCIDAS'
                  )
                }
                className={`p-4 rounded-xl shadow-sm border border-gray-800 cursor-pointer transition-all ${
                  statusFiltroCard === 'VENCIDAS' ? 'ring-4 ring-red-500' : ''
                } bg-gradient-to-br from-red-500 to-rose-600 text-white`}
              >
                <div className="flex justify-between items-start">
                  <span className="text-[11px] font-black">VENCIDAS</span>
                  <span className="text-[10px] bg-white/20 px-1.5 py-0.5 rounded font-bold">
                    {pctVencidas}%
                  </span>
                </div>
                <h3 className="text-3xl font-black mt-3">
                  {String(qtdVencidasComData).padStart(2, '0')}{' '}
                  <span className="text-[10px] font-normal opacity-80">
                    Placas
                  </span>
                </h3>
              </div>
              <div
                onClick={() =>
                  setStatusFiltroCard(
                    statusFiltroCard === 'SEM DATA' ? 'TODOS' : 'SEM DATA'
                  )
                }
                className={`p-4 rounded-xl shadow-sm border border-gray-400 cursor-pointer transition-all ${
                  statusFiltroCard === 'SEM DATA' ? 'ring-4 ring-gray-400' : ''
                } bg-gradient-to-br from-gray-600 to-gray-700 text-white`}
              >
                <div className="flex justify-between items-start">
                  <span className="text-[11px] font-black text-gray-300">
                    SEM MÊS
                  </span>
                  <span className="text-[9px] bg-black/20 text-orange-400 px-1 py-0.5 rounded font-black">
                    FILA VIVA
                  </span>
                </div>
                <h3 className="text-3xl font-black mt-3">
                  {String(preventivasSemDataFilaGeral.length).padStart(2, '0')}{' '}
                  <span className="text-[10px] font-normal opacity-80">
                    Placas
                  </span>
                </h3>
              </div>
            </div>

            {/* GRÁFICO ESTILO PREMIUM DE PREVENTIVAS */}
            <div className="p-6 rounded-2xl border border-gray-800/80 bg-[#131B35] shadow-2xl relative overflow-hidden">
              <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-6 gap-2">
                <div>
                  <h3 className="text-lg font-black text-white tracking-tight flex items-center gap-2">
                    <span>📊 Visão Geral</span>
                  </h3>
                  <p className="text-xs text-gray-400">
                    Frota por mês (Clique em uma coluna para selecionar o mês)
                  </p>
                </div>
                <div className="flex items-center gap-2 bg-[#0B1329] p-1.5 rounded-xl border border-gray-800">
                  <span className="text-[10px] font-bold text-gray-400 px-2 uppercase">
                    Visão 2026
                  </span>
                </div>
              </div>

              <div className="relative pt-6 pb-2">
                <div className="absolute inset-0 flex flex-col justify-between pointer-events-none opacity-20 border-b border-gray-700">
                  <div className="border-b border-gray-600 w-full"></div>
                  <div className="border-b border-gray-600 w-full"></div>
                  <div className="border-b border-gray-600 w-full"></div>
                  <div className="border-b border-gray-600 w-full"></div>
                </div>

                <div className="h-48 flex items-end justify-between gap-1.5 sm:gap-2 px-2 relative z-10">
                  {contagemPorMesPrev.map((item) => {
                    const estaSelecionado =
                      Number(mesSelecionadoHistorico) === item.num;
                    const alturaPct = Math.max(
                      (item.qtd / maxQtdMesesPrev) * 100,
                      4
                    );

                    return (
                      <div
                        key={item.num}
                        onClick={() => {
                          setMesSelecionadoHistorico(item.num);
                          setStatusFiltroCard('TODOS');
                        }}
                        className="flex-1 flex flex-col items-center h-full justify-end cursor-pointer group"
                      >
                        <span
                          className={`text-[10px] font-black mb-1 transition-all ${
                            estaSelecionado
                              ? 'text-orange-400 scale-125'
                              : 'text-gray-400 group-hover:text-white'
                          }`}
                        >
                          {item.qtd > 0 ? item.qtd : ''}
                        </span>

                        <div
                          className={`w-full max-w-[36px] rounded-t-lg transition-all duration-300 ${
                            estaSelecionado
                              ? 'bg-gradient-to-t from-rose-600 via-orange-500 to-amber-300 ring-2 ring-orange-400 shadow-lg shadow-orange-500/30'
                              : 'bg-gradient-to-t from-rose-900/60 via-orange-600/70 to-amber-500/80 group-hover:from-rose-600 group-hover:to-amber-400'
                          }`}
                          style={{ height: `${alturaPct}%` }}
                        ></div>

                        <span
                          className={`mt-3 text-[11px] font-bold uppercase transition-all ${
                            estaSelecionado
                              ? 'text-orange-400 font-black'
                              : 'text-gray-400 group-hover:text-white'
                          }`}
                        >
                          {item.sigla}
                        </span>
                      </div>
                    );
                  })}

                  <div
                    onClick={() => setStatusFiltroCard('SEM DATA')}
                    className="flex-1 flex flex-col items-center h-full justify-end cursor-pointer group border-l border-gray-800 pl-1"
                  >
                    <span className="text-[10px] font-black text-gray-400 mb-1 group-hover:scale-110 transition-all">
                      {preventivasSemDataFilaGeral.length > 0
                        ? preventivasSemDataFilaGeral.length
                        : ''}
                    </span>
                    <div
                      className="w-full max-w-[36px] rounded-t-lg bg-gradient-to-t from-gray-700 via-gray-600 to-gray-400 transition-all group-hover:brightness-125"
                      style={{
                        height: `${Math.max(
                          (preventivasSemDataFilaGeral.length /
                            maxQtdMesesPrev) *
                            100,
                          4
                        )}%`,
                      }}
                    ></div>
                    <span className="mt-3 text-[10px] font-bold uppercase text-gray-400">
                      s/m
                    </span>
                  </div>
                </div>
              </div>
            </div>

            <div className="p-4 rounded-xl border border-gray-800 bg-[#111A33] flex items-center gap-3">
              <Search size={18} className="text-gray-400" />
              <input
                type="text"
                value={termoBuscaGerencial}
                onChange={(e) => setTermoBuscaGerencial(e.target.value)}
                placeholder="Filtrar por Placa ou CDV"
                className="bg-transparent w-full outline-none text-xs text-white"
              />
            </div>

            <div className="rounded-2xl border border-gray-800 bg-[#1C2541] overflow-hidden shadow-sm">
              <div className="p-4 border-b border-gray-800 flex justify-between items-center bg-gray-500/5">
                <div className="font-bold text-xs uppercase text-gray-400">
                  Lista:{' '}
                  <span className="text-blue-400 font-black">
                    {statusFiltroCard}
                  </span>{' '}
                  ({listagemFiltradaCardsChefe.length} Resultados)
                </div>
              </div>
              <div className="max-h-[500px] overflow-y-auto">
                <table className="w-full text-left text-xs table-auto">
                  <thead className="sticky top-0 bg-[#111A33] font-bold text-gray-400 uppercase text-[10px] border-b border-gray-800">
                    <tr>
                      <th className="p-4">Placa</th>
                      <th className="p-4">Filial</th>
                      <th className="p-4">CDV</th>
                      <th className="p-4">KM Atual</th>
                      <th className="p-4 text-orange-400">Mês Meta</th>
                      <th className="p-4">Oficina Destino</th>
                      <th className="p-4">Status</th>
                      <th className="p-4">Obs 1</th>
                      <th className="p-4">Obs 2</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-800">
                    {listagemFiltradaCardsChefe.map((p) => (
                      <tr key={`prev-${p.id}`} className="hover:bg-gray-800/10">
                        <td className="p-4 font-mono font-bold text-sm">
                          {p.placa}
                        </td>
                        <td className="p-4 uppercase font-semibold text-blue-400">
                          {p.regional}
                        </td>
                        <td className="p-4 uppercase text-gray-400 font-semibold">
                          {p.cdv}
                        </td>
                        <td className="p-4 font-bold text-gray-400">
                          {p.km || '--'}
                        </td>
                        <td className="p-4 text-orange-400 font-black uppercase">
                          {p.mes_referencia || 'SEM MÊS'}
                        </td>
                        <td className="p-4 uppercase text-gray-500">
                          {p.oficina || '--'}
                        </td>
                        <td className="p-4">
                          <span
                            className={`text-[10px] font-black px-2 py-0.5 rounded ${
                              p.status === 'Concluído' || p.status === 'EM DIA'
                                ? 'bg-green-500/10 text-green-500'
                                : p.status === 'EM EXECUÇÃO' ||
                                  p.status === 'OFICINA'
                                ? 'bg-orange-500/10 text-orange-500'
                                : p.status === 'A VENCER'
                                ? 'bg-yellow-500/10 text-yellow-600'
                                : 'bg-red-50/10 text-red-500'
                            }`}
                          >
                            {p.status === 'Concluído' || p.status === 'EM DIA'
                              ? 'EM DIA'
                              : p.status}
                          </span>
                        </td>
                        <td className="p-4 text-gray-400 italic">
                          {p.observacao || '--'}
                        </td>
                        <td className="p-4 text-gray-400 italic">
                          {p.observacao2 || '--'}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* --- TELA 3: ABA STATUS TACOGRAFOS --- */}
        {activeTab === 'status_tacografos' && (
          <div className="max-w-6xl mx-auto animate-fadeIn space-y-6">
            <div className="flex justify-between items-center">
              <div>
                <h2 className="text-2xl font-bold tracking-tight text-white">
                  Cronograma de Tacógrafos
                </h2>
                <p className="text-gray-400 text-xs mt-0.5">
                  Indicadores operacionais e controle por mês e ano.
                </p>
              </div>
              <div className="flex items-center gap-2 p-2 rounded-xl border border-gray-800 bg-[#1C2541] text-xs font-bold text-white">
                <Clock size={14} className="text-[#0000FF]" />
                <span className="text-gray-400 text-[11px]">Mês:</span>
                <select
                  value={mesSelecionadoTacoGerencial}
                  onChange={(e) => {
                    setMesSelecionadoTacoGerencial(e.target.value);
                    setStatusFiltroCardTaco('TODOS');
                  }}
                  className="bg-transparent outline-none cursor-pointer font-black text-amber-400 uppercase"
                >
                  {mesesDoAno.map((m) => (
                    <option
                      key={m.num}
                      value={m.num}
                      className="bg-[#1C2541] text-white"
                    >
                      {m.nome}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            {/* CARDS MESTRES DE TACOGRAFOS */}
            <div className="grid grid-cols-1 sm:grid-cols-3 lg:grid-cols-6 gap-3">
              <div
                onClick={() =>
                  setStatusFiltroCardTaco(
                    statusFiltroCardTaco === 'EM DIA' ? 'TODOS' : 'EM DIA'
                  )
                }
                className={`p-3.5 rounded-xl shadow-sm border border-gray-800 cursor-pointer transition-all ${
                  statusFiltroCardTaco === 'EM DIA'
                    ? 'ring-4 ring-green-400'
                    : ''
                } bg-gradient-to-br from-green-500 to-emerald-600 text-white`}
              >
                <div className="flex justify-between items-start">
                  <span className="text-[10px] font-black">EM DIA</span>
                  <span className="text-[9px] bg-white/20 px-1 py-0.5 rounded font-bold">
                    {pctTacoEmDia}%
                  </span>
                </div>
                <h3 className="text-2xl font-black mt-2">
                  {String(qtdTacoEmDiaComData).padStart(2, '0')}{' '}
                  <span className="text-[9px] font-normal opacity-80">
                    Placas
                  </span>
                </h3>
              </div>

              <div
                onClick={() =>
                  setStatusFiltroCardTaco(
                    statusFiltroCardTaco === 'EM EXECUÇÃO'
                      ? 'TODOS'
                      : 'EM EXECUÇÃO'
                  )
                }
                className={`p-3.5 rounded-xl shadow-sm border border-gray-800 cursor-pointer transition-all ${
                  statusFiltroCardTaco === 'EM EXECUÇÃO'
                    ? 'ring-4 ring-orange-400'
                    : ''
                } bg-gradient-to-br from-orange-500 to-amber-600 text-white`}
              >
                <div className="flex justify-between items-start">
                  <span className="text-[10px] font-black">EM EXEC.</span>
                  <span className="text-[9px] bg-white/20 px-1 py-0.5 rounded font-bold">
                    {pctTacoExecucao}%
                  </span>
                </div>
                <h3 className="text-2xl font-black mt-2">
                  {String(qtdTacoExecucaoComData).padStart(2, '0')}{' '}
                  <span className="text-[9px] font-normal opacity-80">
                    Placas
                  </span>
                </h3>
              </div>

              <div
                onClick={() =>
                  setStatusFiltroCardTaco(
                    statusFiltroCardTaco === 'A VENCER' ? 'TODOS' : 'A VENCER'
                  )
                }
                className={`p-3.5 rounded-xl shadow-sm border border-gray-800 cursor-pointer transition-all ${
                  statusFiltroCardTaco === 'A VENCER'
                    ? 'ring-4 ring-yellow-400'
                    : ''
                } bg-gradient-to-br from-yellow-400 to-yellow-500 text-gray-900`}
              >
                <div className="flex justify-between items-start">
                  <span className="text-[10px] font-black">A VENCER</span>
                  <span className="text-[9px] bg-black/10 px-1 py-0.5 rounded font-bold">
                    {pctTacoAVencer}%
                  </span>
                </div>
                <h3 className="text-2xl font-black mt-2">
                  {String(qtdTacoAVencerComData).padStart(2, '0')}{' '}
                  <span className="text-[9px] font-normal opacity-70">
                    Placas
                  </span>
                </h3>
              </div>

              <div
                onClick={() =>
                  setStatusFiltroCardTaco(
                    statusFiltroCardTaco === 'VENCIDAS' ? 'TODOS' : 'VENCIDAS'
                  )
                }
                className={`p-3.5 rounded-xl shadow-sm border border-gray-800 cursor-pointer transition-all ${
                  statusFiltroCardTaco === 'VENCIDAS'
                    ? 'ring-4 ring-red-500'
                    : ''
                } bg-gradient-to-br from-red-500 to-rose-600 text-white`}
              >
                <div className="flex justify-between items-start">
                  <span className="text-[10px] font-black">VENCIDAS</span>
                  <span className="text-[9px] bg-white/20 px-1 py-0.5 rounded font-bold">
                    {pctTacoVencidas}%
                  </span>
                </div>
                <h3 className="text-2xl font-black mt-2">
                  {String(qtdTacoVencidasComData).padStart(2, '0')}{' '}
                  <span className="text-[9px] font-normal opacity-80">
                    Placas
                  </span>
                </h3>
              </div>

              <div
                onClick={() =>
                  setStatusFiltroCardTaco(
                    statusFiltroCardTaco === 'FUTUROS' ? 'TODOS' : 'FUTUROS'
                  )
                }
                className={`p-3.5 rounded-xl shadow-sm border border-blue-500/30 cursor-pointer transition-all ${
                  statusFiltroCardTaco === 'FUTUROS'
                    ? 'ring-4 ring-blue-400'
                    : ''
                } bg-gradient-to-br from-blue-600 to-indigo-700 text-white`}
              >
                <div className="flex justify-between items-start">
                  <span className="text-[10px] font-black text-blue-200">
                    2027+ FUTUROS
                  </span>
                  <span className="text-[8px] bg-black/20 text-blue-200 px-1 py-0.5 rounded font-black">
                    VALIDADE
                  </span>
                </div>
                <h3 className="text-2xl font-black mt-2">
                  {String(tacosAnosFuturos.length).padStart(2, '0')}{' '}
                  <span className="text-[9px] font-normal opacity-80">
                    Placas
                  </span>
                </h3>
              </div>

              <div
                onClick={() =>
                  setStatusFiltroCardTaco(
                    statusFiltroCardTaco === 'SEM DATA' ? 'TODOS' : 'SEM DATA'
                  )
                }
                className={`p-3.5 rounded-xl shadow-sm border border-gray-400 cursor-pointer transition-all ${
                  statusFiltroCardTaco === 'SEM DATA'
                    ? 'ring-4 ring-gray-400'
                    : ''
                } bg-gradient-to-br from-gray-600 to-gray-700 text-white`}
              >
                <div className="flex justify-between items-start">
                  <span className="text-[10px] font-black text-gray-300">
                    SEM MÊS
                  </span>
                  <span className="text-[8px] bg-black/20 text-blue-400 px-1 py-0.5 rounded font-black">
                    FILA VIVA
                  </span>
                </div>
                <h3 className="text-2xl font-black mt-2">
                  {String(tacosSemDataFilaGeral.length).padStart(2, '0')}{' '}
                  <span className="text-[9px] font-normal opacity-80">
                    Placas
                  </span>
                </h3>
              </div>
            </div>

            {/* GRÁFICO ESTILO PREMIUM DE TACOGRAFOS */}
            <div className="p-6 rounded-2xl border border-gray-800/80 bg-[#131B35] shadow-2xl relative overflow-hidden">
              <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-6 gap-2">
                <div>
                  <h3 className="text-lg font-black text-white tracking-tight flex items-center gap-2">
                    <span>📊 Visão Geral</span>
                  </h3>
                  <p className="text-xs text-gray-400">
                    Frota por mês (Clique em uma coluna para selecionar o mês)
                  </p>
                </div>
                <div className="flex items-center gap-2 bg-[#0B1329] p-1.5 rounded-xl border border-gray-800">
                  <span className="text-[10px] font-bold text-gray-400 px-2 uppercase">
                    Visão 2026
                  </span>
                </div>
              </div>

              <div className="relative pt-6 pb-2">
                <div className="absolute inset-0 flex flex-col justify-between pointer-events-none opacity-20 border-b border-gray-700">
                  <div className="border-b border-gray-600 w-full"></div>
                  <div className="border-b border-gray-600 w-full"></div>
                  <div className="border-b border-gray-600 w-full"></div>
                  <div className="border-b border-gray-600 w-full"></div>
                </div>

                <div className="h-48 flex items-end justify-between gap-1.5 sm:gap-2 px-2 relative z-10">
                  {contagemPorMesTaco.map((item) => {
                    const estaSelecionado =
                      Number(mesSelecionadoTacoGerencial) === item.num;
                    const alturaPct = Math.max(
                      (item.qtd / maxQtdMesesTaco) * 100,
                      4
                    );

                    return (
                      <div
                        key={item.num}
                        onClick={() => {
                          setMesSelecionadoTacoGerencial(item.num);
                          setStatusFiltroCardTaco('TODOS');
                        }}
                        className="flex-1 flex flex-col items-center h-full justify-end cursor-pointer group"
                      >
                        <span
                          className={`text-[10px] font-black mb-1 transition-all ${
                            estaSelecionado
                              ? 'text-amber-400 scale-125'
                              : 'text-gray-400 group-hover:text-white'
                          }`}
                        >
                          {item.qtd > 0 ? item.qtd : ''}
                        </span>

                        <div
                          className={`w-full max-w-[36px] rounded-t-lg transition-all duration-300 ${
                            estaSelecionado
                              ? 'bg-gradient-to-t from-rose-600 via-orange-500 to-amber-300 ring-2 ring-amber-400 shadow-lg shadow-orange-500/30'
                              : 'bg-gradient-to-t from-rose-900/60 via-orange-600/70 to-amber-500/80 group-hover:from-rose-600 group-hover:to-amber-400'
                          }`}
                          style={{ height: `${alturaPct}%` }}
                        ></div>

                        <span
                          className={`mt-3 text-[11px] font-bold uppercase transition-all ${
                            estaSelecionado
                              ? 'text-amber-400 font-black'
                              : 'text-gray-400 group-hover:text-white'
                          }`}
                        >
                          {item.sigla}
                        </span>
                      </div>
                    );
                  })}

                  <div
                    onClick={() => setStatusFiltroCardTaco('FUTUROS')}
                    className="flex-1 flex flex-col items-center h-full justify-end cursor-pointer group border-l border-gray-800 pl-1"
                  >
                    <span className="text-[10px] font-black text-blue-400 mb-1 group-hover:scale-110 transition-all">
                      {tacosAnosFuturos.length > 0
                        ? tacosAnosFuturos.length
                        : ''}
                    </span>
                    <div
                      className="w-full max-w-[36px] rounded-t-lg bg-gradient-to-t from-blue-900 via-indigo-600 to-cyan-400 transition-all group-hover:brightness-125"
                      style={{
                        height: `${Math.max(
                          (tacosAnosFuturos.length / maxQtdMesesTaco) * 100,
                          4
                        )}%`,
                      }}
                    ></div>
                    <span className="mt-3 text-[10px] font-black uppercase text-blue-400">
                      27+
                    </span>
                  </div>

                  <div
                    onClick={() => setStatusFiltroCardTaco('SEM DATA')}
                    className="flex-1 flex flex-col items-center h-full justify-end cursor-pointer group border-l border-gray-800 pl-1"
                  >
                    <span className="text-[10px] font-black text-gray-400 mb-1 group-hover:scale-110 transition-all">
                      {tacosSemDataFilaGeral.length > 0
                        ? tacosSemDataFilaGeral.length
                        : ''}
                    </span>
                    <div
                      className="w-full max-w-[36px] rounded-t-lg bg-gradient-to-t from-gray-700 via-gray-600 to-gray-400 transition-all group-hover:brightness-125"
                      style={{
                        height: `${Math.max(
                          (tacosSemDataFilaGeral.length / maxQtdMesesTaco) *
                            100,
                          4
                        )}%`,
                      }}
                    ></div>
                    <span className="mt-3 text-[10px] font-bold uppercase text-gray-400">
                      s/m
                    </span>
                  </div>
                </div>
              </div>
            </div>

            <div className="p-4 rounded-xl border border-gray-800 bg-[#111A33] flex items-center gap-3">
              <Search size={18} className="text-gray-400" />
              <input
                type="text"
                value={termoBuscaGerencialTaco}
                onChange={(e) => setTermoBuscaGerencialTaco(e.target.value)}
                placeholder="Filtrar por Placa ou CDV"
                className="bg-transparent w-full outline-none text-xs text-white"
              />
            </div>

            <div className="rounded-2xl border border-gray-800 bg-[#1C2541] overflow-hidden shadow-sm">
              <div className="p-4 border-b border-gray-800 flex justify-between items-center bg-gray-500/5">
                <div className="font-bold text-xs uppercase text-gray-400">
                  Lista:{' '}
                  <span className="text-blue-400 font-black">
                    {statusFiltroCardTaco}
                  </span>{' '}
                  ({listagemFiltradaTacosChefe.length} Resultados)
                </div>
              </div>
              <div className="max-h-[500px] overflow-y-auto">
                <table className="w-full text-left text-xs table-auto">
                  <thead className="sticky top-0 bg-[#111A33] font-bold text-gray-400 uppercase text-[10px] border-b border-gray-800">
                    <tr>
                      <th className="p-4">Placa</th>
                      <th className="p-4">Modelo</th>
                      <th className="p-4">CDV</th>
                      <th className="p-4 text-blue-400">Filial</th>
                      <th className="p-4">Mês / Ano Meta</th>
                      <th className="p-4">Validade</th>
                      <th className="p-4">Status</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-800">
                    {listagemFiltradaTacosChefe.map((t) => (
                      <tr
                        key={`taco-ger-${t.id}`}
                        className="hover:bg-gray-800/10"
                      >
                        <td className="p-4 font-mono font-bold text-sm">
                          {t.placa}
                        </td>
                        <td className="p-4 text-gray-400 uppercase">
                          {t.modelo || 'CAMINHÃO'}
                        </td>
                        <td className="p-4 uppercase">{t.cdv}</td>
                        <td className="p-4 uppercase text-blue-400 font-semibold">
                          {t.regional}
                        </td>
                        <td className="p-4 text-blue-400 font-black uppercase">
                          {t.mes_referencia || 'SEM MÊS'}
                        </td>
                        <td className="p-4 font-bold">
                          {t.validade
                            ? new Date(
                                t.validade + 'T00:00:00'
                              ).toLocaleDateString('pt-BR')
                            : '--'}
                        </td>
                        <td className="p-4">
                          <span
                            className={`text-[10px] font-black px-2 py-0.5 rounded ${
                              t.status === 'Concluído' || t.status === 'EM DIA'
                                ? 'bg-green-500/10 text-green-500'
                                : t.status === 'EM EXECUÇÃO' ||
                                  t.status === 'OFICINA'
                                ? 'bg-orange-500/10 text-orange-500'
                                : t.status === 'A VENCER'
                                ? 'bg-yellow-500/10 text-yellow-600'
                                : 'bg-red-50/10 text-red-500'
                            }`}
                          >
                            {t.status === 'Concluído' || t.status === 'EM DIA'
                              ? 'EM DIA'
                              : t.status}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* ============================================================================ */}
        {/* --- TELA 4: NOVA ABA STATUS FINANCEIRO (DIVISÃO RATEIO + CTA + BASE) --- */}
        {/* ============================================================================ */}
        {activeTab === 'status_financeiro' && (
          <div className="max-w-6xl mx-auto animate-fadeIn space-y-6">
            {/* CABEÇALHO FINANCEIRO */}
            <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 border-b border-gray-800 pb-4">
              <div>
                <h2 className="text-2xl font-bold tracking-tight text-white flex items-center gap-2">
                  <DollarSign className="text-rose-400" size={26} />
                  <span>Status Financeiro & Divisão de Rateio</span>
                </h2>
                <p className="text-gray-400 text-xs mt-0.5">
                  Leitura direta da aba "DIVISÃO RATEIO", alertas críticos "CTA", Orçamento por CDV e auditoria "BASE (SUL II)".
                </p>
              </div>
              <div className="flex items-center gap-2 p-2 rounded-xl border border-gray-800 bg-[#1C2541] text-xs font-bold text-white">
                <Clock size={14} className="text-rose-400" />
                <span className="text-gray-400 text-[11px]">Mês Foco:</span>
                <select
                  value={mesSelecionadoFinanceiro}
                  onChange={(e) => {
                    setMesSelecionadoFinanceiro(Number(e.target.value));
                    setContaFiltroFinanceiro('TODAS');
                    setNfFiltroFinanceiro('TODAS');
                    setCdvFiltroFinanceiro('TODOS');
                  }}
                  className="bg-transparent outline-none cursor-pointer font-black text-rose-400 uppercase"
                >
                  {mesesDoAno.map((m) => (
                    <option
                      key={m.num}
                      value={m.num}
                      className="bg-[#1C2541] text-white"
                    >
                      {m.nome}{' '}
                      {mesOficialRateio === m.num ? '(Rateio Oficial)' : ''}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            {/* CAIXA DE IMPORTAÇÃO DA PLANILHA OS's - SUL2 */}
            <div className="p-5 bg-[#1C2541] border border-gray-800 rounded-2xl flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
              <div className="flex-1">
                <h3 className="text-sm font-black uppercase text-rose-400">
                  💸 Importar Planilha Financeira (OS's - SUL2 - 2026)
                </h3>
                <p className="text-[11px] opacity-70 mt-1">
                  Carrega simultaneamente as abas "DIVISÃO RATEIO", "CTA", "Posição Nacional" e "BASE (SUL II)".
                </p>
                <label className="flex items-center gap-2 mt-2 cursor-pointer text-xs text-gray-400">
                  <input
                    type="checkbox"
                    checked={substituirFinanceiro}
                    onChange={(e) => setSubstituirFinanceiro(e.target.checked)}
                    className="accent-rose-500"
                  />
                  Substituir dados financeiros atuais
                </label>
              </div>
              <div className="flex flex-col sm:flex-row gap-2 shrink-0">
                <button
                  onClick={limparTodaTabelaFinanceiro}
                  className="bg-red-600/10 text-red-500 hover:bg-red-600 hover:text-white border border-red-500/20 text-xs font-black px-4 py-2.5 rounded-xl transition-all flex items-center gap-1.5"
                >
                  <Trash2 size={14} /> Limpar Banco
                </button>
                <label className="bg-gradient-to-r from-rose-600 to-orange-500 hover:from-rose-500 hover:to-orange-400 text-white font-black text-xs px-4 py-2.5 rounded-xl cursor-pointer flex items-center gap-1.5 transition-all shadow-lg shadow-rose-600/20">
                  <Upload size={14} />
                  Carregar Excel Financeiro
                  <input
                    type="file"
                    accept=".xlsx, .xls"
                    onChange={lidarComImportacaoFinanceiro}
                    className="hidden"
                  />
                </label>
              </div>
            </div>

            {/* BANNER DE ALERTA VERMELHO: ABA CTA (CALL TO ACTION) */}
            {finCta.length > 0 && (
              <div className="p-5 rounded-2xl border-2 border-red-500/60 bg-gradient-to-r from-red-950/80 via-[#1C2541] to-[#1C2541] shadow-xl relative overflow-hidden">
                <div className="flex items-center justify-between mb-3">
                  <div className="flex items-center gap-2">
                    <span className="p-2 rounded-xl bg-red-600 text-white animate-pulse">
                      <AlertTriangle size={18} />
                    </span>
                    <div>
                      <h3 className="text-sm font-black uppercase text-red-400 tracking-wide">
                        🚨 Call To Action (CTA) — Veículos Críticos Aguardando Decisão
                      </h3>
                      <p className="text-[11px] text-gray-300">
                        Ordens de serviço de alto impacto sinalizadas na aba CTA da planilha
                      </p>
                    </div>
                  </div>
                  <span className="text-xs font-black bg-red-600/20 text-red-300 border border-red-500/30 px-3 py-1 rounded-full">
                    Total Travado:{' '}
                    {formatarMoeda(
                      finCta.reduce((acc, c) => acc + (Number(c.valor) || 0), 0)
                    )}
                  </span>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mt-3">
                  {finCta.map((item, idx) => (
                    <div
                      key={item.id || idx}
                      onClick={() => setTermoBuscaFinanceiro(item.placa)}
                      className="p-3.5 rounded-xl bg-[#0B1329]/90 border border-red-500/30 hover:border-red-400 cursor-pointer transition-all flex justify-between items-center gap-3"
                    >
                      <div className="space-y-1">
                        <div className="flex items-center gap-2">
                          <span className="font-mono font-black text-sm text-white bg-red-600/30 px-2 py-0.5 rounded border border-red-500/40">
                            {item.placa}
                          </span>
                          <span className="text-xs font-black text-amber-400 uppercase">
                            {item.servico}
                          </span>
                          <span className="text-[10px] font-bold text-gray-400">
                            OS: {item.os}
                          </span>
                        </div>
                        <p className="text-[11px] text-gray-300 font-semibold">
                          🏢 {item.oficina}
                        </p>
                        <p className="text-[11px] text-red-300 italic">
                          ⚠️ Impacto: {item.acao}
                        </p>
                      </div>
                      <div className="text-right shrink-0">
                        <span className="text-[10px] font-black px-2 py-0.5 rounded bg-red-500/20 text-red-400 uppercase block mb-1">
                          {item.status}
                        </span>
                        <span className="text-base font-black text-white">
                          {formatarMoeda(item.valor)}
                        </span>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* 4 CARDS EXECUTIVOS DO MÊS (ORÇAMENTO AOP, FATURADO NFs, PENDENTE E SALDO) */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              <div
                onClick={() => {
                  setContaFiltroFinanceiro('TODAS');
                  setNfFiltroFinanceiro('TODAS');
                  setCdvFiltroFinanceiro('TODOS');
                }}
                className="p-4 rounded-xl shadow-sm border border-gray-800 cursor-pointer bg-gradient-to-br from-blue-600 to-indigo-700 text-white"
              >
                <div className="flex justify-between items-start">
                  <span className="text-[11px] font-black uppercase">
                    ORÇAMENTO (AOP + PROV)
                  </span>
                  <span className="text-[9px] bg-white/20 px-1.5 py-0.5 rounded font-bold">
                    META MÊS {mesSelecionadoFinanceiro}
                  </span>
                </div>
                <h3 className="text-2xl font-black mt-3">
                  {formatarMoeda(totalOrcamentoMes)}
                </h3>
                <p className="text-[10px] opacity-80 mt-1">
                  Teto liberado para SUL2
                </p>
              </div>

              <div
                onClick={() =>
                  setNfFiltroFinanceiro(
                    nfFiltroFinanceiro === 'COBRADAS' ? 'TODAS' : 'COBRADAS'
                  )
                }
                className={`p-4 rounded-xl shadow-sm border border-gray-800 cursor-pointer transition-all ${
                  nfFiltroFinanceiro === 'COBRADAS'
                    ? 'ring-4 ring-emerald-400'
                    : ''
                } bg-gradient-to-br from-emerald-600 to-teal-700 text-white`}
              >
                <div className="flex justify-between items-start">
                  <span className="text-[11px] font-black uppercase">
                    FATURADO (NFs 1/4 A 4/4)
                  </span>
                  <span className="text-[10px] bg-white/20 px-1.5 py-0.5 rounded font-bold">
                    {totalOrcamentoMes > 0
                      ? (
                          (totalFaturadoNFsMes / totalOrcamentoMes) *
                          100
                        ).toFixed(1)
                      : '0.0'}
                    %
                  </span>
                </div>
                <h3 className="text-2xl font-black mt-3">
                  {formatarMoeda(totalFaturadoNFsMes)}
                </h3>
                <p className="text-[10px] opacity-80 mt-1">
                  Notas Fiscais já cobradas
                </p>
              </div>

              <div
                onClick={() =>
                  setNfFiltroFinanceiro(
                    nfFiltroFinanceiro === 'PEND' ? 'TODAS' : 'PEND'
                  )
                }
                className={`p-4 rounded-xl shadow-sm border border-gray-800 cursor-pointer transition-all ${
                  nfFiltroFinanceiro === 'PEND' ? 'ring-4 ring-amber-400' : ''
                } bg-gradient-to-br from-orange-500 to-amber-600 text-white`}
              >
                <div className="flex justify-between items-start">
                  <span className="text-[11px] font-black uppercase">
                    PENDENTE (PEND)
                  </span>
                  <span className="text-[10px] bg-white/20 px-1.5 py-0.5 rounded font-bold">
                    {totalOrcamentoMes > 0
                      ? ((totalPendenteMes / totalOrcamentoMes) * 100).toFixed(
                          1
                        )
                      : '0.0'}
                    %
                  </span>
                </div>
                <h3 className="text-2xl font-black mt-3">
                  {formatarMoeda(totalPendenteMes)}
                </h3>
                <p className="text-[10px] opacity-80 mt-1">
                  Concluídas não cobradas / Em aberto
                </p>
              </div>

              <div
                className={`p-4 rounded-xl shadow-sm border border-gray-800 ${
                  saldoFinalMes >= 0
                    ? 'bg-gradient-to-br from-green-600 to-emerald-700'
                    : 'bg-gradient-to-br from-red-600 to-rose-700'
                } text-white`}
              >
                <div className="flex justify-between items-start">
                  <span className="text-[11px] font-black uppercase">
                    SALDO FINAL (RATEIO)
                  </span>
                  <span className="text-[9px] bg-black/20 px-1.5 py-0.5 rounded font-black uppercase">
                    {saldoFinalMes >= 0 ? 'DENTRO DO AOP' : 'ESTOURO DE META'}
                  </span>
                </div>
                <h3 className="text-2xl font-black mt-3">
                  {formatarMoeda(saldoFinalMes)}
                </h3>
                <p className="text-[10px] opacity-80 mt-1">
                  (AOP + PROV) - (PEND + NFs)
                </p>
              </div>
            </div>

            {/* ============================================================================ */}
            {/* --- TABELA MATRIZ OFICIAL: ABA DIVISÃO RATEIO --- */}
            {/* ============================================================================ */}
            <div className="rounded-2xl border border-gray-800 bg-[#1C2541] overflow-hidden shadow-xl">
              <div className="p-4 border-b border-gray-800 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-2 bg-[#111A33]">
                <div>
                  <h3 className="text-sm font-black uppercase text-white flex items-center gap-2">
                    <span>
                      📋 Matriz Divisão de Rateio (Mês {mesSelecionadoFinanceiro})
                    </span>
                    {usandoRateioDireto && (
                      <span className="text-[10px] bg-emerald-500/20 text-emerald-400 border border-emerald-500/30 px-2 py-0.5 rounded-full">
                        Lido Diretamente da Aba DIVISÃO RATEIO
                      </span>
                    )}
                  </h3>
                  <p className="text-[11px] text-gray-400">
                    Clique em qualquer conta abaixo (ex: MANUTENÇÃO, PNEUS, SINISTRO) para filtrar as OSs correspondentes na Base.
                  </p>
                </div>
                {contaFiltroFinanceiro !== 'TODAS' && (
                  <button
                    onClick={() => setContaFiltroFinanceiro('TODAS')}
                    className="text-[10px] font-black bg-rose-500/20 text-rose-300 border border-rose-500/30 px-3 py-1 rounded-lg hover:bg-rose-500 hover:text-white transition-all"
                  >
                    Limpar Filtro: {contaFiltroFinanceiro} ✕
                  </button>
                )}
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs table-auto">
                  <thead className="bg-[#0B1329] font-black text-gray-400 uppercase text-[10px] border-b border-gray-800">
                    <tr>
                      <th className="p-3.5">Conta Correção</th>
                      <th className="p-3.5 text-right text-blue-400">
                        AOP (Meta)
                      </th>
                      <th className="p-3.5 text-right text-indigo-400">PROV</th>
                      <th className="p-3.5 text-right text-orange-400">PEND</th>
                      <th className="p-3.5 text-right text-emerald-400">
                        {labelNf1}
                      </th>
                      <th className="p-3.5 text-right text-emerald-400">
                        {labelNf2}
                      </th>
                      <th className="p-3.5 text-right text-emerald-400">
                        {labelNf3}
                      </th>
                      <th className="p-3.5 text-right text-emerald-400">
                        {labelNf4}
                      </th>
                      <th className="p-3.5 text-right text-white">SALDO</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-800">
                    {linhasMatrizRateio.map((linha, index) => {
                      const isTotal = String(linha.conta)
                        .toUpperCase()
                        .includes('TOTAL GERAL');
                      const selecionada =
                        contaFiltroFinanceiro ===
                        String(linha.conta).toUpperCase();

                      return (
                        <tr
                          key={index}
                          onClick={() => {
                            if (!isTotal) {
                              const c = String(linha.conta).toUpperCase();
                              setContaFiltroFinanceiro(
                                contaFiltroFinanceiro === c ? 'TODAS' : c
                              );
                            }
                          }}
                          className={`transition-all ${
                            isTotal
                              ? 'bg-[#111A33] font-black text-white border-t-2 border-gray-700'
                              : selecionada
                              ? 'bg-rose-500/20 border-l-4 border-rose-500 cursor-pointer'
                              : 'hover:bg-gray-800/30 cursor-pointer'
                          }`}
                        >
                          <td className="p-3.5 font-black uppercase text-white">
                            {linha.conta}
                          </td>
                          <td className="p-3.5 text-right font-mono text-blue-300">
                            {formatarMoeda(linha.aop)}
                          </td>
                          <td className="p-3.5 text-right font-mono text-gray-400">
                            {linha.prov ? formatarMoeda(linha.prov) : '--'}
                          </td>
                          <td className="p-3.5 text-right font-mono text-orange-400 font-bold">
                            {linha.pend ? formatarMoeda(linha.pend) : '--'}
                          </td>
                          <td className="p-3.5 text-right font-mono text-gray-200">
                            {linha.nf1 ? formatarMoeda(linha.nf1) : '--'}
                          </td>
                          <td className="p-3.5 text-right font-mono text-gray-200">
                            {linha.nf2 ? formatarMoeda(linha.nf2) : '--'}
                          </td>
                          <td className="p-3.5 text-right font-mono text-gray-200">
                            {linha.nf3 ? formatarMoeda(linha.nf3) : '--'}
                          </td>
                          <td className="p-3.5 text-right font-mono text-gray-200">
                            {linha.nf4 ? formatarMoeda(linha.nf4) : '--'}
                          </td>
                          <td
                            className={`p-3.5 text-right font-mono font-black ${
                              Number(linha.saldo) >= 0
                                ? 'text-green-400'
                                : 'text-red-400'
                            }`}
                          >
                            {formatarMoeda(linha.saldo)}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>

            {/* ============================================================================ */}
            {/* --- NOVA SEÇÃO: ORÇAMENTO PROPORCIONAL & RANKING DE GASTO POR CDV --- */}
            {/* ============================================================================ */}
            <div className="rounded-2xl border border-gray-800 bg-[#1C2541] overflow-hidden shadow-xl">
              <div className="p-4 border-b border-gray-800 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-2 bg-[#111A33]">
                <div>
                  <h3 className="text-sm font-black uppercase text-white flex items-center gap-2">
                    <span>
                      🏢 Desempenho Orçamentário & Ranking por CDV ({totalPlacasSul2} Placas SUL2)
                    </span>
                  </h3>
                  <p className="text-[11px] text-gray-400">
                    Comparativo entre o Orçamento Proporcional por frota (Posição Nacional) vs. Gasto Real (BASE SUL II). Clique em um CDV para auditar suas OSs.
                  </p>
                </div>
                {cdvFiltroFinanceiro !== 'TODOS' && (
                  <button
                    onClick={() => setCdvFiltroFinanceiro('TODOS')}
                    className="text-[10px] font-black bg-blue-500/20 text-blue-300 border border-blue-500/30 px-3 py-1 rounded-lg hover:bg-blue-500 hover:text-white transition-all"
                  >
                    Limpar Filtro CDV: {cdvFiltroFinanceiro} ✕
                  </button>
                )}
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs table-auto">
                  <thead className="bg-[#0B1329] font-black text-gray-400 uppercase text-[10px] border-b border-gray-800">
                    <tr>
                      <th className="p-3.5">Ranking / CDV</th>
                      <th className="p-3.5 text-center">Placas</th>
                      <th className="p-3.5 text-right">% Frota</th>
                      <th className="p-3.5 text-right text-blue-400">
                        Orçamento (Teto)
                      </th>
                      <th className="p-3.5 text-right text-amber-400">
                        Gasto Real (Mês)
                      </th>
                      <th className="p-3.5 text-center">% do Orçamento</th>
                      <th className="p-3.5 text-right">
                        Estouro / Economia (%)
                      </th>
                      <th className="p-3.5 w-44">
                        Peso no Gasto Geral (Quem gasta +)
                      </th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-800">
                    {analisePorCdv.map((item, idx) => {
                      const estourou = item.pctDoOrcamento > 100;
                      const semOrcamento =
                        item.orcamentoCdv === 0 && item.gastoReal > 0;
                      const selecionado = cdvFiltroFinanceiro === item.cdv;

                      return (
                        <tr
                          key={item.cdv}
                          onClick={() =>
                            setCdvFiltroFinanceiro(
                              selecionado ? 'TODOS' : item.cdv
                            )
                          }
                          className={`transition-all cursor-pointer ${
                            selecionado
                              ? 'bg-blue-500/20 border-l-4 border-blue-400'
                              : 'hover:bg-gray-800/30'
                          }`}
                        >
                          <td className="p-3.5 font-black text-white flex items-center gap-2">
                            <span className="text-[10px] font-mono bg-[#0B1329] text-gray-400 px-1.5 py-0.5 rounded border border-gray-700">
                              #{idx + 1}
                            </span>
                            <span>{item.cdv}</span>
                          </td>
                          <td className="p-3.5 text-center font-mono font-bold text-gray-300">
                            {item.placas > 0 ? item.placas : '--'}
                          </td>
                          <td className="p-3.5 text-right font-mono text-gray-400">
                            {item.placas > 0
                              ? `${item.pctFrota.toFixed(2)}%`
                              : '--'}
                          </td>
                          <td className="p-3.5 text-right font-mono text-blue-300 font-semibold">
                            {item.orcamentoCdv > 0
                              ? formatarMoeda(item.orcamentoCdv)
                              : '--'}
                          </td>
                          <td className="p-3.5 text-right font-mono font-black text-white">
                            {formatarMoeda(item.gastoReal)}
                          </td>
                          <td className="p-3.5 text-center">
                            {item.orcamentoCdv > 0 ? (
                              <span
                                className={`text-[11px] font-mono font-black px-2 py-0.5 rounded ${
                                  estourou
                                    ? 'bg-red-500/20 text-red-400 border border-red-500/30'
                                    : 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/30'
                                }`}
                              >
                                {item.pctDoOrcamento.toFixed(2)}%
                              </span>
                            ) : (
                              <span className="text-[10px] text-gray-500">
                                Extra SUL2
                              </span>
                            )}
                          </td>
                          <td className="p-3.5 text-right font-mono font-black">
                            {item.orcamentoCdv > 0 ? (
                              estourou ? (
                                <span className="text-red-400">
                                  +{item.pctEstouro.toFixed(2)}% acima 🔴
                                </span>
                              ) : (
                                <span className="text-emerald-400">
                                  {item.pctEstouro.toFixed(2)}% (Economia) 🟢
                                </span>
                              )
                            ) : semOrcamento ? (
                              <span className="text-amber-400 text-[10px]">
                                Sem Teto Local
                              </span>
                            ) : (
                              '--'
                            )}
                          </td>
                          <td className="p-3.5">
                            <div className="flex items-center gap-2">
                              <div className="flex-1 h-2 bg-[#0B1329] rounded-full overflow-hidden border border-gray-800">
                                <div
                                  className={`h-full rounded-full ${
                                    estourou
                                      ? 'bg-gradient-to-r from-orange-500 to-red-500'
                                      : 'bg-gradient-to-r from-blue-500 to-emerald-400'
                                  }`}
                                  style={{
                                    width: `${Math.min(
                                      item.shareGastoTotal,
                                      100
                                    )}%`,
                                  }}
                                ></div>
                              </div>
                              <span className="font-mono text-[10px] font-black text-amber-400 w-11 text-right">
                                {item.shareGastoTotal.toFixed(1)}%
                              </span>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>

            {/* GRÁFICO ANUAL DE DESPESAS POR MÊS (ESTILO PREMIUM) */}
            <div className="p-6 rounded-2xl border border-gray-800/80 bg-[#131B35] shadow-2xl relative overflow-hidden">
              <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-6 gap-2">
                <div>
                  <h3 className="text-lg font-black text-white tracking-tight flex items-center gap-2">
                    <span>📊 Evolução Mensal de OSs (BASE SUL II)</span>
                  </h3>
                  <p className="text-xs text-gray-400">
                    Volume financeiro de OSs por mês em 2026 (Clique em uma barra para auditar o mês)
                  </p>
                </div>
                <div className="flex items-center gap-2 bg-[#0B1329] p-1.5 rounded-xl border border-gray-800">
                  <span className="text-[10px] font-bold text-gray-400 px-2 uppercase">
                    Ano 2026
                  </span>
                </div>
              </div>

              <div className="relative pt-6 pb-2">
                <div className="absolute inset-0 flex flex-col justify-between pointer-events-none opacity-20 border-b border-gray-700">
                  <div className="border-b border-gray-600 w-full"></div>
                  <div className="border-b border-gray-600 w-full"></div>
                  <div className="border-b border-gray-600 w-full"></div>
                  <div className="border-b border-gray-600 w-full"></div>
                </div>

                <div className="h-48 flex items-end justify-between gap-1.5 sm:gap-2 px-2 relative z-10">
                  {gastosPorMesAnual.map((item) => {
                    const estaSelecionado =
                      Number(mesSelecionadoFinanceiro) === item.num;
                    const alturaPct = Math.max(
                      (item.valorGasto / maxValorMesFinanceiro) * 100,
                      4
                    );

                    return (
                      <div
                        key={item.num}
                        onClick={() => setMesSelecionadoFinanceiro(item.num)}
                        className="flex-1 flex flex-col items-center h-full justify-end cursor-pointer group"
                      >
                        <span
                          className={`text-[9px] font-black mb-1 transition-all ${
                            estaSelecionado
                              ? 'text-rose-400 scale-110'
                              : 'text-gray-400 group-hover:text-white'
                          }`}
                        >
                          {item.valorGasto > 0
                            ? `${(item.valorGasto / 1000).toFixed(0)}k`
                            : ''}
                        </span>

                        <div
                          className={`w-full max-w-[36px] rounded-t-lg transition-all duration-300 ${
                            estaSelecionado
                              ? 'bg-gradient-to-t from-rose-600 via-orange-500 to-amber-300 ring-2 ring-rose-400 shadow-lg shadow-rose-500/30'
                              : 'bg-gradient-to-t from-rose-900/60 via-orange-600/70 to-amber-500/80 group-hover:from-rose-600 group-hover:to-amber-400'
                          }`}
                          style={{ height: `${alturaPct}%` }}
                        ></div>

                        <span
                          className={`mt-3 text-[11px] font-bold uppercase transition-all ${
                            estaSelecionado
                              ? 'text-rose-400 font-black'
                              : 'text-gray-400 group-hover:text-white'
                          }`}
                        >
                          {item.sigla}
                        </span>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>

            {/* BARRA DE PESQUISA E AUDITORIA DE ORDENS DE SERVIÇO (BASE SUL II) */}
            <div className="p-4 rounded-xl border border-gray-800 bg-[#111A33] flex items-center gap-3">
              <Search size={18} className="text-gray-400" />
              <input
                type="text"
                value={termoBuscaFinanceiro}
                onChange={(e) => setTermoBuscaFinanceiro(e.target.value)}
                placeholder="Pesquisar OS por Placa, Número da OS, CDV, Oficina ou NF..."
                className="bg-transparent w-full outline-none text-xs text-white"
              />
              {termoBuscaFinanceiro && (
                <button
                  onClick={() => setTermoBuscaFinanceiro('')}
                  className="text-xs text-gray-400 hover:text-white font-bold"
                >
                  Limpar
                </button>
              )}
            </div>

            {/* TABELA DETALHADA DE ORDENS DE SERVIÇO (ABA BASE SUL II) */}
            <div className="rounded-2xl border border-gray-800 bg-[#1C2541] overflow-hidden shadow-sm">
              <div className="p-4 border-b border-gray-800 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-2 bg-gray-500/5">
                <div className="font-bold text-xs uppercase text-gray-400">
                  Auditoria BASE (SUL II) — Conta:{' '}
                  <span className="text-rose-400 font-black">
                    {contaFiltroFinanceiro}
                  </span>{' '}
                  | CDV:{' '}
                  <span className="text-blue-400 font-black">
                    {cdvFiltroFinanceiro}
                  </span>{' '}
                  | Filtro NF:{' '}
                  <span className="text-amber-400 font-black">
                    {nfFiltroFinanceiro}
                  </span>{' '}
                  ({listagemOsFinanceiroFiltrada.length} OSs)
                </div>
                <div className="text-xs font-black text-emerald-400">
                  Soma Filtrada:{' '}
                  {formatarMoeda(
                    listagemOsFinanceiroFiltrada.reduce(
                      (acc, b) => acc + (Number(b.valor_total) || 0),
                      0
                    )
                  )}
                </div>
              </div>
              <div className="max-h-[500px] overflow-y-auto">
                <table className="w-full text-left text-xs table-auto">
                  <thead className="sticky top-0 bg-[#111A33] font-bold text-gray-400 uppercase text-[10px] border-b border-gray-800">
                    <tr>
                      <th className="p-3.5">OS</th>
                      <th className="p-3.5">Data</th>
                      <th className="p-3.5">Placa</th>
                      <th className="p-3.5">CDV</th>
                      <th className="p-3.5">Estabelecimento / Oficina</th>
                      <th className="p-3.5">Conta</th>
                      <th className="p-3.5">Status</th>
                      <th className="p-3.5">NF / Lote</th>
                      <th className="p-3.5 text-right">Valor Total</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-800">
                    {listagemOsFinanceiroFiltrada.slice(0, 300).map((b) => (
                      <tr key={b.id} className="hover:bg-gray-800/20">
                        <td className="p-3.5 font-mono text-gray-300 font-bold">
                          {b.os}
                        </td>
                        <td className="p-3.5 text-gray-400">
                          {b.data_os
                            ? new Date(
                                b.data_os + 'T00:00:00'
                              ).toLocaleDateString('pt-BR')
                            : '--'}
                        </td>
                        <td className="p-3.5 font-mono font-black text-white">
                          {b.placa}
                        </td>
                        <td className="p-3.5 uppercase text-blue-400 font-semibold">
                          {b.cdv}
                        </td>
                        <td className="p-3.5 text-gray-300 truncate max-w-[200px]">
                          {b.estabelecimento}
                        </td>
                        <td className="p-3.5">
                          <span className="text-[10px] font-black px-2 py-0.5 rounded bg-gray-800 text-amber-400 uppercase">
                            {b.conta_correcao}
                          </span>
                        </td>
                        <td className="p-3.5 text-gray-300 text-[11px]">
                          {b.status}
                        </td>
                        <td className="p-3.5">
                          <span
                            className={`text-[10px] font-black px-2 py-0.5 rounded ${
                              b.nf === 'PEND'
                                ? 'bg-orange-500/15 text-orange-400'
                                : 'bg-emerald-500/15 text-emerald-400'
                            }`}
                          >
                            {b.nf}
                          </span>
                        </td>
                        <td className="p-3.5 text-right font-mono font-black text-white">
                          {formatarMoeda(b.valor_total)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}

export default App; 
