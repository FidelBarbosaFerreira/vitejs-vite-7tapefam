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
} from 'lucide-react';

const LINK_POWER_BI = 'https://app.powerbi.com';

// ============================================================================
// --- FUNÇÃO AUXILIAR DE TRATAMENTO DE DATAS INTERNAS DO EXCEL ---
// ============================================================================
function formatarDataVersaoBancoNoTopo(raw, dataPadrao) {
  if (!raw || String(raw).trim() === '') return null;
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
      const ano = partes[2];
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

  // --- SELETORES DE DATA ---
  const [mesSelecionadoHistorico, setMesSelecionadoHistorico] = useState(
    new Date().getMonth() + 1
  );
  const [mesSelecionadoTacoGerencial, setMesSelecionadoTacoGerencial] =
    useState(new Date().getMonth() + 1);

  // --- INTERAÇÃO DE FILTRO DOS CARDS ---
  const [statusFiltroCard, setStatusFiltroCard] = useState('TODOS');
  const [statusFiltroCardTaco, setStatusFiltroCardTaco] = useState('TODOS');

  // --- INTERRUPTORES DE SUBSTITUIÇÃO ---
  const [substituirFrota, setSubstituirFrota] = useState(false);
  const [substituirTaco, setSubstituirTaco] = useState(false);

  // --- DADOS REAIS DO BANCO ---
  const [veiculos, setVeiculos] = useState([]);
  const [preventivas, setPreventivas] = useState([]);
  const [tacografos, setTacografos] = useState([]);

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

  // Mapeamento dos 12 meses para o gráfico estilo "Meta vs Receita" de Preventivas
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
  // TRATAMENTO UNIFICADO: VENCIDA / VENCIDAS / VENCIDO
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
  // TRATAMENTO UNIFICADO TACOGRAFO: VENCIDA / VENCIDAS / VENCIDO
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
                      ? 'bg-blue-50/10 text-[#0000FF]'
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
      </main>
    </div>
  );
}

export default App;
