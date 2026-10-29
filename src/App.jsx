import { useState, useEffect, useRef } from 'react';
import { supabase } from './supabase';
import * as XLSX from 'xlsx';
import {
  BarChart3,
  CalendarDays,
  Trash2,
  RefreshCw,
  Clock,
  ExternalLink,
  Upload,
  Truck,
  Search,
  History,
  DollarSign,
  AlertTriangle,
  Building2,
  Box,
  Layers,
  ChevronLeft,
  ChevronRight,
} from 'lucide-react';

const LINK_POWER_BI = 'https://app.powerbi.com';

// ============================================================================
// --- MAPEADOR DAS 7 ZONAS ANATÔMICAS DO CAMINHÃO VW 8.120 PEPSICO ---
// ============================================================================
function identificarZonaDefeito3D(statusTexto) {
  const s = String(statusTexto || '').toUpperCase();
  if (
    s.includes('MOTOR') ||
    s.includes('CABEÇOTE') ||
    s.includes('CABECOTE') ||
    s.includes('RADIADOR') ||
    s.includes('AQUECENDO') ||
    s.includes('TURBINA') ||
    s.includes('BOMBA')
  ) {
    return {
      id: 'MOTOR',
      nome: 'MOTOR, TURBINA & ARREFECIMENTO',
      cor: '#ef4444',
      corHex: 0xef4444,
      simbolo: '⊗',
      descricao:
        'Bloco do motor diesel 8.120, cabeçote, turbocompressor, radiador e intercooler sob a cabine basculante.',
      bascularCabine: true,
      abrirBau: false,
      raioXRecomendado: false,
      cameraPreset: 'MOTOR',
    };
  }
  if (
    s.includes('CÂMBIO') ||
    s.includes('CAMBIO') ||
    s.includes('EMBREAGEM') ||
    s.includes('TRANSMISSÃO') ||
    s.includes('TRANSMISSAO')
  ) {
    return {
      id: 'TRANSMISSAO',
      nome: 'TRANSMISSÃO, CÂMBIO & EMBREAGEM',
      cor: '#f97316',
      corHex: 0xf97316,
      simbolo: '⊗',
      descricao:
        'Caixa seletora de marchas, capa seca, platô/disco de embreagem e acoplamento do eixo cardan.',
      bascularCabine: false,
      abrirBau: false,
      raioXRecomendado: true,
      cameraPreset: 'CAMBIO',
    };
  }
  if (
    s.includes('IGNIÇÃO') ||
    s.includes('IGNICAO') ||
    s.includes('PAINEL') ||
    s.includes('ELÉTRICA') ||
    s.includes('ELETRICA') ||
    s.includes('PARABRISA') ||
    s.includes('TACÓGRAFO') ||
    s.includes('TACOGRAFO') ||
    s.includes('PARTIDA') ||
    s.includes('ALTERNADOR') ||
    s.includes('BATERIA')
  ) {
    return {
      id: 'PARTIDA',
      nome: 'PARTIDA, ELÉTRICA & CABINE',
      cor: '#eab308',
      corHex: 0xeab308,
      simbolo: '!',
      descricao:
        'Sistema de ignição/partida, chicote elétrico, painel de instrumentos, tacógrafo e estrutura da cabine.',
      bascularCabine: false,
      abrirBau: false,
      raioXRecomendado: false,
      cameraPreset: 'FRENTE',
    };
  }
  if (
    s.includes('RODA') ||
    s.includes('PNEU') ||
    s.includes('BORRACHARIA') ||
    s.includes('FREIO') ||
    s.includes('CUBO')
  ) {
    return {
      id: 'PNEU',
      nome: 'PNEUS, CUBOS & RODAGEM',
      cor: '#ec4899',
      corHex: 0xec4899,
      simbolo: '⊗',
      descricao:
        'Rodado traseiro duplo e dianteiro, retentores de roda, tambores de freio, cubos e pneus.',
      bascularCabine: false,
      abrirBau: false,
      raioXRecomendado: true,
      cameraPreset: 'PNEUS',
    };
  }
  if (
    s.includes('SUSPENSÃO') ||
    s.includes('SUSPENSAO') ||
    s.includes('MOLA') ||
    s.includes('CATALISADOR') ||
    s.includes('ESCAPAMENTO') ||
    s.includes('CHASSI') ||
    s.includes('CARDAN') ||
    s.includes('DIFERENCIAL') ||
    s.includes('ARLA')
  ) {
    return {
      id: 'UNDERTRUCK',
      nome: 'UNDERTRUCK (SUSPENSÃO / CARDAN / EXAUSTÃO)',
      cor: '#a855f7',
      corHex: 0xa855f7,
      simbolo: '⊗',
      descricao:
        'Subestrutura Undertruck: feixes de mola parabólicos, amortecedores, cardan, diferencial, tanque e catalisador.',
      bascularCabine: false,
      abrirBau: false,
      raioXRecomendado: true,
      cameraPreset: 'CAMBIO',
    };
  }
  if (
    s.includes('BAU') ||
    s.includes('BAÚ') ||
    s.includes('AGUA') ||
    s.includes('ÁGUA') ||
    s.includes('LAVAGEM') ||
    s.includes('PORTA')
  ) {
    return {
      id: 'BAU',
      nome: 'BAÚ DE CARGA & CARROCERIA PEPSICO',
      cor: '#3b82f6',
      corHex: 0x3b82f6,
      simbolo: '⊗',
      descricao:
        'Carroceria de alumínio, portas traseiras e lateral de entrega, vedação contra infiltração e assoalho.',
      bascularCabine: false,
      abrirBau: true,
      raioXRecomendado: false,
      cameraPreset: 'TRASEIRA',
    };
  }
  if (s.includes('SINISTRO')) {
    return {
      id: 'SINISTRO',
      nome: 'SINISTRO / COLISÃO ESTRUTURAL',
      cor: '#dc2626',
      corHex: 0xdc2626,
      simbolo: '⊗',
      descricao:
        'Avaria estrutural por colisão/acidente exigindo reparo de funilaria, alinhamento de chassi e cabine.',
      bascularCabine: false,
      abrirBau: false,
      raioXRecomendado: false,
      cameraPreset: 'FRENTE',
    };
  }
  return {
    id: 'PREVENTIVA',
    nome: 'PREVENTIVA / DIAGNÓSTICO GERAL',
    cor: '#06b6d4',
    corHex: 0x06b6d4,
    simbolo: '!',
    descricao:
      'Manutenção preventiva programada, check-up multiponto ou aguardando aprovação/liberação.',
    bascularCabine: false,
    abrirBau: false,
    raioXRecomendado: false,
    cameraPreset: 'ISO',
  };
}

// ============================================================================
// --- COMPONENTE 1: PLANTA TÉCNICA 2D COMPLEXA (VW 8.120 PEPSICO) ---
// ============================================================================
function PlantaTecnica2DCaminhao({ veiculo, zonaAtiva, filtroZona3D, setFiltroZona3D }) {
  const estaDisponivel =
    String(veiculo?.disponibilidade || '').toUpperCase().includes('DISPON') ||
    String(veiculo?.status || '').toUpperCase() === 'PRONTO';

  const obterStatusCaixa = (idCaixa) => {
    if (!veiculo || estaDisponivel) {
      return { tipo: 'OK', corBg: '#10b981' };
    }
    if (
      zonaAtiva.id === idCaixa ||
      (zonaAtiva.id === 'SINISTRO' && (idCaixa === 'PARTIDA' || idCaixa === 'BAU'))
    ) {
      return { tipo: 'CRITICO', corBg: '#ef4444' };
    }
    if (idCaixa === 'PREVENTIVA' && zonaAtiva.id !== 'PREVENTIVA') {
      return { tipo: 'ATENCAO', corBg: '#f59e0b' };
    }
    return { tipo: 'OK', corBg: '#10b981' };
  };

  const caixasAnatomicas = [
    { id: 'BAU', label: 'BAU', x: 90, y: 42 },
    { id: 'PREVENTIVA', label: 'PREVENTIVA', x: 375, y: 78 },
    { id: 'PARTIDA', label: 'PARTIDA', x: 585, y: 108 },
    { id: 'TRANSMISSAO', label: 'TRANSMISSÃO', x: 472, y: 196 },
    { id: 'MOTOR', label: 'MOTOR', x: 588, y: 196 },
    { id: 'UNDERTRUCK', label: 'UNDERTRUCK', x: 325, y: 222 },
    { id: 'PNEU', label: 'PNEU', x: 115, y: 222 },
  ];

  return (
    <div className="w-full rounded-xl overflow-hidden border-2 border-gray-600 bg-[#f8fafc] text-gray-900 shadow-inner">
      <div className="bg-[#d1d5db] border-b-2 border-gray-500 py-2 px-4 text-center">
        <h4 className="text-lg font-black tracking-wide text-gray-950 uppercase">
          PLACA : {veiculo?.placa || 'AGUARDANDO PLANILHA'} —{' '}
          <span className="text-sm font-bold text-gray-700">
            VW 8.120 PEPSICO ({veiculo?.status || 'CARREGUE O EXCEL DE DISPONIBILIDADE'})
          </span>
        </h4>
      </div>

      <div className="relative p-3 flex items-center justify-center bg-white">
        <svg viewBox="0 0 720 320" className="w-full h-auto max-h-[340px] select-none">
          <line x1="35" y1="265" x2="685" y2="265" stroke="#94a3b8" strokeWidth="2" />

          {/* 1. BAÚ DE CARGA PEPSICO */}
          <g
            onClick={() => setFiltroZona3D(filtroZona3D === 'BAU' ? 'TODAS' : 'BAU')}
            className="cursor-pointer"
          >
            <rect
              x="42"
              y="32"
              width="465"
              height="162"
              fill={zonaAtiva.id === 'BAU' ? '#fee2e2' : '#ffffff'}
              stroke={zonaAtiva.id === 'BAU' ? '#dc2626' : '#334155'}
              strokeWidth={zonaAtiva.id === 'BAU' ? '3' : '2'}
            />
            <rect x="40" y="28" width="469" height="6" fill="#e2e8f0" stroke="#334155" strokeWidth="1.5" />
            <rect x="40" y="190" width="469" height="7" fill="#cbd5e1" stroke="#334155" strokeWidth="1.5" />
            <line x1="48" y1="34" x2="48" y2="190" stroke="#475569" strokeWidth="1.5" />
            <line x1="501" y1="34" x2="501" y2="190" stroke="#475569" strokeWidth="1.5" />

            {/* Porta Lateral Central */}
            <rect x="262" y="56" width="76" height="130" fill="#f8fafc" stroke="#334155" strokeWidth="2" />
            <rect x="265" y="59" width="70" height="124" fill="none" stroke="#94a3b8" strokeWidth="1" />
            <rect x="326" y="70" width="12" height="5" fill="#cbd5e1" stroke="#334155" strokeWidth="1" />
            <rect x="326" y="168" width="12" height="5" fill="#cbd5e1" stroke="#334155" strokeWidth="1" />
            <rect x="272" y="172" width="22" height="9" fill="#cbd5e1" stroke="#334155" strokeWidth="1.2" />

            {/* Globo e Logo PepsiCo */}
            <g transform="translate(88, 108)">
              <circle cx="16" cy="8" r="14" fill="none" stroke="#1e293b" strokeWidth="1.8" />
              <ellipse cx="16" cy="8" rx="14" ry="5" fill="none" stroke="#1e293b" strokeWidth="1.2" />
              <ellipse cx="16" cy="8" rx="6" ry="14" fill="none" stroke="#1e293b" strokeWidth="1.2" />
              <line x1="2" y1="8" x2="30" y2="8" stroke="#1e293b" strokeWidth="1.2" />
              <text
                x="36"
                y="22"
                fill="#1e293b"
                fontSize="21"
                fontWeight="900"
                fontFamily="sans-serif"
                letterSpacing="1"
              >
                PEPSICO
              </text>
            </g>
          </g>

          {/* 2. CABINE VW 8.120 + QUEBRA-VENTO */}
          <g
            onClick={() => setFiltroZona3D(filtroZona3D === 'PARTIDA' ? 'TODAS' : 'PARTIDA')}
            className="cursor-pointer"
          >
            <path
              d="M 512,96 L 512,42 Q 555,52 594,96 Z"
              fill="#f1f5f9"
              stroke="#334155"
              strokeWidth="2"
            />
            <path d="M 515,78 Q 552,82 586,96" fill="none" stroke="#64748b" strokeWidth="1.2" />

            <path
              d="M 512,96 L 605,96 L 642,148 L 648,205 L 512,205 Z"
              fill={
                zonaAtiva.id === 'PARTIDA' || zonaAtiva.id === 'SINISTRO'
                  ? '#fef3c7'
                  : '#ffffff'
              }
              stroke={
                zonaAtiva.id === 'PARTIDA' || zonaAtiva.id === 'SINISTRO'
                  ? '#d97706'
                  : '#334155'
              }
              strokeWidth="2.2"
            />
            <path
              d="M 526,104 L 592,104 L 612,145 L 612,196 L 526,196 Z"
              fill="none"
              stroke="#475569"
              strokeWidth="1.5"
            />
            <polygon
              points="532,108 588,108 604,142 532,142"
              fill="#e2e8f0"
              stroke="#334155"
              strokeWidth="1.5"
            />
            <rect x="586" y="110" width="8" height="26" rx="2" fill="#94a3b8" stroke="#1e293b" strokeWidth="1.5" />
            <rect x="535" y="156" width="14" height="5" rx="1" fill="#475569" />
            <text x="580" y="154" fill="#334155" fontSize="8" fontWeight="bold">
              8.120
            </text>
            <line x1="605" y1="96" x2="640" y2="145" stroke="#1e293b" strokeWidth="3" />
            <rect x="634" y="182" width="18" height="26" rx="3" fill="#cbd5e1" stroke="#1e293b" strokeWidth="1.8" />
            <rect x="636" y="164" width="10" height="15" fill="#fef08a" stroke="#334155" strokeWidth="1.2" />
          </g>

          {/* 3. UNDERTRUCK & CHASSI */}
          <g
            onClick={() =>
              setFiltroZona3D(filtroZona3D === 'UNDERTRUCK' ? 'TODAS' : 'UNDERTRUCK')
            }
            className="cursor-pointer"
          >
            <rect x="42" y="197" width="470" height="12" fill="#94a3b8" stroke="#334155" strokeWidth="1.5" />
            <rect x="85" y="213" width="115" height="5" fill="#e2e8f0" stroke="#475569" strokeWidth="1.2" />
            <rect x="85" y="223" width="115" height="5" fill="#e2e8f0" stroke="#475569" strokeWidth="1.2" />
            <rect x="305" y="213" width="175" height="5" fill="#e2e8f0" stroke="#475569" strokeWidth="1.2" />
            <rect x="305" y="223" width="175" height="5" fill="#e2e8f0" stroke="#475569" strokeWidth="1.2" />
            <rect x="42" y="197" width="8" height="36" fill="#cbd5e1" stroke="#334155" strokeWidth="1.5" />
          </g>

          {/* 4. RODAS E PNEUS */}
          <g
            onClick={() => setFiltroZona3D(filtroZona3D === 'PNEU' ? 'TODAS' : 'PNEU')}
            className="cursor-pointer"
          >
            <path d="M 206,235 A 42,42 0 0,1 290,235" fill="none" stroke="#334155" strokeWidth="4" />
            <circle
              cx="248"
              cy="232"
              r="33"
              fill="#e2e8f0"
              stroke={zonaAtiva.id === 'PNEU' ? '#dc2626' : '#1e293b'}
              strokeWidth="5"
            />
            <circle cx="248" cy="232" r="19" fill="#ffffff" stroke="#334155" strokeWidth="2" />
            <circle cx="248" cy="232" r="7" fill="#475569" />

            <path d="M 538,225 A 42,42 0 0,1 622,225" fill="none" stroke="#334155" strokeWidth="3.5" />
            <circle
              cx="580"
              cy="232"
              r="33"
              fill="#e2e8f0"
              stroke={zonaAtiva.id === 'PNEU' ? '#dc2626' : '#1e293b'}
              strokeWidth="5"
            />
            <circle cx="580" cy="232" r="19" fill="#ffffff" stroke="#334155" strokeWidth="2" />
            <circle cx="580" cy="232" r="7" fill="#475569" />
          </g>

          {/* 5. AS 7 CAIXAS DE DIAGNÓSTICO ANATÔMICO */}
          {caixasAnatomicas.map((cx) => {
            const st = obterStatusCaixa(cx.id);
            const ativa = zonaAtiva.id === cx.id || filtroZona3D === cx.id;

            return (
              <g
                key={cx.id}
                transform={`translate(${cx.x}, ${cx.y})`}
                onClick={() =>
                  setFiltroZona3D(filtroZona3D === cx.id ? 'TODAS' : cx.id)
                }
                className="cursor-pointer hover:opacity-90"
              >
                <rect
                  x="0"
                  y="0"
                  width="112"
                  height="30"
                  rx="3"
                  fill={ativa ? '#1e293b' : '#cbd5e1'}
                  stroke={ativa ? st.corBg : '#475569'}
                  strokeWidth={ativa ? '2.5' : '1.5'}
                />
                <circle cx="16" cy="15" r="9" fill={st.corBg} stroke="#ffffff" strokeWidth="1.5" />
                <text
                  x="16"
                  y="19"
                  textAnchor="middle"
                  fill="#ffffff"
                  fontSize="11"
                  fontWeight="900"
                >
                  {st.tipo === 'OK' ? '✓' : st.tipo === 'ATENCAO' ? '!' : '✕'}
                </text>
                <text
                  x="64"
                  y="19"
                  textAnchor="middle"
                  fill={ativa ? '#ffffff' : '#0f172a'}
                  fontSize="10"
                  fontWeight="900"
                >
                  {cx.label}
                </text>
              </g>
            );
          })}

          {/* 6. STATUS GERAL DO VEÍCULO */}
          <g transform="translate(265, 275)">
            <rect
              x="0"
              y="0"
              width="195"
              height="34"
              rx="4"
              fill="#cbd5e1"
              stroke="#475569"
              strokeWidth="1.8"
            />
            <circle
              cx="20"
              cy="17"
              r="10"
              fill={estaDisponivel ? '#10b981' : '#f59e0b'}
              stroke="#ffffff"
              strokeWidth="1.5"
            />
            <text x="20" y="21" textAnchor="middle" fill="#ffffff" fontSize="12" fontWeight="900">
              {estaDisponivel ? '✓' : '!'}
            </text>
            <text x="110" y="21" textAnchor="middle" fill="#0f172a" fontSize="11" fontWeight="900">
              STATUS: {estaDisponivel ? 'OPERACIONAL' : 'EM OFICINA'}
            </text>
          </g>
        </svg>
      </div>
    </div>
  );
}

// ============================================================================
// --- COMPONENTE 2: VITRINE 3D REALISTA ESCULPIDA (MALHA CHANFRADA + NEON) ---
// ============================================================================
function VitrineCaminhao3D({ veiculo, zonaAtiva, filtroZona3D, setFiltroZona3D }) {
  const mountRef = useRef(null);
  const [modoVisualizacao, setModoVisualizacao] = useState('3D'); // '3D' | '2D'
  const stateRef = useRef({
    autoGirar: false,
    raioX: false,
    mostrarRotulos: true,
    cabineBasculada: false,
    portasBauAbertas: false,
    zonaDestacada: 'PREVENTIVA',
    corHexDestacada: 0x06b6d4,
    theta: -0.68,
    phi: 1.15,
    radius: 11.2,
    targetTheta: -0.68,
    targetPhi: 1.15,
    targetRadius: 11.2,
    lookAtTarget: { x: 0, y: 1.25, z: 0 },
    currentLookAt: { x: 0, y: 1.25, z: 0 },
  });

  const [autoGirarUI, setAutoGirarUI] = useState(false);
  const [raioXUI, setRaioXUI] = useState(false);
  const [rotulosUI, setRotulosUI] = useState(true);
  const [cabineUI, setCabineUI] = useState(false);
  const [bauAbertoUI, setBauAbertoUI] = useState(false);
  const [cameraAtivaUI, setCameraAtivaUI] = useState('ISO');
  const [threeCarregado, setThreeCarregado] = useState(!!window.THREE);

  useEffect(() => {
    if (!zonaAtiva) return;
    stateRef.current.zonaDestacada = zonaAtiva.id;
    stateRef.current.corHexDestacada = zonaAtiva.corHex || 0xef4444;

    const deveBascular = !!zonaAtiva.bascularCabine;
    stateRef.current.cabineBasculada = deveBascular;
    setCabineUI(deveBascular);

    const deveAbrirBau = !!zonaAtiva.abrirBau;
    stateRef.current.portasBauAbertas = deveAbrirBau;
    setBauAbertoUI(deveAbrirBau);

    const deveRaioX = !!zonaAtiva.raioXRecomendado;
    stateRef.current.raioX = deveRaioX;
    setRaioXUI(deveRaioX);

    moverCameraParaPreset(zonaAtiva.cameraPreset || 'ISO', deveBascular, deveRaioX);
  }, [veiculo, zonaAtiva?.id]);

  const moverCameraParaPreset = (preset, forcarCabine, forcarRaioX) => {
    setCameraAtivaUI(preset);
    const s = stateRef.current;

    if (preset === 'MOTOR') {
      s.cabineBasculada = true;
      setCabineUI(true);
      s.targetTheta = -0.82;
      s.targetPhi = 1.15;
      s.targetRadius = 7.6;
      s.lookAtTarget = { x: -2.4, y: 1.15, z: 0 };
    } else if (preset === 'CAMBIO') {
      s.raioX = true;
      setRaioXUI(true);
      s.targetTheta = -1.05;
      s.targetPhi = 1.22;
      s.targetRadius = 8.0;
      s.lookAtTarget = { x: -0.3, y: 0.9, z: 0 };
    } else if (preset === 'PNEUS') {
      s.raioX = true;
      setRaioXUI(true);
      s.targetTheta = -2.05;
      s.targetPhi = 1.3;
      s.targetRadius = 8.2;
      s.lookAtTarget = { x: 1.9, y: 0.75, z: 0 };
    } else if (preset === 'FRENTE') {
      s.targetTheta = -0.25;
      s.targetPhi = 1.22;
      s.targetRadius = 9.0;
      s.lookAtTarget = { x: -2.3, y: 1.45, z: 0 };
    } else if (preset === 'TRASEIRA') {
      s.targetTheta = Math.PI - 0.35;
      s.targetPhi = 1.18;
      s.targetRadius = 10.2;
      s.lookAtTarget = { x: 1.8, y: 1.65, z: 0 };
    } else if (preset === 'ESQ') {
      s.targetTheta = -Math.PI / 2;
      s.targetPhi = 1.32;
      s.targetRadius = 11.2;
      s.lookAtTarget = { x: 0, y: 1.3, z: 0 };
    } else if (preset === 'DIR') {
      s.targetTheta = Math.PI / 2;
      s.targetPhi = 1.32;
      s.targetRadius = 11.2;
      s.lookAtTarget = { x: 0, y: 1.3, z: 0 };
    } else if (preset === 'TOPO') {
      s.targetTheta = -Math.PI / 2;
      s.targetPhi = 0.16;
      s.targetRadius = 13.0;
      s.lookAtTarget = { x: 0, y: 1.0, z: 0 };
    } else {
      if (forcarCabine === undefined) {
        s.cabineBasculada = false;
        setCabineUI(false);
      }
      if (forcarRaioX === undefined) {
        s.raioX = false;
        setRaioXUI(false);
      }
      s.targetTheta = -0.68;
      s.targetPhi = 1.15;
      s.targetRadius = 11.2;
      s.lookAtTarget = { x: 0, y: 1.3, z: 0 };
    }
  };

  useEffect(() => {
    if (window.THREE) {
      setThreeCarregado(true);
      return;
    }
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js';
    script.async = true;
    script.onload = () => setThreeCarregado(true);
    document.body.appendChild(script);
  }, []);

  useEffect(() => {
    if (modoVisualizacao !== '3D' || !threeCarregado || !mountRef.current || !window.THREE) return;
    const THREE = window.THREE;
    const container = mountRef.current;
    const width = container.clientWidth || 720;
    const height = container.clientHeight || 360;

    const scene = new THREE.Scene();
    const camera = new THREE.PerspectiveCamera(40, width / height, 0.1, 100);

    const renderer = new THREE.WebGLRenderer({ antialias: true, alpha: true });
    renderer.setSize(width, height);
    renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
    renderer.shadowMap.enabled = true;
    container.innerHTML = '';
    container.appendChild(renderer.domElement);

    // --- ILUMINAÇÃO DE ESTÚDIO AUTOMOTIVO 3 PONTOS ---
    scene.add(new THREE.AmbientLight(0xffffff, 0.72));

    const sunLight = new THREE.DirectionalLight(0xffffff, 1.05);
    sunLight.position.set(12, 18, 14);
    scene.add(sunLight);

    const fillLight = new THREE.DirectionalLight(0x38bdf8, 0.5);
    fillLight.position.set(-14, 10, -12);
    scene.add(fillLight);

    const underLight = new THREE.DirectionalLight(0xfbbf24, 0.35);
    underLight.position.set(0, -6, 6);
    scene.add(underLight);

    // Piso Técnico de Engenharia
    const gridHelper = new THREE.GridHelper(20, 28, 0x0284c7, 0x1e293b);
    gridHelper.position.y = 0.01;
    scene.add(gridHelper);

    const truckGroup = new THREE.Group();
    scene.add(truckGroup);

    const materiaisExternos = [];
    const todosMeshesClicaveis = [];
    const meshesPorZona = {
      MOTOR: [],
      TRANSMISSAO: [],
      PARTIDA: [],
      UNDERTRUCK: [],
      PNEU: [],
      BAU: [],
      PREVENTIVA: [],
      SINISTRO: [],
    };

    // Clona material individual para permitir que cada peça brilhe em Neon quando selecionada
    const registrarPeca = (mesh, zonaId) => {
      mesh.material = mesh.material.clone();
      mesh.userData = {
        zonaId,
        baseColor: mesh.material.color ? mesh.material.color.getHex() : 0xffffff,
      };
      if (meshesPorZona[zonaId]) {
        meshesPorZona[zonaId].push(mesh);
      }
      todosMeshesClicaveis.push(mesh);
    };

    // ========================================================================
    // 1. CHASSI ESCALONADO E SUBESTRUTURA UNDERTRUCK
    // ========================================================================
    const matChassi = new THREE.MeshStandardMaterial({
      color: 0x64748b,
      roughness: 0.45,
      metalness: 0.65,
    });

    [-0.52, 0.52].forEach((zPos) => {
      const longarina = new THREE.Mesh(new THREE.BoxGeometry(7.35, 0.22, 0.12), matChassi);
      longarina.position.set(0.18, 0.86, zPos);
      registrarPeca(longarina, 'UNDERTRUCK');
      truckGroup.add(longarina);
    });

    for (let x = -2.9; x <= 3.5; x += 1.05) {
      const travessa = new THREE.Mesh(new THREE.BoxGeometry(0.12, 0.18, 1.12), matChassi);
      travessa.position.set(x, 0.86, 0);
      registrarPeca(travessa, 'UNDERTRUCK');
      truckGroup.add(travessa);
    }

    // Tanque Cilíndrico de Diesel + Cintas Metálicas + Tanque Azul de Arla 32
    const matTanqueInox = new THREE.MeshStandardMaterial({
      color: 0x94a3b8,
      roughness: 0.25,
      metalness: 0.85,
    });
    const tanqueDiesel = new THREE.Mesh(
      new THREE.CylinderGeometry(0.28, 0.28, 1.25, 24),
      matTanqueInox
    );
    tanqueDiesel.rotation.z = Math.PI / 2;
    tanqueDiesel.position.set(-0.25, 0.78, 0.88);
    registrarPeca(tanqueDiesel, 'UNDERTRUCK');
    truckGroup.add(tanqueDiesel);

    const matArlaAzul = new THREE.MeshStandardMaterial({
      color: 0x0284c7,
      roughness: 0.4,
      metalness: 0.2,
    });
    const tanqueArla = new THREE.Mesh(new THREE.BoxGeometry(0.38, 0.42, 0.38), matArlaAzul);
    tanqueArla.position.set(-1.1, 0.8, 0.88);
    registrarPeca(tanqueArla, 'UNDERTRUCK');
    truckGroup.add(tanqueArla);

    // Catalisador & Sistema de Exaustão DPF
    const catalisador = new THREE.Mesh(new THREE.BoxGeometry(0.85, 0.34, 0.38), matTanqueInox);
    catalisador.position.set(-0.35, 0.8, -0.88);
    registrarPeca(catalisador, 'UNDERTRUCK');
    truckGroup.add(catalisador);

    const ponteiraEscape = new THREE.Mesh(
      new THREE.CylinderGeometry(0.055, 0.055, 1.85, 14),
      matTanqueInox
    );
    ponteiraEscape.rotation.z = Math.PI / 2;
    ponteiraEscape.position.set(0.75, 0.72, -0.82);
    registrarPeca(ponteiraEscape, 'UNDERTRUCK');
    truckGroup.add(ponteiraEscape);

    // ========================================================================
    // 2. MOTOR CUMMINS / MWM 4 CILINDROS + TURBINA + RADIADOR COM HÉLICE VIVA
    // ========================================================================
    const matBlocoMotor = new THREE.MeshStandardMaterial({
      color: 0xdc2626,
      roughness: 0.35,
      metalness: 0.4,
    });
    const matAluminio = new THREE.MeshStandardMaterial({
      color: 0xe2e8f0,
      roughness: 0.3,
      metalness: 0.8,
    });
    const matPretoFosco = new THREE.MeshStandardMaterial({
      color: 0x1e293b,
      roughness: 0.6,
      metalness: 0.4,
    });

    const motorGroup = new THREE.Group();
    motorGroup.position.set(-2.38, 1.12, 0);
    truckGroup.add(motorGroup);

    // Bloco Principal e Cárter Inferior
    const bloco = new THREE.Mesh(new THREE.BoxGeometry(1.12, 0.68, 0.68), matBlocoMotor);
    registrarPeca(bloco, 'MOTOR');
    motorGroup.add(bloco);

    const carter = new THREE.Mesh(new THREE.BoxGeometry(0.95, 0.25, 0.52), matPretoFosco);
    carter.position.set(0, -0.42, 0);
    registrarPeca(carter, 'MOTOR');
    motorGroup.add(carter);

    // Cabeçote Valvulado + 4 Tampas de Válvula Individuais
    const cabecote = new THREE.Mesh(new THREE.BoxGeometry(1.04, 0.22, 0.62), matAluminio);
    cabecote.position.set(0, 0.44, 0);
    registrarPeca(cabecote, 'MOTOR');
    motorGroup.add(cabecote);

    for (let i = -0.36; i <= 0.36; i += 0.24) {
      const tampaValvula = new THREE.Mesh(
        new THREE.BoxGeometry(0.18, 0.12, 0.48),
        matBlocoMotor
      );
      tampaValvula.position.set(i, 0.58, 0);
      registrarPeca(tampaValvula, 'MOTOR');
      motorGroup.add(tampaValvula);
    }

    // Turbocompressor Lateral (Caracol) + Tubo Intercooler
    const turbina = new THREE.Mesh(new THREE.TorusGeometry(0.14, 0.06, 12, 24), matAluminio);
    turbina.position.set(0.1, 0.25, -0.42);
    registrarPeca(turbina, 'MOTOR');
    motorGroup.add(turbina);

    // Radiador Colmeia + Intercooler + Ventoinha Giratória
    const radiador = new THREE.Mesh(new THREE.BoxGeometry(0.16, 0.88, 0.98), matPretoFosco);
    radiador.position.set(-0.78, 0.12, 0);
    registrarPeca(radiador, 'MOTOR');
    motorGroup.add(radiador);

    const intercooler = new THREE.Mesh(new THREE.BoxGeometry(0.08, 0.65, 0.82), matAluminio);
    intercooler.position.set(-0.88, 0.12, 0);
    registrarPeca(intercooler, 'MOTOR');
    motorGroup.add(intercooler);

    const heliceGroup = new THREE.Group();
    heliceGroup.position.set(-0.66, 0.12, 0);
    motorGroup.add(heliceGroup);
    for (let a = 0; a < 6; a++) {
      const pa = new THREE.Mesh(new THREE.BoxGeometry(0.03, 0.34, 0.08), matBlocoMotor);
      pa.rotation.x = (a * Math.PI) / 3;
      registrarPeca(pa, 'MOTOR');
      heliceGroup.add(pa);
    }

    // ========================================================================
    // 3. TRANSMISSÃO: CAPA SECA DE EMBREAGEM + CAIXA ZF + CARDAN GIRATÓRIO
    // ========================================================================
    const matTransmissao = new THREE.MeshStandardMaterial({
      color: 0xcbd5e1,
      roughness: 0.28,
      metalness: 0.82,
    });

    // Capa Seca Cônica da Embreagem (entre o motor e o câmbio)
    const capaSeca = new THREE.Mesh(
      new THREE.CylinderGeometry(0.24, 0.36, 0.38, 24),
      matTransmissao
    );
    capaSeca.rotation.z = Math.PI / 2;
    capaSeca.position.set(-1.62, 0.98, 0);
    registrarPeca(capaSeca, 'TRANSMISSAO');
    truckGroup.add(capaSeca);

    // Corpo da Caixa de Câmbio com Trambulador Superior
    const corpoCambio = new THREE.Mesh(
      new THREE.CylinderGeometry(0.26, 0.24, 0.78, 20),
      matTransmissao
    );
    corpoCambio.rotation.z = Math.PI / 2;
    corpoCambio.position.set(-1.1, 0.96, 0);
    registrarPeca(corpoCambio, 'TRANSMISSAO');
    truckGroup.add(corpoCambio);

    const trambulador = new THREE.Mesh(
      new THREE.BoxGeometry(0.45, 0.22, 0.32),
      matTransmissao
    );
    trambulador.position.set(-1.1, 1.18, 0);
    registrarPeca(trambulador, 'TRANSMISSAO');
    truckGroup.add(trambulador);

    // Eixo Cardan Articulado (gira em tempo real!)
    const cardanMesh = new THREE.Mesh(
      new THREE.CylinderGeometry(0.075, 0.075, 2.75, 16),
      matTransmissao
    );
    cardanMesh.position.set(0.65, 0.78, 0);
    cardanMesh.rotation.z = Math.PI / 2 - 0.06;
    registrarPeca(cardanMesh, 'TRANSMISSAO');
    truckGroup.add(cardanMesh);

    // ========================================================================
    // 4. RODAS DETALHADAS (COM PARAFUSOS CROMADOS), EIXOS E FEIXE DE MOLAS
    // ========================================================================
    const matBorrachaPneu = new THREE.MeshStandardMaterial({
      color: 0x0f172a,
      roughness: 0.88,
      metalness: 0.1,
    });

    const criarRodaRealista = (x, y, z, dupla = false) => {
      const rGroup = new THREE.Group();
      rGroup.position.set(x, y, z);

      const largura = dupla ? 0.62 : 0.34;
      const pneu = new THREE.Mesh(
        new THREE.CylinderGeometry(0.58, 0.58, largura, 32),
        matBorrachaPneu
      );
      pneu.rotation.x = Math.PI / 2;
      registrarPeca(pneu, 'PNEU');
      rGroup.add(pneu);

      // Ombro arredondado lateral do pneu (Torus)
      const ombro = new THREE.Mesh(
        new THREE.TorusGeometry(0.46, 0.11, 14, 32),
        matBorrachaPneu
      );
      ombro.position.z = z > 0 ? largura / 2 - 0.04 : -largura / 2 + 0.04;
      registrarPeca(ombro, 'PNEU');
      rGroup.add(ombro);

      // Aro de Aço Prata + Cubo Redutor + 8 Parafusos Cromados
      const aro = new THREE.Mesh(
        new THREE.CylinderGeometry(0.35, 0.35, largura + 0.02, 24),
        matAluminio
      );
      aro.rotation.x = Math.PI / 2;
      registrarPeca(aro, 'PNEU');
      rGroup.add(aro);

      const cubo = new THREE.Mesh(
        new THREE.CylinderGeometry(0.15, 0.15, largura + 0.07, 16),
        matPretoFosco
      );
      cubo.rotation.x = Math.PI / 2;
      registrarPeca(cubo, 'PNEU');
      rGroup.add(cubo);

      const faceZ = z > 0 ? largura / 2 + 0.02 : -(largura / 2 + 0.02);
      for (let b = 0; b < 8; b++) {
        const ang = (b * Math.PI) / 4;
        const parafuso = new THREE.Mesh(
          new THREE.CylinderGeometry(0.022, 0.022, 0.04, 8),
          matAluminio
        );
        parafuso.rotation.x = Math.PI / 2;
        parafuso.position.set(Math.cos(ang) * 0.24, Math.sin(ang) * 0.24, faceZ);
        rGroup.add(parafuso);
      }

      truckGroup.add(rGroup);
    };

    criarRodaRealista(-2.38, 0.58, 1.04, false);
    criarRodaRealista(-2.38, 0.58, -1.04, false);
    criarRodaRealista(2.05, 0.58, 0.98, true);
    criarRodaRealista(2.05, 0.58, -0.98, true);

    // Estepe sob o chassi
    const estepe = new THREE.Mesh(
      new THREE.CylinderGeometry(0.52, 0.52, 0.28, 24),
      matBorrachaPneu
    );
    estepe.position.set(0.95, 0.58, 0.68);
    registrarPeca(estepe, 'PNEU');
    truckGroup.add(estepe);

    // Eixo Diferencial Traseiro + Feixe de Molas Parabólicas + Amortecedores
    const eixoTraseiro = new THREE.Mesh(
      new THREE.CylinderGeometry(0.11, 0.11, 2.05, 16),
      matChassi
    );
    eixoTraseiro.rotation.x = Math.PI / 2;
    eixoTraseiro.position.set(2.05, 0.58, 0);
    registrarPeca(eixoTraseiro, 'UNDERTRUCK');
    truckGroup.add(eixoTraseiro);

    const diferencial = new THREE.Mesh(new THREE.SphereGeometry(0.28, 20, 20), matChassi);
    diferencial.position.set(2.05, 0.58, 0);
    registrarPeca(diferencial, 'UNDERTRUCK');
    truckGroup.add(diferencial);

    [-0.58, 0.58].forEach((zPos) => {
      // Feixe de 3 lâminas escalonadas
      [0, 0.05, 0.1].forEach((offsetY, idx) => {
        const lamina = new THREE.Mesh(
          new THREE.BoxGeometry(1.45 - idx * 0.3, 0.045, 0.14),
          matChassi
        );
        lamina.position.set(2.05, 0.72 + offsetY, zPos);
        registrarPeca(lamina, 'UNDERTRUCK');
        truckGroup.add(lamina);
      });
      // Amortecedor inclinado
      const amortecedor = new THREE.Mesh(
        new THREE.CylinderGeometry(0.04, 0.04, 0.55, 10),
        matBlocoMotor
      );
      amortecedor.position.set(1.75, 0.78, zPos * 1.15);
      amortecedor.rotation.z = 0.35;
      registrarPeca(amortecedor, 'UNDERTRUCK');
      truckGroup.add(amortecedor);
    });

    // ========================================================================
    // 5. CABINE ESCULPIDA EM MALHA CHANFRADA (EXTRUDE GEOMETRY VW DELIVERY / 8.120)
    // ========================================================================
    const cabinePivot = new THREE.Group();
    cabinePivot.position.set(-3.35, 0.92, 0);
    truckGroup.add(cabinePivot);

    const matPinturaCabine = new THREE.MeshStandardMaterial({
      color: 0xf8fafc,
      roughness: 0.18,
      metalness: 0.15,
      transparent: true,
      opacity: 1.0,
    });
    const matVidroFume = new THREE.MeshStandardMaterial({
      color: 0x0f172a,
      roughness: 0.08,
      metalness: 0.92,
      transparent: true,
      opacity: 0.86,
    });

    // Perfil lateral anatômico real da cabine VW (frente chanfrada e teto curvo)
    const cabShape = new THREE.Shape();
    cabShape.moveTo(0, 0);
    cabShape.lineTo(1.72, 0);
    cabShape.lineTo(1.72, 1.78);
    cabShape.lineTo(0.48, 1.78);
    cabShape.quadraticCurveTo(0.28, 1.78, 0.18, 1.55);
    cabShape.lineTo(0.0, 0.78);
    cabShape.lineTo(0, 0);

    const cabExtrudeGeo = new THREE.ExtrudeGeometry(cabShape, {
      depth: 2.02,
      bevelEnabled: true,
      bevelSegments: 4,
      steps: 1,
      bevelSize: 0.07,
      bevelThickness: 0.07,
    });
    cabExtrudeGeo.translate(0, 0, -1.01);

    const malhaCabine = new THREE.Mesh(cabExtrudeGeo, matPinturaCabine);
    registrarPeca(malhaCabine, 'PARTIDA');
    materiaisExternos.push(malhaCabine.material);
    cabinePivot.add(malhaCabine);

    // Quebra-Vento Aerodinâmico Esculpido no Teto da Cabine
    const aeroShape = new THREE.Shape();
    aeroShape.moveTo(0.45, 1.82);
    aeroShape.lineTo(1.74, 1.82);
    aeroShape.lineTo(1.74, 2.46);
    aeroShape.quadraticCurveTo(1.05, 2.38, 0.45, 1.82);

    const aeroGeo = new THREE.ExtrudeGeometry(aeroShape, {
      depth: 1.92,
      bevelEnabled: true,
      bevelSegments: 3,
      steps: 1,
      bevelSize: 0.04,
      bevelThickness: 0.04,
    });
    aeroGeo.translate(0, 0, -0.96);

    const malhaQuebraVento = new THREE.Mesh(aeroGeo, matPinturaCabine);
    registrarPeca(malhaQuebraVento, 'PARTIDA');
    materiaisExternos.push(malhaQuebraVento.material);
    cabinePivot.add(malhaQuebraVento);

    // Para-brisa Curvo Panorâmico + Janelas Laterais Chanfradas
    const parabrisa = new THREE.Mesh(new THREE.BoxGeometry(0.1, 0.82, 1.88), matVidroFume);
    parabrisa.position.set(0.08, 1.22, 0);
    parabrisa.rotation.z = -0.22;
    registrarPeca(parabrisa, 'PARTIDA');
    materiaisExternos.push(parabrisa.material);
    cabinePivot.add(parabrisa);

    const janelasLaterais = new THREE.Mesh(
      new THREE.BoxGeometry(1.08, 0.68, 2.18),
      matVidroFume
    );
    janelasLaterais.position.set(0.88, 1.24, 0);
    materiaisExternos.push(janelasLaterais.material);
    cabinePivot.add(janelasLaterais);

    // Interior da Cabine (Bancos Duplos + Painel Digital + Volante Inclinado)
    const matInterior = new THREE.MeshStandardMaterial({ color: 0x1e293b, roughness: 0.7 });
    const painel = new THREE.Mesh(new THREE.BoxGeometry(0.42, 0.36, 1.88), matInterior);
    painel.position.set(0.38, 0.78, 0);
    registrarPeca(painel, 'PARTIDA');
    cabinePivot.add(painel);

    const volante = new THREE.Mesh(new THREE.TorusGeometry(0.18, 0.025, 10, 24), matPretoFosco);
    volante.position.set(0.55, 1.02, 0.52);
    volante.rotation.y = Math.PI / 2;
    volante.rotation.x = 0.45;
    cabinePivot.add(volante);

    [-0.52, 0.52].forEach((zSeat) => {
      const banco = new THREE.Mesh(new THREE.BoxGeometry(0.45, 0.68, 0.52), matInterior);
      banco.position.set(1.05, 0.62, zSeat);
      cabinePivot.add(banco);
    });

    // Grade Frontal VW com Frisos Cromados + Emblema Redondo Central + Faróis Duplos
    const gradeFrontal = new THREE.Mesh(
      new THREE.BoxGeometry(0.08, 0.52, 1.65),
      matPretoFosco
    );
    gradeFrontal.position.set(-0.04, 0.44, 0);
    registrarPeca(gradeFrontal, 'PARTIDA');
    cabinePivot.add(gradeFrontal);

    const emblemaVW = new THREE.Mesh(
      new THREE.CylinderGeometry(0.13, 0.13, 0.06, 24),
      matAluminio
    );
    emblemaVW.rotation.z = Math.PI / 2;
    emblemaVW.position.set(-0.07, 0.48, 0);
    cabinePivot.add(emblemaVW);

    const matFarolLed = new THREE.MeshStandardMaterial({
      color: 0xfef08a,
      emissive: 0xfacc15,
      emissiveIntensity: 0.75,
    });
    [-0.84, 0.84].forEach((zPos) => {
      const blocoFarol = new THREE.Mesh(
        new THREE.BoxGeometry(0.1, 0.24, 0.34),
        matFarolLed
      );
      blocoFarol.position.set(-0.04, 0.28, zPos);
      registrarPeca(blocoFarol, 'PARTIDA');
      cabinePivot.add(blocoFarol);
    });

    // Para-lamas Redondos Arqueados + Degrau de Estribo + Retrovisores com Haste Tubular
    [-1.08, 1.08].forEach((zPos) => {
      const paralama = new THREE.Mesh(
        new THREE.TorusGeometry(0.62, 0.09, 12, 24, Math.PI),
        matPretoFosco
      );
      paralama.position.set(0.97, -0.22, zPos);
      cabinePivot.add(paralama);

      // Retrovisor Principal + Auxiliar Convexo + Braço Metálico
      const espelhoGrande = new THREE.Mesh(
        new THREE.BoxGeometry(0.12, 0.46, 0.16),
        matPretoFosco
      );
      espelhoGrande.position.set(0.28, 1.28, zPos * 1.15);
      cabinePivot.add(espelhoGrande);

      const hasteEspelho = new THREE.Mesh(
        new THREE.CylinderGeometry(0.018, 0.018, 0.25, 8),
        matPretoFosco
      );
      hasteEspelho.rotation.x = Math.PI / 2;
      hasteEspelho.position.set(0.28, 1.46, zPos * 1.05);
      cabinePivot.add(hasteEspelho);
    });

    // Para-choque Dianteiro Envolvente (associado à zona SINISTRO / PARTIDA)
    const parachoqueFrontal = new THREE.Mesh(
      new THREE.BoxGeometry(0.32, 0.36, 2.26),
      matPinturaCabine
    );
    parachoqueFrontal.position.set(-3.45, 0.74, 0);
    registrarPeca(parachoqueFrontal, 'SINISTRO');
    truckGroup.add(parachoqueFrontal);

    // ========================================================================
    // 6. BAÚ DE CARGA PEPSICO + PORTAS TRASEIRAS ARTICULADAS QUE ABREM EM 3D!
    // ========================================================================
    const criarTexturaBauAltaDefinicao = () => {
      const canvas = document.createElement('canvas');
      canvas.width = 1024;
      canvas.height = 512;
      const ctx = canvas.getContext('2d');

      ctx.fillStyle = '#f8fafc';
      ctx.fillRect(0, 0, 1024, 512);

      // Frisos horizontais de alumínio
      ctx.strokeStyle = '#e2e8f0';
      ctx.lineWidth = 3;
      for (let y = 36; y < 512; y += 38) {
        ctx.beginPath();
        ctx.moveTo(0, y);
        ctx.lineTo(1024, y);
        ctx.stroke();
      }

      // Cantoneiras estruturais e rebites
      ctx.strokeStyle = '#64748b';
      ctx.lineWidth = 18;
      ctx.strokeRect(9, 9, 1006, 494);

      // Porta Lateral de Entrega com Dobradiças e Fecho
      ctx.fillStyle = '#f1f5f9';
      ctx.fillRect(465, 80, 165, 402);
      ctx.strokeStyle = '#334155';
      ctx.lineWidth = 6;
      ctx.strokeRect(465, 80, 165, 402);
      ctx.fillStyle = '#475569';
      ctx.fillRect(612, 120, 26, 14);
      ctx.fillRect(612, 420, 26, 14);
      ctx.fillRect(488, 445, 44, 18);

      // Globo PepsiCo + Tipografia Oficial
      ctx.save();
      ctx.strokeStyle = '#0f172a';
      ctx.lineWidth = 5;
      ctx.beginPath();
      ctx.arc(135, 240, 40, 0, Math.PI * 2);
      ctx.stroke();
      ctx.beginPath();
      ctx.ellipse(135, 240, 40, 16, -0.2, 0, Math.PI * 2);
      ctx.stroke();
      ctx.beginPath();
      ctx.ellipse(135, 240, 16, 40, -0.2, 0, Math.PI * 2);
      ctx.stroke();

      ctx.fillStyle = '#0f172a';
      ctx.font = '900 56px sans-serif';
      ctx.fillText('PEPSICO', 192, 275);
      ctx.font = 'bold 22px sans-serif';
      ctx.fillStyle = '#0284c7';
      ctx.fillText('LOGÍSTICA SUL2 • VW 8.120', 195, 312);
      ctx.restore();

      return new THREE.CanvasTexture(canvas);
    };

    const texLateralBau = criarTexturaBauAltaDefinicao();
    const matBauLateral = new THREE.MeshStandardMaterial({
      map: texLateralBau,
      roughness: 0.28,
      metalness: 0.12,
      transparent: true,
      opacity: 1.0,
    });
    const matBauTeto = new THREE.MeshStandardMaterial({
      color: 0xf1f5f9,
      roughness: 0.35,
      metalness: 0.18,
      transparent: true,
      opacity: 1.0,
    });

    const bauCarga = new THREE.Mesh(new THREE.BoxGeometry(5.05, 2.55, 2.32), [
      matBauTeto,
      matBauTeto,
      matBauTeto,
      matBauTeto,
      matBauLateral,
      matBauLateral,
    ]);
    bauCarga.position.set(1.15, 2.28, 0);
    // Clona e registra todos os materiais do baú para Raio-X e Neon
    bauCarga.material = bauCarga.material.map((m) => {
      const c = m.clone();
      materiaisExternos.push(c);
      return c;
    });
    todosMeshesClicaveis.push(bauCarga);
    bauCarga.userData = { zonaId: 'BAU', baseColor: 0xf8fafc };
    meshesPorZona.BAU.push(bauCarga);
    truckGroup.add(bauCarga);

    // Portas Traseiras Duplas Articuladas (Abrem quando seleciona BAÚ!)
    const portaTraseiraEsqPivot = new THREE.Group();
    portaTraseiraEsqPivot.position.set(3.68, 2.28, 1.14);
    truckGroup.add(portaTraseiraEsqPivot);

    const folhaPortaEsq = new THREE.Mesh(
      new THREE.BoxGeometry(0.08, 2.42, 1.12),
      matBauTeto
    );
    folhaPortaEsq.position.set(0, 0, -0.56);
    registrarPeca(folhaPortaEsq, 'BAU');
    materiaisExternos.push(folhaPortaEsq.material);
    portaTraseiraEsqPivot.add(folhaPortaEsq);

    const portaTraseiraDirPivot = new THREE.Group();
    portaTraseiraDirPivot.position.set(3.68, 2.28, -1.14);
    truckGroup.add(portaTraseiraDirPivot);

    const folhaPortaDir = new THREE.Mesh(
      new THREE.BoxGeometry(0.08, 2.42, 1.12),
      matBauTeto
    );
    folhaPortaDir.position.set(0, 0, 0.56);
    registrarPeca(folhaPortaDir, 'BAU');
    materiaisExternos.push(folhaPortaDir.material);
    portaTraseiraDirPivot.add(folhaPortaDir);

    // Barras Duplas de Proteção Ciclista Laterais
    [-1.14, 1.14].forEach((zPos) => {
      [0.86, 0.68].forEach((yPos) => {
        const protetor = new THREE.Mesh(
          new THREE.BoxGeometry(2.15, 0.08, 0.04),
          matAluminio
        );
        protetor.position.set(0.15, yPos, zPos);
        registrarPeca(protetor, 'UNDERTRUCK');
        truckGroup.add(protetor);
      });
    });

    // ========================================================================
    // 7. ANEL HOLOGRÁFICO E PULSO NEON NA PEÇA AFETADA
    // ========================================================================
    const matAnelAlerta = new THREE.MeshBasicMaterial({
      color: 0xff3300,
      side: THREE.DoubleSide,
      transparent: true,
      opacity: 0.9,
    });
    const anelLaser = new THREE.Mesh(new THREE.RingGeometry(0.32, 0.42, 32), matAnelAlerta);
    anelLaser.rotation.x = Math.PI / 2;
    truckGroup.add(anelLaser);

    const coordsPorZona3D = {
      MOTOR: { x: -2.38, y: 1.25, z: 0 },
      TRANSMISSAO: { x: -1.15, y: 0.96, z: 0 },
      PARTIDA: { x: -2.55, y: 1.85, z: 0 },
      UNDERTRUCK: { x: 0.4, y: 0.78, z: 0.78 },
      PNEU: { x: 2.05, y: 0.58, z: 1.05 },
      BAU: { x: 1.45, y: 2.35, z: 0 },
      PREVENTIVA: { x: -0.5, y: 2.35, z: 0 },
      SINISTRO: { x: -3.35, y: 1.05, z: 0 },
    };

    // ========================================================================
    // INTERAÇÃO DE MOUSE (ARRASTAR 360°, SCROLL ZOOM E CLIQUE DIRETO NA PEÇA 3D)
    // ========================================================================
    const raycaster = new THREE.Raycaster();
    const mouseVec = new THREE.Vector2();
    let isDragging = false;
    let moveuMouse = false;
    let prevX = 0;
    let prevY = 0;

    const onPointerDown = (e) => {
      isDragging = true;
      moveuMouse = false;
      prevX = e.clientX;
      prevY = e.clientY;
    };

    const onPointerMove = (e) => {
      if (!isDragging) return;
      const dx = e.clientX - prevX;
      const dy = e.clientY - prevY;
      if (Math.abs(dx) > 3 || Math.abs(dy) > 3) moveuMouse = true;
      prevX = e.clientX;
      prevY = e.clientY;

      stateRef.current.targetTheta -= dx * 0.0085;
      stateRef.current.targetPhi = Math.max(
        0.12,
        Math.min(Math.PI / 2 - 0.04, stateRef.current.targetPhi - dy * 0.0085)
      );
    };

    const onPointerUp = (e) => {
      if (isDragging && !moveuMouse) {
        // Se foi um clique rápido na peça 3D, descobre qual peça foi clicada!
        const rect = renderer.domElement.getBoundingClientRect();
        mouseVec.x = ((e.clientX - rect.left) / rect.width) * 2 - 1;
        mouseVec.y = -((e.clientY - rect.top) / rect.height) * 2 + 1;
        raycaster.setFromCamera(mouseVec, camera);
        const intersects = raycaster.intersectObjects(todosMeshesClicaveis, false);
        if (intersects.length > 0) {
          const zClicada = intersects[0].object.userData?.zonaId;
          if (zClicada) {
            stateRef.current.zonaDestacada = zClicada;
            setFiltroZona3D((atual) => (atual === zClicada ? 'TODAS' : zClicada));
          }
        }
      }
      isDragging = false;
    };

    const onWheel = (e) => {
      e.preventDefault();
      stateRef.current.targetRadius = Math.max(
        5.2,
        Math.min(18.0, stateRef.current.targetRadius + e.deltaY * 0.006)
      );
    };

    const domElem = renderer.domElement;
    domElem.addEventListener('pointerdown', onPointerDown);
    window.addEventListener('pointermove', onPointerMove);
    window.addEventListener('pointerup', onPointerUp);
    domElem.addEventListener('wheel', onWheel, { passive: false });

    // ========================================================================
    // LOOP DE ANIMAÇÃO 60 FPS (NEON PULSANTE + PEÇAS VIVAS)
    // ========================================================================
    let reqId;
    let clock = 0;

    const animate = () => {
      reqId = requestAnimationFrame(animate);
      clock += 0.065;
      const s = stateRef.current;

      if (s.autoGirar && !isDragging) {
        s.targetTheta += 0.006;
      }

      s.theta += (s.targetTheta - s.theta) * 0.1;
      s.phi += (s.targetPhi - s.phi) * 0.1;
      s.radius += (s.targetRadius - s.radius) * 0.1;

      s.currentLookAt.x += (s.lookAtTarget.x - s.currentLookAt.x) * 0.1;
      s.currentLookAt.y += (s.lookAtTarget.y - s.currentLookAt.y) * 0.1;
      s.currentLookAt.z += (s.lookAtTarget.z - s.currentLookAt.z) * 0.1;

      camera.position.x =
        s.currentLookAt.x + s.radius * Math.sin(s.phi) * Math.cos(s.theta);
      camera.position.y = s.currentLookAt.y + s.radius * Math.cos(s.phi);
      camera.position.z =
        s.currentLookAt.z + s.radius * Math.sin(s.phi) * Math.sin(s.theta);
      camera.lookAt(s.currentLookAt.x, s.currentLookAt.y, s.currentLookAt.z);

      // 1. Basculamento suave da Cabine (42 graus)
      const anguloAlvoCabine = s.cabineBasculada ? 0.74 : 0.0;
      cabinePivot.rotation.z += (anguloAlvoCabine - cabinePivot.rotation.z) * 0.09;

      // 2. Abertura suave das Portas Traseiras do Baú
      const anguloPorta = s.portasBauAbertas ? 1.65 : 0.0;
      portaTraseiraEsqPivot.rotation.y +=
        (anguloPorta - portaTraseiraEsqPivot.rotation.y) * 0.09;
      portaTraseiraDirPivot.rotation.y +=
        (-anguloPorta - portaTraseiraDirPivot.rotation.y) * 0.09;

      // 3. Rotação viva da Hélice do Radiador e do Eixo Cardan
      heliceGroup.rotation.x += 0.18;
      cardanMesh.rotation.x += 0.12;

      // 4. Modo Raio-X inteligente
      const opacidadeAlvo = s.raioX ? 0.16 : 1.0;
      materiaisExternos.forEach((mat) => {
        mat.opacity += (opacidadeAlvo - mat.opacity) * 0.12;
        mat.depthWrite = mat.opacity > 0.5;
      });

      // 5. BRILHO NEON PULSANTE APENAS NA PEÇA COM DEFEITO!
      const pulsoNeon = 0.45 + Math.abs(Math.sin(clock * 2.4)) * 0.55;
      Object.keys(meshesPorZona).forEach((zKey) => {
        const ehZonaAfetada = zKey === s.zonaDestacada;
        meshesPorZona[zKey].forEach((m) => {
          const mats = Array.isArray(m.material) ? m.material : [m.material];
          mats.forEach((mat) => {
            if (!mat.emissive) return;
            if (ehZonaAfetada && s.zonaDestacada !== 'PREVENTIVA') {
              mat.emissive.setHex(s.corHexDestacada);
              mat.emissiveIntensity = pulsoNeon * 0.85;
            } else {
              mat.emissive.setHex(0x000000);
              mat.emissiveIntensity = 0;
            }
          });
        });
      });

      // 6. Anel Laser no ponto da avaria
      const posAlerta = coordsPorZona3D[s.zonaDestacada] || coordsPorZona3D.PREVENTIVA;
      anelLaser.position.set(posAlerta.x, posAlerta.y, posAlerta.z);
      matAnelAlerta.color.setHex(s.corHexDestacada);
      const escalaAnel = 1 + Math.sin(clock * 2.5) * 0.3;
      anelLaser.scale.set(escalaAnel, escalaAnel, 1);

      renderer.render(scene, camera);
    };

    animate();

    const handleResize = () => {
      if (!container) return;
      const w = container.clientWidth;
      const h = container.clientHeight;
      camera.aspect = w / h;
      camera.updateProjectionMatrix();
      renderer.setSize(w, h);
    };
    window.addEventListener('resize', handleResize);

    return () => {
      cancelAnimationFrame(reqId);
      domElem.removeEventListener('pointerdown', onPointerDown);
      window.removeEventListener('pointermove', onPointerMove);
      window.removeEventListener('pointerup', onPointerUp);
      domElem.removeEventListener('wheel', onWheel);
      window.removeEventListener('resize', handleResize);
      renderer.dispose();
    };
  }, [threeCarregado, modoVisualizacao]);

  const toggleGirar = () => {
    const novo = !autoGirarUI;
    setAutoGirarUI(novo);
    stateRef.current.autoGirar = novo;
  };

  const toggleRaioX = () => {
    const novo = !raioXUI;
    setRaioXUI(novo);
    stateRef.current.raioX = novo;
  };

  const toggleRotulos = () => {
    const novo = !rotulosUI;
    setRotulosUI(novo);
    stateRef.current.mostrarRotulos = novo;
  };

  const toggleCabine = () => {
    const novo = !cabineUI;
    setCabineUI(novo);
    stateRef.current.cabineBasculada = novo;
  };

  const togglePortasBau = () => {
    const novo = !bauAbertoUI;
    setBauAbertoUI(novo);
    stateRef.current.portasBauAbertas = novo;
  };

  const limparFiltros3D = () => {
    setAutoGirarUI(false);
    setRaioXUI(false);
    setCabineUI(false);
    setBauAbertoUI(false);
    stateRef.current.autoGirar = false;
    stateRef.current.raioX = false;
    stateRef.current.cabineBasculada = false;
    stateRef.current.portasBauAbertas = false;
    setFiltroZona3D('TODAS');
    moverCameraParaPreset('ISO', false, false);
  };

  return (
    <div className="lg:col-span-8 rounded-2xl border border-gray-700/80 bg-[#141d2b] p-4 shadow-2xl relative overflow-hidden flex flex-col justify-between">
      {/* BARRA DE CONTROLE SUPERIOR */}
      <div className="space-y-2.5 relative z-20">
        <div className="flex flex-wrap items-center justify-between gap-2 border-b border-gray-800 pb-2.5">
          {/* ALTERNADOR VITRINE 3D vs ESQUEMA 2D */}
          <div className="flex items-center gap-1.5 bg-[#0B1329] p-1 rounded-xl border border-gray-700">
            <button
              onClick={() => setModoVisualizacao('3D')}
              className={`px-3 py-1.5 rounded-lg text-[11px] font-black uppercase flex items-center gap-1.5 transition-all ${
                modoVisualizacao === '3D'
                  ? 'bg-cyan-500 text-gray-950 shadow-md'
                  : 'text-gray-300 hover:text-white'
              }`}
            >
              <Box size={14} />
              <span>Vitrine 3D (360°)</span>
            </button>
            <button
              onClick={() => setModoVisualizacao('2D')}
              className={`px-3 py-1.5 rounded-lg text-[11px] font-black uppercase flex items-center gap-1.5 transition-all ${
                modoVisualizacao === '2D'
                  ? 'bg-yellow-400 text-gray-950 shadow-md'
                  : 'text-gray-300 hover:text-white'
              }`}
            >
              <Layers size={14} />
              <span>Esquema 2D (VW 8.120)</span>
            </button>
          </div>

          {modoVisualizacao === '3D' && (
            <div className="flex flex-wrap items-center gap-1.5">
              <button
                onClick={toggleGirar}
                className={`px-2.5 py-1 rounded text-[10px] font-black uppercase border transition-all ${
                  autoGirarUI
                    ? 'bg-yellow-400 text-gray-950 border-yellow-300'
                    : 'bg-[#1f2937]/90 text-yellow-400 border-yellow-500/40'
                }`}
              >
                GIRAR
              </button>
              <button
                onClick={toggleRaioX}
                className={`px-2.5 py-1 rounded text-[10px] font-black uppercase border transition-all ${
                  raioXUI
                    ? 'bg-yellow-400 text-gray-950 border-yellow-300'
                    : 'bg-[#1f2937]/90 text-yellow-400 border-yellow-500/40'
                }`}
              >
                RAIO-X
              </button>
              <button
                onClick={toggleRotulos}
                className={`px-2.5 py-1 rounded text-[10px] font-black uppercase border transition-all ${
                  rotulosUI
                    ? 'bg-yellow-400 text-gray-950 border-yellow-300'
                    : 'bg-[#1f2937]/90 text-yellow-400 border-yellow-500/40'
                }`}
              >
                RÓTULOS
              </button>
              <button
                onClick={toggleCabine}
                className={`px-2.5 py-1 rounded text-[10px] font-black uppercase border transition-all ${
                  cabineUI
                    ? 'bg-yellow-400 text-gray-950 border-yellow-300 shadow-lg'
                    : 'bg-[#1f2937]/90 text-yellow-400 border-yellow-500/40'
                }`}
              >
                CABINE {cabineUI ? '(ABERTA)' : ''}
              </button>
              <button
                onClick={togglePortasBau}
                className={`px-2.5 py-1 rounded text-[10px] font-black uppercase border transition-all ${
                  bauAbertoUI
                    ? 'bg-yellow-400 text-gray-950 border-yellow-300 shadow-lg'
                    : 'bg-[#1f2937]/90 text-yellow-400 border-yellow-500/40'
                }`}
              >
                PORTAS BAÚ
              </button>
              <button
                onClick={limparFiltros3D}
                className="px-2.5 py-1 rounded text-[10px] font-black uppercase border bg-[#1f2937]/90 text-yellow-400 border-yellow-500/40 hover:bg-yellow-400 hover:text-gray-950"
              >
                LIMPAR
              </button>
            </div>
          )}
        </div>

        {/* ATALHOS DE CÂMERA 3D */}
        {modoVisualizacao === '3D' && (
          <div className="flex flex-wrap items-center justify-between gap-1.5">
            <div className="flex flex-wrap items-center gap-1.5">
              {[
                { id: 'MOTOR', label: 'MOTOR', zona: 'MOTOR' },
                { id: 'CAMBIO', label: 'TRANSMISSÃO / UNDERTRUCK', zona: 'TRANSMISSAO' },
                { id: 'PNEUS', label: 'PNEUS E EIXOS', zona: 'PNEU' },
                { id: 'FRENTE', label: 'PARTIDA / FRENTE', zona: 'PARTIDA' },
                { id: 'TRASEIRA', label: 'BAÚ / TRASEIRA', zona: 'BAU' },
                { id: 'ESQ', label: 'ESQ.' },
                { id: 'DIR', label: 'DIR.' },
                { id: 'TOPO', label: 'TOPO' },
              ].map((btn) => (
                <button
                  key={btn.id}
                  onClick={() => {
                    moverCameraParaPreset(btn.id);
                    if (btn.zona) {
                      stateRef.current.zonaDestacada = btn.zona;
                      setFiltroZona3D(btn.zona);
                    }
                  }}
                  className={`px-2.5 py-1 rounded text-[10px] font-black uppercase border transition-all ${
                    cameraAtivaUI === btn.id
                      ? 'bg-yellow-400 text-gray-950 border-yellow-300'
                      : 'bg-[#1f2937]/70 text-yellow-400/90 border-yellow-500/30 hover:border-yellow-400'
                  }`}
                >
                  {btn.label}
                </button>
              ))}
            </div>
            <span className="text-[10px] font-bold text-gray-400 hidden xl:inline">
              🖱️ Arraste 360° • Clique na peça 3D p/ filtrar
            </span>
          </div>
        )}
      </div>

      {/* PALCO CENTRAL 3D OU 2D */}
      {modoVisualizacao === '3D' ? (
        <div className="relative my-2 flex-1 min-h-[350px] w-full rounded-xl overflow-hidden bg-gradient-to-b from-[#1a2436] to-[#0f172a] border border-gray-800/80">
          <div
            ref={mountRef}
            className="w-full h-[350px] cursor-grab active:cursor-grabbing"
          />

          {rotulosUI && veiculo && (
            <div className="absolute bottom-3 left-3 right-3 flex flex-wrap items-center justify-between gap-2 pointer-events-none">
              <div className="bg-[#0B1329]/90 backdrop-blur-md border border-yellow-400/40 px-3.5 py-2 rounded-xl shadow-lg flex items-center gap-2.5">
                <span
                  className="w-6 h-6 rounded-full flex items-center justify-center text-xs font-black text-white"
                  style={{ backgroundColor: zonaAtiva.cor }}
                >
                  {zonaAtiva.simbolo}
                </span>
                <div>
                  <div className="flex items-center gap-2">
                    <span className="font-mono font-black text-xs text-yellow-400">
                      PLACA : {veiculo.placa}
                    </span>
                    <span className="text-[10px] font-black text-white uppercase">
                      • {zonaAtiva.nome}
                    </span>
                  </div>
                  <p className="text-[11px] font-bold text-rose-400">
                    Diagnóstico: {veiculo.status || 'Em diagnóstico'}
                  </p>
                </div>
              </div>

              <div className="bg-[#0B1329]/85 border border-gray-700 px-3 py-1.5 rounded-lg text-[10px] font-bold text-gray-300">
                Oficina: <span className="text-white font-black">{veiculo.oficina}</span> | CDV:{' '}
                <span className="text-cyan-400 font-black">{veiculo.cdv}</span>
              </div>
            </div>
          )}
        </div>
      ) : (
        <div className="my-2">
          <PlantaTecnica2DCaminhao
            veiculo={veiculo}
            zonaAtiva={zonaAtiva}
            filtroZona3D={filtroZona3D}
            setFiltroZona3D={setFiltroZona3D}
          />
        </div>
      )}

      {/* BARRA DE FILTROS INFERIOR COM FLEX-WRAP (SEM CORTAR OS TEXTOS!) */}
      <div className="pt-2.5 border-t border-gray-800 flex flex-wrap items-center justify-between gap-2">
        <span className="text-[10px] font-black text-gray-400 uppercase tracking-wider shrink-0">
          Filtrar Tabela por Sistema:
        </span>
        <div className="flex flex-wrap items-center gap-1.5">
          {[
            { id: 'TODAS', label: 'Todos', cor: '#06b6d4' },
            { id: 'MOTOR', label: 'Motor', cor: '#ef4444' },
            { id: 'TRANSMISSAO', label: 'Transmissão', cor: '#f97316' },
            { id: 'PARTIDA', label: 'Partida/Elétrica', cor: '#eab308' },
            { id: 'UNDERTRUCK', label: 'Undertruck', cor: '#a855f7' },
            { id: 'PNEU', label: 'Pneus/Rodas', cor: '#ec4899' },
            { id: 'BAU', label: 'Baú', cor: '#3b82f6' },
            { id: 'PREVENTIVA', label: 'Preventiva', cor: '#10b981' },
          ].map((z) => {
            const ativo = filtroZona3D === z.id;
            return (
              <button
                key={z.id}
                onClick={() => {
                  setFiltroZona3D(z.id);
                  if (z.id !== 'TODAS') stateRef.current.zonaDestacada = z.id;
                }}
                className={`text-[11px] font-black px-3 py-1.5 rounded-lg border flex items-center gap-1.5 whitespace-nowrap transition-all ${
                  ativo
                    ? 'bg-cyan-500 text-gray-950 border-cyan-300 shadow-md'
                    : 'bg-[#1C2541] text-gray-200 border-gray-700 hover:border-cyan-500/50'
                }`}
              >
                <span
                  className="w-2.5 h-2.5 rounded-full shrink-0"
                  style={{ backgroundColor: ativo ? '#082f49' : z.cor }}
                />
                <span>{z.label}</span>
              </button>
            );
          })}
        </div>
      </div>
    </div>
  );
}

// ============================================================================
// --- FUNÇÕES AUXILIARES DE TRATAMENTO DO EXCEL ---
// ============================================================================
function formatarDataVersaoBancoNoTopo(raw) {
  if (!raw || String(raw).trim() === '' || String(raw).trim() === '-') return null;
  if (typeof raw === 'number') {
    const UTC_DAYS = Math.floor(raw - 25569);
    const UTC_VALUE = UTC_DAYS * 86400;
    const dt = new Date(UTC_VALUE * 1000);
    if (isNaN(dt.getTime())) return null;
    return dt.toISOString().split('T')[0];
  }
  const textoData = String(raw).trim();
  if (textoData.includes('/')) {
    const partes = textoData.split('/');
    if (partes.length === 3) {
      const dia = partes[0].padStart(2, '0');
      const mes = partes[1].padStart(2, '0');
      const ano = partes[2].length === 2 ? `20${partes[2]}` : partes[2];
      const candidato = `${ano}-${mes}-${dia}`;
      return /^\d{4}-\d{2}-\d{2}$/.test(candidato) ? candidato : null;
    }
  }
  if (/^\d{4}-\d{2}-\d{2}/.test(textoData)) {
    return textoData.slice(0, 10);
  }
  return null;
}

function extrairNumeroMesPorExtenso(texto) {
  if (!texto) return null;
  const t = String(texto).toUpperCase().trim();
  if (t.includes('JAN') || t === '1' || t === '01') return 1;
  if (t.includes('FEV') || t === '2' || t === '02') return 2;
  if (t.includes('MAR') || t === '3' || t === '03') return 3;
  if (t.includes('ABR') || t === '4' || t === '04') return 4;
  if (t.includes('MAI') || t === '5' || t === '05') return 5;
  if (t.includes('JUN') || t === '6' || t === '06') return 6;
  if (t.includes('JUL') || t === '7' || t === '07') return 7;
  if (t.includes('AGO') || t === '8' || t === '08') return 8;
  if (t.includes('SET') || t === '9' || t === '09') return 9;
  if (t.includes('OUT') || t === '10') return 10;
  if (t.includes('NOV') || t === '11') return 11;
  if (t.includes('DEZ') || t === '12') return 12;
  return null;
}

function formatarMoeda(valor) {
  const num = Number(valor) || 0;
  return num.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function limparNumeroMoeda(val) {
  if (val === undefined || val === null || val === '') return 0;
  if (typeof val === 'number') return val;
  const limpo = String(val)
    .replace(/[R$\s]/g, '')
    .replace(/\./g, '')
    .replace(',', '.');
  return Number(limpo) || 0;
}

function gerarCarimboDataHoraAtual() {
  const agora = new Date();
  return `${agora.toLocaleDateString('pt-BR')} às ${agora.toLocaleTimeString('pt-BR', {
    hour: '2-digit',
    minute: '2-digit',
  })}`;
}

function formatarCreatedAtBanco(isoStr) {
  if (!isoStr) return null;
  const d = new Date(isoStr);
  if (isNaN(d.getTime())) return null;
  return `${d.toLocaleDateString('pt-BR')} às ${d.toLocaleTimeString('pt-BR', {
    hour: '2-digit',
    minute: '2-digit',
  })}`;
}

// ============================================================================
// --- COMPONENTE PRINCIPAL (MODO COCKPIT WIDESCREEN ZERO-SCROLL) ---
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

  // --- ESTADOS GERAIS & COCKPIT WIDESCREEN ---
  const [activeTab, setActiveTab] = useState('frota');
  const [filialAtual, setFilialAtual] = useState('Todas');
  const [subAbaFrota, setSubAbaFrota] = useState('preventivas');
  const [carregando, setCarregando] = useState(false);
  const [rotacionar, setRotacionar] = useState(false);

  // Controles de Layout Zero-Scroll (Sidebar Retrátil e Gavetas)
  const [sidebarAberta, setSidebarAberta] = useState(false);
  const [mostrarGavetaUploadFin, setMostrarGavetaUploadFin] = useState(false);
  const [mostrarCtaExpandido, setMostrarCtaExpandido] = useState(false);
  const [subVisaoFinanceiro, setSubVisaoFinanceiro] = useState('MATRIZES'); // 'MATRIZES' | 'AUDITORIA_OS'
  const [abaLateralParados, setAbaLateralParados] = useState('PLACAS'); // 'PLACAS' | 'RESUMO_CDV' | 'LAUDO'

  // --- CARIMBOS DE ÚLTIMA ATUALIZAÇÃO DOS 4 MÓDULOS ---
  const [ultimaAtuPrev, setUltimaAtuPrev] = useState(
    () => localStorage.getItem('atu_prev_sul2') || 'Aguardando planilha'
  );
  const [ultimaAtuTaco, setUltimaAtuTaco] = useState(
    () => localStorage.getItem('atu_taco_sul2') || 'Aguardando planilha'
  );
  const [ultimaAtuFin, setUltimaAtuFin] = useState(
    () => localStorage.getItem('atu_fin_sul2') || 'Aguardando planilha'
  );
  const [ultimaAtuParados, setUltimaAtuParados] = useState(
    () => localStorage.getItem('atu_parados_sul2') || 'Aguardando planilha'
  );

  // --- FILTROS DE BUSCA ---
  const [termoBusca, setTermoBusca] = useState('');
  const [termoBuscaTaco, setTermoBuscaTaco] = useState('');
  const [termoBuscaGerencial, setTermoBuscaGerencial] = useState('');
  const [termoBuscaGerencialTaco, setTermoBuscaGerencialTaco] = useState('');
  const [termoBuscaFinanceiro, setTermoBuscaFinanceiro] = useState('');
  const [termoBuscaParados, setTermoBuscaParados] = useState('');
  const [buscaPlaca3DTopo, setBuscaPlaca3DTopo] = useState('');

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

  // --- ESTADOS DE VEÍCULOS PARADOS & RESUMO (PERSISTÊNCIA DUPLA: CACHE + SUPABASE) ---
  const [veiculosParados, setVeiculosParados] = useState(() => {
    try {
      const salvo = localStorage.getItem('cache_veiculos_parados_sul2');
      return salvo ? JSON.parse(salvo) : [];
    } catch {
      return [];
    }
  });
  const [resumoDispo, setResumoDispo] = useState(() => {
    try {
      const salvo = localStorage.getItem('cache_resumo_dispo_sul2');
      return salvo ? JSON.parse(salvo) : [];
    } catch {
      return [];
    }
  });
  const [veiculo3dSelecionado, setVeiculo3dSelecionado] = useState(() => {
    try {
      const salvo = localStorage.getItem('cache_veiculos_parados_sul2');
      const lista = salvo ? JSON.parse(salvo) : [];
      return lista.length > 0 ? lista[0] : null;
    } catch {
      return null;
    }
  });
  const [filtroDispoParados, setFiltroDispoParados] = useState('TODOS');
  const [filtroCdvParados, setFiltroCdvParados] = useState('TODOS');
  const [filtroOficinaParados, setFiltroOficinaParados] = useState('TODAS');
  const [filtroZona3D, setFiltroZona3D] = useState('TODAS');
  const [substituirParados, setSubstituirParados] = useState(true);

  // --- FROTA POR CDV EXTRAÍDA DINAMICAMENTE DA PLANILHA / BANCO ---
  const [frotaPorCdv, setFrotaPorCdv] = useState(() => {
    try {
      const salvo = localStorage.getItem('frota_cdv_sul2');
      return salvo ? JSON.parse(salvo) : {};
    } catch {
      return {};
    }
  });

  // --- INTERRUPTORES DE SUBSTITUIÇÃO ---
  const [substituirFrota, setSubstituirFrota] = useState(false);
  const [substituirTaco, setSubstituirTaco] = useState(false);
  const [substituirFinanceiro, setSubstituirFinanceiro] = useState(true);

  // --- DADOS REAIS DO SUPABASE ---
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
      if (prevData && prevData.length > 0) {
        setPreventivas(prevData);
        if (prevData[0].created_at && !localStorage.getItem('atu_prev_sul2')) {
          setUltimaAtuPrev(formatarCreatedAtBanco(prevData[0].created_at));
        }
      }

      const { data: tacoData } = await supabase
        .from('tacografos')
        .select('*')
        .order('placa', { ascending: true });
      if (tacoData && tacoData.length > 0) {
        setTacografos(tacoData);
        if (tacoData[0].created_at && !localStorage.getItem('atu_taco_sul2')) {
          setUltimaAtuTaco(formatarCreatedAtBanco(tacoData[0].created_at));
        }
      }

      const { data: rateioData } = await supabase
        .from('financeiro_rateio')
        .select('*')
        .order('id', { ascending: true });
      if (rateioData && rateioData.length > 0) {
        setFinRateio(rateioData);
        if (rateioData[0].mes)
          setMesSelecionadoFinanceiro(Number(rateioData[0].mes));
        if (rateioData[0].created_at && !localStorage.getItem('atu_fin_sul2')) {
          setUltimaAtuFin(formatarCreatedAtBanco(rateioData[0].created_at));
        }
      }

      const { data: ctaData } = await supabase
        .from('financeiro_cta')
        .select('*')
        .order('valor', { ascending: false });
      setFinCta(ctaData || []);

      const { data: paradosData } = await supabase
        .from('veiculos_parados')
        .select('*')
        .order('id', { ascending: true });
      if (paradosData && paradosData.length > 0) {
        setVeiculosParados(paradosData);
        localStorage.setItem(
          'cache_veiculos_parados_sul2',
          JSON.stringify(paradosData)
        );
        setVeiculo3dSelecionado((atual) => atual || paradosData[0]);
        if (
          paradosData[0].created_at &&
          !localStorage.getItem('atu_parados_sul2')
        ) {
          setUltimaAtuParados(
            formatarCreatedAtBanco(paradosData[0].created_at)
          );
        }
      }

      const { data: resDispoData } = await supabase
        .from('resumo_disponibilidade')
        .select('*')
        .order('id', { ascending: true });
      if (resDispoData && resDispoData.length > 0) {
        setResumoDispo(resDispoData);
        localStorage.setItem(
          'cache_resumo_dispo_sul2',
          JSON.stringify(resDispoData)
        );
      }

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
      if (todasBase.length > 0) setFinBase(todasBase);
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
      setErroLogin('Login ou senha inválidos!');
    }
  };

  // --- BOTÕES MASTER DE EXCLUSÃO ---
  const limparTodaTabelaPreventivas = async () => {
    if (
      confirm(
        '⚠️ CUIDADO: Vai apagar TODOS os registos da tabela de preventivas no Supabase. Tem a certeza absoluta?'
      )
    ) {
      setCarregando(true);
      try {
        const { error } = await supabase
          .from('preventivas')
          .delete()
          .neq('placa', 'REMOVER_TUDO_MASTER');
        if (error) throw error;
        setPreventivas([]);
        alert('Tabela de preventivas completamente limpa!');
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
        '⚠️ CUIDADO: Vai apagar TODOS os registos da tabela de tacógrafos no Supabase. Tem a certeza absoluta?'
      )
    ) {
      setCarregando(true);
      try {
        const { error } = await supabase
          .from('tacografos')
          .delete()
          .neq('placa', 'REMOVER_TUDO_MASTER');
        if (error) throw error;
        setTacografos([]);
        alert('Tabela de tacógrafos completamente limpa!');
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
        '⚠️ CUIDADO: Vai apagar TODOS os registos Financeiros no Supabase. Tem a certeza absoluta?'
      )
    ) {
      setCarregando(true);
      try {
        await supabase
          .from('financeiro_rateio')
          .delete()
          .neq('conta', 'REMOVER_TUDO');
        await supabase
          .from('financeiro_cta')
          .delete()
          .neq('placa', 'REMOVER_TUDO');
        await supabase
          .from('financeiro_base')
          .delete()
          .neq('os', 'REMOVER_TUDO');
        setFinRateio([]);
        setFinCta([]);
        setFinBase([]);
        alert('Tabelas financeiras completamente limpas!');
      } catch (err) {
        alert('Erro ao limpar banco financeiro: ' + err.message);
      } finally {
        setCarregando(false);
      }
    }
  };

  const limparTodaTabelaParados = async () => {
    if (
      confirm(
        '⚠️ CUIDADO: Vai apagar TODOS os registos de Veículos Parados e Resumo de Disponibilidade. Tem a certeza?'
      )
    ) {
      setCarregando(true);
      try {
        localStorage.removeItem('cache_veiculos_parados_sul2');
        localStorage.removeItem('cache_resumo_dispo_sul2');
        setVeiculosParados([]);
        setResumoDispo([]);
        setVeiculo3dSelecionado(null);
        await supabase
          .from('veiculos_parados')
          .delete()
          .neq('placa', 'REMOVER_TUDO');
        await supabase
          .from('resumo_disponibilidade')
          .delete()
          .neq('cdv', 'REMOVER_TUDO');
        alert('Tabelas de disponibilidade completamente limpas!');
      } catch (err) {
        alert('Erro ao limpar banco de parados: ' + err.message);
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
        '⚠️ ATENÇÃO MÁXIMA: Marcou para substituir a malha ativa. Isto limpará todos os registos existentes. Deseja prosseguir?'
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
        const nomeFolhaAlvo = livro.SheetNames.includes('PREVENTIVAS SUL2')
          ? 'PREVENTIVAS SUL2'
          : livro.SheetNames[0];

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
          if (placaRaw === 'PLACA' || placaRaw === 'PLACAS') continue;

          const filialRaw = String(linha[1] || 'GERAL').toUpperCase().trim();
          const cdvRaw = String(linha[2] || '').toUpperCase().trim();
          const oficinaRaw = String(linha[4] || '').trim();
          const kmRaw = String(linha[7] || '').trim();
          const statusPlanilha = String(linha[10] || '').toUpperCase().trim();
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

        const arrayPrevs = Array.from(mapaPreventivasUnicas.values());
        setPreventivas(
          arrayPrevs.map((item, idx) => ({ id: idx + 1, ...item }))
        );

        const carimbo = gerarCarimboDataHoraAtual();
        setUltimaAtuPrev(carimbo);
        localStorage.setItem('atu_prev_sul2', carimbo);

        if (substituirFrota) {
          await supabase
            .from('preventivas')
            .delete()
            .neq('placa', 'REMOVER_TUDO_MASTER');
        }
        const { error: errIns } = await supabase
          .from('preventivas')
          .insert(arrayPrevs);
        if (!errIns) {
          fetchDados();
        }
        alert(
          `✅ Malha de Preventivas sincronizada (${arrayPrevs.length} placas)!\n🕒 Atualizado em: ${carimbo}`
        );
      } catch (err) {
        alert(err.message);
      } finally {
        setCarregando(false);
        e.target.value = '';
      }
    };
    leitor.readAsBinaryString(arquivo);
  };

  // --- IMPORTAÇÃO 2: TACÓGRAFOS ---
  const lidarComImportacaoTacografos = (e) => {
    const arquivo = e.target.files[0];
    if (!arquivo) return;
    if (
      substituirTaco &&
      !confirm(
        '⚠️ ATENÇÃO MÁXIMA: Marcou para substituir a malha ativa de Tacógrafos. Deseja prosseguir?'
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

        let abaEscolhida = livro.SheetNames[0];
        for (const nomeAba of livro.SheetNames) {
          const mTeste = XLSX.utils.sheet_to_json(livro.Sheets[nomeAba], {
            header: 1,
          });
          if (mTeste && mTeste.length > 2) {
            abaEscolhida = nomeAba;
            break;
          }
        }

        const matrizDados = XLSX.utils.sheet_to_json(
          livro.Sheets[abaEscolhida],
          { header: 1 }
        );
        const mapaTacoUnicos = new Map();

        let linhaInicial = 1;
        for (let r = 0; r < Math.min(15, matrizDados.length); r++) {
          const l = matrizDados[r];
          if (
            l &&
            String(l[0] || '')
              .toUpperCase()
              .trim()
              .includes('PLACA')
          ) {
            linhaInicial = r + 1;
            break;
          }
        }

        for (let i = linhaInicial; i < matrizDados.length; i++) {
          const linha = matrizDados[i];
          if (!linha || !linha[0]) continue;

          const placaRaw = String(linha[0]).toUpperCase().trim();
          if (placaRaw === 'PLACA' || placaRaw === 'PLACAS' || placaRaw === '')
            continue;

          const modeloRaw = String(linha[1] || 'CAMINHÃO').trim();
          const cdvRaw = String(linha[2] || 'GERAL').toUpperCase().trim();
          const filialRaw = String(linha[3] || 'GERAL').toUpperCase().trim();
          const validadeRaw = linha[4]
            ? formatarDataVersaoBancoNoTopo(linha[4])
            : null;
          const statusPlanilha = String(linha[5] || '').toUpperCase().trim();
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

        const arrayTacos = Array.from(mapaTacoUnicos.values());
        if (arrayTacos.length === 0) {
          throw new Error('Nenhuma placa encontrada na folha de cálculo selecionada!');
        }

        setTacografos(
          arrayTacos.map((item, idx) => ({ id: idx + 1, ...item }))
        );

        const primeiroComMes = arrayTacos.find((t) =>
          extrairNumeroMesPorExtenso(t.mes_referencia)
        );
        if (primeiroComMes) {
          setMesSelecionadoTacoGerencial(
            extrairNumeroMesPorExtenso(primeiroComMes.mes_referencia)
          );
        }

        const carimbo = gerarCarimboDataHoraAtual();
        setUltimaAtuTaco(carimbo);
        localStorage.setItem('atu_taco_sul2', carimbo);

        if (substituirTaco) {
          await supabase
            .from('tacografos')
            .delete()
            .neq('placa', 'REMOVER_TUDO_MASTER');
        }

        const { error: errTaco } = await supabase
          .from('tacografos')
          .insert(arrayTacos);
        if (errTaco) {
          alert(
            `✅ ${arrayTacos.length} Tacógrafos carregados no ecrã com sucesso!\n🕒 Atualizado em: ${carimbo}\n\n(Aviso Supabase: ${errTaco.message})`
          );
        } else {
          await fetchDados();
          alert(
            `✅ Painel de Tacógrafos atualizado com sucesso (${arrayTacos.length} veículos)!\n🕒 Última atualização: ${carimbo}`
          );
        }
      } catch (err) {
        alert('Erro ao carregar Tacógrafos: ' + err.message);
      } finally {
        setCarregando(false);
        e.target.value = '';
      }
    };
    leitor.readAsBinaryString(arquivo);
  };

  // --- IMPORTAÇÃO 3: FINANCEIRO ---
  const lidarComImportacaoFinanceiro = (e) => {
    const arquivo = e.target.files[0];
    if (!arquivo) return;
    if (
      substituirFinanceiro &&
      !confirm(
        '⚠️ Deseja atualizar e substituir a base financeira (DIVISÃO RATEIO + BASE SUL II + CTA + POSIÇÃO NACIONAL)?'
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

        if (!livro.SheetNames.includes('DIVISÃO RATEIO')) {
          throw new Error('O separador "DIVISÃO RATEIO" não foi encontrado!');
        }
        if (!livro.SheetNames.includes('BASE (SUL II)')) {
          throw new Error('O separador "BASE (SUL II)" não foi encontrado!');
        }

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
            localStorage.setItem('frota_cdv_sul2', JSON.stringify(novoMapaCdv));
          }
        }

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

        const matrizBase = XLSX.utils.sheet_to_json(
          livro.Sheets['BASE (SUL II)'],
          { header: 1 }
        );
        let linhaHeaderBase = 10;
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

          const osRaw =
            row[1] !== undefined && row[1] !== null ? String(row[1]).trim() : '';
          const placaRaw =
            row[3] !== undefined && row[3] !== null
              ? String(row[3]).toUpperCase().trim()
              : '';
          const valorRaw = Number(row[5]) || 0;
          const mesRaw = Number(row[13]) || null;
          const nfRaw =
            row[15] !== undefined && row[15] !== null
              ? String(row[15]).trim()
              : '';

          if (!osRaw && !placaRaw && valorRaw === 0) continue;

          listaBaseParaInserir.push({
            os: osRaw || '-',
            data_os: formatarDataVersaoBancoNoTopo(row[2]),
            placa: placaRaw || '-',
            estabelecimento: row[4] ? String(row[4]).trim() : '-',
            valor_total: valorRaw,
            ccusto: row[6] ? String(row[6]).trim() : null,
            cdv: row[7] ? String(row[7]).toUpperCase().trim() : 'GERAL',
            area: row[8] ? String(row[8]).toUpperCase().trim() : 'SUL2',
            conta_tl: row[9] ? String(row[9]).trim() : null,
            conta_correcao: row[10]
              ? String(row[10]).toUpperCase().trim()
              : 'OUTROS',
            status: row[11] ? String(row[11]).trim() : 'Pendente',
            data_faturamento: formatarDataVersaoBancoNoTopo(row[12]),
            mes: mesRaw,
            titulo:
              row[14] && !String(row[14]).includes('#')
                ? String(row[14]).trim()
                : null,
            nf: nfRaw,
          });
        }

        const carimbo = gerarCarimboDataHoraAtual();
        setUltimaAtuFin(carimbo);
        localStorage.setItem('atu_fin_sul2', carimbo);

        if (substituirFinanceiro) {
          await supabase
            .from('financeiro_rateio')
            .delete()
            .neq('conta', 'REMOVER_TUDO');
          await supabase
            .from('financeiro_cta')
            .delete()
            .neq('placa', 'REMOVER_TUDO');
          await supabase
            .from('financeiro_base')
            .delete()
            .neq('os', 'REMOVER_TUDO');
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

        const tamanhoLote = 500;
        for (let i = 0; i < listaBaseParaInserir.length; i += tamanhoLote) {
          const lote = listaBaseParaInserir.slice(i, i + tamanhoLote);
          const { error: errB } = await supabase
            .from('financeiro_base')
            .insert(lote);
          if (errB) throw errB;
        }

        setMesSelecionadoFinanceiro(mesRateioPlanilha);
        setMostrarGavetaUploadFin(false);
        alert(
          `✅ Sincronização Financeira Concluída!\n🕒 Atualizado em: ${carimbo}\n- Divisão Rateio (Mês ${mesRateioPlanilha}): ${listaRateioParaInserir.length} linhas\n- Alertas CTA: ${listaCtaParaInserir.length} veículos críticos\n- Base SUL II: ${listaBaseParaInserir.length} registos`
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

  // --- IMPORTAÇÃO 4: CONTROLE DE VEÍCULOS PARADOS (SALVA CACHE + SUPABASE SEM SUMIR!) ---
  const lidarComImportacaoVeiculosParados = (e) => {
    const arquivo = e.target.files[0];
    if (!arquivo) return;
    setCarregando(true);
    const leitor = new FileReader();
    leitor.onload = async (evento) => {
      try {
        const dadosBinarios = evento.target.result;
        const livro = XLSX.read(dadosBinarios, { type: 'binary' });

        if (!livro.SheetNames.includes('Em Oficina')) {
          throw new Error('O separador "Em Oficina" não foi encontrado!');
        }

        const matrizOficina = XLSX.utils.sheet_to_json(
          livro.Sheets['Em Oficina'],
          { header: 1 }
        );
        const listaParados = [];
        for (let i = 1; i < matrizOficina.length; i++) {
          const row = matrizOficina[i];
          if (!row || !row[0]) continue;
          const placaRaw = String(row[0]).toUpperCase().trim();
          if (placaRaw === 'PLACAS' || placaRaw === 'PLACA' || placaRaw === '')
            continue;

          listaParados.push({
            placa: placaRaw,
            filial: String(row[1] || 'GERAL').toUpperCase().trim(),
            oficina: String(row[2] || 'SEM OFICINA').toUpperCase().trim(),
            cdv: String(row[3] || 'GERAL').toUpperCase().trim(),
            status: String(row[4] || 'EM DIAGNÓSTICO').trim(),
            os: row[5] ? String(row[5]).trim() : '--',
            valor: limparNumeroMoeda(row[6]),
            previsao: formatarDataVersaoBancoNoTopo(row[7]),
            disponibilidade: String(row[8] || 'Em oficina').trim(),
          });
        }

        let listaResumo = [];
        if (livro.SheetNames.includes('Resumo')) {
          const matrizResumo = XLSX.utils.sheet_to_json(
            livro.Sheets['Resumo'],
            { header: 1 }
          );
          for (let i = 0; i < matrizResumo.length; i++) {
            const row = matrizResumo[i];
            if (!row || !row[0]) continue;
            const nomeCdv = String(row[0])
              .replace("'", '')
              .toUpperCase()
              .trim();
            if (nomeCdv.startsWith('CDV ')) {
              const disp = Number(row[1]) || 0;
              const ofic = Number(row[2]) || 0;
              const lib = Number(row[3]) || 0;
              const tot = Number(row[4]) || disp + ofic + lib;
              const pct =
                row[5] !== undefined && row[5] !== null && !isNaN(Number(row[5]))
                  ? Number(row[5]) <= 1
                    ? Number(row[5]) * 100
                    : Number(row[5])
                  : tot > 0
                  ? (disp / tot) * 100
                  : 100;
              listaResumo.push({
                cdv: nomeCdv,
                disponivel: disp,
                em_oficina: ofic,
                para_liberacao: lib,
                total_geral: tot,
                pct_dispo: Number(pct.toFixed(2)),
              });
            }
          }
        }

        if (
          listaResumo.length === 0 &&
          livro.SheetNames.includes('Posição Nacional')
        ) {
          const matrizPN = XLSX.utils.sheet_to_json(
            livro.Sheets['Posição Nacional'],
            { header: 1 }
          );
          const headerPN = (matrizPN[0] || []).map((h) =>
            String(h || '').toUpperCase().trim()
          );
          const idxCdv = headerPN.findIndex((h) => h.includes('CDV'));
          const idxStatusOfic = headerPN.findIndex((h) =>
            h.includes('STATUS OFICINA')
          );

          if (idxCdv !== -1) {
            const agrupado = {};
            for (let r = 1; r < matrizPN.length; r++) {
              const row = matrizPN[r];
              if (!row || !row[idxCdv]) continue;
              const c = String(row[idxCdv]).toUpperCase().trim();
              if (!c.startsWith('CDV')) continue;
              if (!agrupado[c]) {
                agrupado[c] = {
                  disponivel: 0,
                  em_oficina: 0,
                  para_liberacao: 0,
                  total_geral: 0,
                };
              }
              const st =
                idxStatusOfic !== -1
                  ? String(row[idxStatusOfic] || '').toUpperCase()
                  : 'DISPON';
              agrupado[c].total_geral += 1;
              if (st.includes('OFICINA')) agrupado[c].em_oficina += 1;
              else if (st.includes('LIBERA')) agrupado[c].para_liberacao += 1;
              else agrupado[c].disponivel += 1;
            }
            listaResumo = Object.keys(agrupado)
              .sort()
              .map((c) => ({
                cdv: c,
                ...agrupado[c],
                pct_dispo: Number(
                  (
                    (agrupado[c].disponivel / (agrupado[c].total_geral || 1)) *
                    100
                  ).toFixed(2)
                ),
              }));
          }
        }

        const paradosComId = listaParados.map((item, idx) => ({
          id: idx + 1,
          ...item,
        }));

        // 1. Persiste imediatamente em estado + localStorage (sobrevive ao F5!)
        setVeiculosParados(paradosComId);
        localStorage.setItem(
          'cache_veiculos_parados_sul2',
          JSON.stringify(paradosComId)
        );
        if (paradosComId.length > 0) {
          setVeiculo3dSelecionado(paradosComId[0]);
        }

        setResumoDispo(listaResumo);
        localStorage.setItem(
          'cache_resumo_dispo_sul2',
          JSON.stringify(listaResumo)
        );

        const carimbo = gerarCarimboDataHoraAtual();
        setUltimaAtuParados(carimbo);
        localStorage.setItem('atu_parados_sul2', carimbo);

        // 2. Persiste no Supabase em segundo plano
        if (substituirParados) {
          await supabase
            .from('veiculos_parados')
            .delete()
            .neq('placa', 'REMOVER_TUDO');
          await supabase
            .from('resumo_disponibilidade')
            .delete()
            .neq('cdv', 'REMOVER_TUDO');
        }
        if (listaParados.length > 0) {
          const { error: errP } = await supabase
            .from('veiculos_parados')
            .insert(listaParados);
          if (errP) console.warn('Aviso veiculos_parados:', errP.message);
        }
        if (listaResumo.length > 0) {
          const { error: errR } = await supabase
            .from('resumo_disponibilidade')
            .insert(listaResumo);
          if (errR) console.warn('Aviso resumo_disponibilidade:', errR.message);
        }

        alert(
          `✅ Folha de cálculo de Veículos Parados sincronizada!\n🕒 Atualizado em: ${carimbo}\n- ${listaParados.length} veículos mapeados na Vitrine 3D/2D\n- ${listaResumo.length} CDVs analisados`
        );
      } catch (err) {
        alert('Erro ao importar folha de cálculo: ' + err.message);
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
      String(p.placa || '').includes(termoBuscaGerencial.toUpperCase()) ||
      String(p.cdv || '').includes(termoBuscaGerencial.toUpperCase());
    return passaCard && passaBusca;
  });

  // ============================================================================
  // --- ABA GERENCIAL 2: REGRAS DE TACÓGRAFOS ---
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
      String(t.placa || '').includes(termoBuscaGerencialTaco.toUpperCase()) ||
      String(t.cdv || '').includes(termoBuscaGerencialTaco.toUpperCase());
    return passaCard && passaBusca;
  });

  // ============================================================================
  // --- ABA GERENCIAL 3: REGRAS DE STATUS FINANCEIRO ---
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

      const aop = regAop.reduce(
        (acc, r) => acc + (Number(r.valor_total) || 0),
        0
      );
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

  const mapaFrotaDinamico = { ...frotaPorCdv };
  if (Object.keys(mapaFrotaDinamico).length === 0 && resumoDispo.length > 0) {
    resumoDispo.forEach((r) => {
      if (r.cdv && r.total_geral > 0) {
        mapaFrotaDinamico[String(r.cdv).toUpperCase().trim()] = Number(
          r.total_geral
        );
      }
    });
  }
  if (Object.keys(mapaFrotaDinamico).length === 0 && osReaisBase.length > 0) {
    const setsPorCdv = {};
    osReaisBase.forEach((b) => {
      if (b.area === 'SUL2' && b.cdv && b.placa && b.placa !== '-') {
        const c = String(b.cdv).toUpperCase().trim();
        if (!setsPorCdv[c]) setsPorCdv[c] = new Set();
        setsPorCdv[c].add(b.placa);
      }
    });
    Object.keys(setsPorCdv).forEach((c) => {
      mapaFrotaDinamico[c] = setsPorCdv[c].size;
    });
  }

  const totalPlacasSul2 =
    Object.values(mapaFrotaDinamico).reduce((acc, qtd) => acc + qtd, 0) || 1;

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
      ...Object.keys(mapaFrotaDinamico),
      ...osDoMesSelecionado.map((b) =>
        String(b.cdv || 'GERAL')
          .toUpperCase()
          .trim()
      ),
    ])
  );

  const analisePorCdv = todosCdvsMes
    .map((cdvNome) => {
      const qtdPlacas = mapaFrotaDinamico[cdvNome] || 0;
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

      const pctDoOrcamento =
        orcamentoCdv > 0 ? (gastoReal / orcamentoCdv) * 100 : 0;
      const pctEstouro = orcamentoCdv > 0 ? pctDoOrcamento - 100 : 0;
      const saldoCdv = orcamentoCdv - gastoReal;
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

  // ============================================================================
  // --- ABA GERENCIAL 4: REGRAS DE VEÍCULOS PARADOS & RESUMO ---
  // ============================================================================
  const resumoDispoCalculado =
    resumoDispo.length > 0
      ? resumoDispo
      : Array.from(
          new Set(
            veiculosParados
              .map((v) => String(v.cdv || '').toUpperCase().trim())
              .filter(Boolean)
          )
        )
          .sort()
          .map((cdvNome) => {
            const carrosCdv = veiculosParados.filter(
              (v) => String(v.cdv || '').toUpperCase().trim() === cdvNome
            );
            const disp = carrosCdv.filter((v) =>
              String(v.disponibilidade || '').toUpperCase().includes('DISPON')
            ).length;
            const lib = carrosCdv.filter((v) =>
              String(v.disponibilidade || '').toUpperCase().includes('LIBERA')
            ).length;
            const ofic = carrosCdv.length - disp - lib;
            const tot = carrosCdv.length || 1;
            return {
              cdv: cdvNome,
              disponivel: disp,
              em_oficina: ofic,
              para_liberacao: lib,
              total_geral: carrosCdv.length,
              pct_dispo: Number(((disp / tot) * 100).toFixed(2)),
            };
          });

  const totalFrotaResumo = resumoDispoCalculado.reduce(
    (acc, r) => acc + (Number(r.total_geral) || 0),
    0
  );
  const totalDisponivelResumo = resumoDispoCalculado.reduce(
    (acc, r) => acc + (Number(r.disponivel) || 0),
    0
  );
  const totalOficinaResumo = resumoDispoCalculado.reduce(
    (acc, r) => acc + (Number(r.em_oficina) || 0),
    0
  );
  const totalLiberacaoResumo = resumoDispoCalculado.reduce(
    (acc, r) => acc + (Number(r.para_liberacao) || 0),
    0
  );
  const pctDispoGeral =
    totalFrotaResumo > 0
      ? ((totalDisponivelResumo / totalFrotaResumo) * 100).toFixed(1)
      : '0.0';

  const oficinasUnicasParados = Array.from(
    new Set(
      veiculosParados
        .map((v) => String(v.oficina || '').toUpperCase().trim())
        .filter((o) => o && o !== '--')
    )
  ).sort();

  const zona3dAtiva = veiculo3dSelecionado
    ? identificarZonaDefeito3D(veiculo3dSelecionado.status)
    : identificarZonaDefeito3D('PREVENTIVA');

  const listaVeiculosParadosFiltrada = veiculosParados
    .filter((v) => {
      if (filtroDispoParados === 'TODOS') return true;
      return String(v.disponibilidade || '')
        .toUpperCase()
        .includes(filtroDispoParados);
    })
    .filter((v) => {
      if (filtroCdvParados === 'TODOS') return true;
      return String(v.cdv || '').toUpperCase() === filtroCdvParados;
    })
    .filter((v) => {
      if (filtroOficinaParados === 'TODAS') return true;
      return (
        String(v.oficina || '').toUpperCase().trim() === filtroOficinaParados
      );
    })
    .filter((v) => {
      if (filtroZona3D === 'TODAS') return true;
      return identificarZonaDefeito3D(v.status).id === filtroZona3D;
    })
    .filter((v) => {
      if (!termoBuscaParados) return true;
      const t = termoBuscaParados.toUpperCase();
      return (
        String(v.placa || '').toUpperCase().includes(t) ||
        String(v.cdv || '').toUpperCase().includes(t) ||
        String(v.oficina || '').toUpperCase().includes(t) ||
        String(v.status || '').toUpperCase().includes(t) ||
        String(v.filial || '').toUpperCase().includes(t)
      );
    });

  const executarBuscaRapidaPlaca3D = (valorDigitado) => {
    setBuscaPlaca3DTopo(valorDigitado);
    const limpo = String(valorDigitado || '').toUpperCase().trim();
    if (!limpo) return;
    const encontrado = veiculosParados.find((v) =>
      String(v.placa || '').toUpperCase().includes(limpo)
    );
    if (encontrado) {
      setVeiculo3dSelecionado(encontrado);
    }
  };

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
              SUL2
            </p>
          </div>
          <form onSubmit={lidarComLogin} className="space-y-5">
            <input
              type="text"
              placeholder="Login (Ex: SUL2CO)"
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
              Acessar
            </button>
          </form>
        </div>
        <div className="mt-8 text-center">
          <p className="text-[11px] font-bold text-gray-500 uppercase tracking-widest">
            Operado por{' '}
            <span className="text-orange-400 font-black">Fidel Barbosa(Opentech)</span>
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className="flex h-screen font-sans bg-[#0B1329] text-gray-100 relative overflow-hidden">
      {/* ============================================================================ */}
      {/* --- BARRA LATERAL RETRÁTIL INTELIGENTE (MINI-DOCK / EXPANSÍVEL) --- */}
      {/* ============================================================================ */}
      <aside
        className={`${
          sidebarAberta ? 'w-60' : 'w-16'
        } border-r border-gray-800 flex flex-col bg-[#1C2541]/95 transition-all duration-300 shrink-0 z-40 select-none`}
      >
        <div className="h-14 flex items-center justify-between px-3 border-b border-gray-800">
          <div className="flex items-center overflow-hidden">
            <div
              onClick={() => setSidebarAberta(!sidebarAberta)}
              className="w-9 h-9 bg-gradient-to-br from-[#0000FF] to-[#FFA500] rounded-xl flex items-center justify-center text-white font-black text-xs shrink-0 cursor-pointer shadow-md"
              title="Expandir / Recolher Menu"
            >
              PF
            </div>
            {sidebarAberta && (
              <span className="font-black text-base ml-2.5 tracking-tight text-white truncate">
                Painel Frota
              </span>
            )}
          </div>
          <button
            onClick={() => setSidebarAberta(!sidebarAberta)}
            className="p-1.5 rounded-lg bg-[#0B1329] text-gray-400 hover:text-white border border-gray-700/80 transition-all"
            title={sidebarAberta ? 'Recolher Barra Lateral' : 'Expandir Barra Lateral'}
          >
            {sidebarAberta ? <ChevronLeft size={14} /> : <ChevronRight size={14} />}
          </button>
        </div>

        {sidebarAberta && (
          <div className="p-3 border-b border-gray-800 animate-fadeIn">
            <label className="text-[9px] font-black text-gray-400 uppercase tracking-wider">
              Unidade / Filial
            </label>
            <select
              className="w-full bg-[#0B1329] border border-gray-700 text-xs font-bold p-1.5 mt-1 rounded-lg uppercase text-white outline-none"
              value={filialAtual}
              onChange={(e) => setFilialAtual(e.target.value)}
            >
              <option value="Todas">Visão Geral (Todas)</option>
              {Array.from(
                new Set([
                  ...preventivas.map((p) => p.regional),
                  ...tacografos.map((t) => t.regional),
                ])
              )
                .filter(Boolean)
                .map((f) => (
                  <option key={f} value={f}>
                    {f}
                  </option>
                ))}
            </select>
          </div>
        )}

        <nav className="p-2 space-y-1.5 flex-1">
          <button
            onClick={() => setActiveTab('frota')}
            title="Fila de Oficina"
            className={`w-full flex items-center ${
              sidebarAberta ? 'justify-start gap-3 px-3' : 'justify-center px-0'
            } py-2.5 rounded-xl text-xs font-bold transition-all ${
              activeTab === 'frota'
                ? 'bg-blue-500/20 text-blue-400 border border-blue-500/30'
                : 'text-gray-400 hover:bg-gray-800/50'
            }`}
          >
            <CalendarDays size={18} className="shrink-0" />
            {sidebarAberta && <span className="truncate">Fila de Oficina</span>}
          </button>

          <button
            onClick={() => setActiveTab('historico_mensal')}
            title="Status Preventivas"
            className={`w-full flex items-center ${
              sidebarAberta ? 'justify-start gap-3 px-3' : 'justify-center px-0'
            } py-2.5 rounded-xl text-xs font-bold transition-all ${
              activeTab === 'historico_mensal'
                ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/30'
                : 'text-gray-400 hover:bg-gray-800/50'
            }`}
          >
            <History size={18} className="shrink-0" />
            {sidebarAberta && <span className="truncate">Status Preventivas</span>}
          </button>

          <button
            onClick={() => setActiveTab('status_tacografos')}
            title="Status Tacógrafos"
            className={`w-full flex items-center ${
              sidebarAberta ? 'justify-start gap-3 px-3' : 'justify-center px-0'
            } py-2.5 rounded-xl text-xs font-bold transition-all ${
              activeTab === 'status_tacografos'
                ? 'bg-amber-500/20 text-amber-400 border border-amber-500/30'
                : 'text-gray-400 hover:bg-gray-800/50'
            }`}
          >
            <Clock size={18} className="shrink-0" />
            {sidebarAberta && <span className="truncate">Status Tacógrafos</span>}
          </button>

          <button
            onClick={() => setActiveTab('status_financeiro')}
            title="Status Financeiro & Rateio"
            className={`w-full flex items-center ${
              sidebarAberta ? 'justify-between px-3' : 'justify-center px-0'
            } py-2.5 rounded-xl text-xs font-bold transition-all relative ${
              activeTab === 'status_financeiro'
                ? 'bg-rose-500/20 text-rose-400 border border-rose-500/30'
                : 'text-gray-400 hover:bg-gray-800/50'
            }`}
          >
            <div className="flex items-center gap-3">
              <DollarSign size={18} className="text-rose-400 shrink-0" />
              {sidebarAberta && <span className="truncate">Status Financeiro</span>}
            </div>
            {finCta.length > 0 && sidebarAberta && (
              <span className="bg-red-600 text-white text-[9px] font-black px-1.5 py-0.5 rounded-full animate-pulse">
                {finCta.length}
              </span>
            )}
            {finCta.length > 0 && !sidebarAberta && (
              <span className="w-2 h-2 rounded-full bg-red-500 absolute top-1.5 right-1.5 animate-ping" />
            )}
          </button>

          <button
            onClick={() => setActiveTab('veiculos_parados')}
            title="Veículos Parados (3D/2D)"
            className={`w-full flex items-center ${
              sidebarAberta ? 'justify-between px-3' : 'justify-center px-0'
            } py-2.5 rounded-xl text-xs font-bold transition-all ${
              activeTab === 'veiculos_parados'
                ? 'bg-cyan-500/20 text-cyan-400 border border-cyan-500/30'
                : 'text-gray-400 hover:bg-gray-800/50'
            }`}
          >
            <div className="flex items-center gap-3">
              <Truck size={18} className="text-cyan-400 shrink-0" />
              {sidebarAberta && <span className="truncate">Veículos Parados</span>}
            </div>
            {sidebarAberta && (
              <span className="bg-yellow-400 text-gray-950 text-[8px] font-black px-1.5 py-0.5 rounded-full">
                3D/2D
              </span>
            )}
          </button>

          <a
            href={LINK_POWER_BI}
            target="_blank"
            rel="noopener noreferrer"
            title="BI de Multas"
            className={`w-full flex items-center ${
              sidebarAberta ? 'justify-start gap-3 px-3' : 'justify-center px-0'
            } py-2.5 rounded-xl text-gray-400 hover:bg-emerald-500/10 hover:text-emerald-400 transition-all`}
          >
            <BarChart3 size={18} className="text-emerald-500 shrink-0" />
            {sidebarAberta && (
              <span className="flex items-center gap-1 font-bold text-xs truncate">
                BI de Multas <ExternalLink size={11} className="opacity-60" />
              </span>
            )}
          </a>
        </nav>

        <div className="p-2.5 border-t border-gray-800 text-center">
          <button
            onClick={fetchDados}
            title="Sincronizar Banco Agora"
            className="w-full py-1.5 rounded-lg bg-[#0B1329] hover:bg-gray-800 border border-gray-700 text-gray-300 flex items-center justify-center gap-1.5 text-[10px] font-bold"
          >
            <RefreshCw
              size={12}
              className={rotacionar ? 'animate-spin text-[#FFA500]' : ''}
            />
            {sidebarAberta && <span>Sincronizar</span>}
          </button>
        </div>
      </aside>

      {/* ============================================================================ */}
      {/* --- ÁREA PRINCIPAL WIDESCREEN (OCUPA 100% DA LARGURA COM PADDING COMPACTO) --- */}
      {/* ============================================================================ */}
      <main className="flex-1 overflow-y-auto px-5 py-4 relative">
        {/* ============================================================================ */}
        {/* --- TELA 1: FILA DE OFICINA --- */}
        {/* ============================================================================ */}
        {activeTab === 'frota' && (
          <div className="w-full space-y-4 animate-fadeIn">
            <div className="flex flex-wrap justify-between items-center border-b border-gray-800 pb-3 gap-2">
              <div>
                <h2 className="text-lg font-black text-white">
                  Painel Operacional Logístico
                </h2>
                <p className="text-[11px] text-gray-400">
                  🕒 Preventivas: <span className="text-orange-400 font-bold">{ultimaAtuPrev}</span> • Tacógrafos:{' '}
                  <span className="text-blue-400 font-bold">{ultimaAtuTaco}</span>
                </p>
              </div>
              <div className="flex p-1 rounded-xl border border-gray-800 bg-[#111A33]">
                <button
                  onClick={() => setSubAbaFrota('preventivas')}
                  className={`px-3 py-1.5 rounded-lg text-xs font-bold ${
                    subAbaFrota === 'preventivas'
                      ? 'bg-orange-500/15 text-[#FFA500]'
                      : 'text-gray-400'
                  }`}
                >
                  🛠️ Preventivas ({listagemPrevsPendentes.length})
                </button>
                <button
                  onClick={() => setSubAbaFrota('tacografos')}
                  className={`px-3 py-1.5 rounded-lg text-xs font-bold ${
                    subAbaFrota === 'tacografos'
                      ? 'bg-blue-500/15 text-blue-400'
                      : 'text-gray-400'
                  }`}
                >
                  ⏱️ Tacógrafos ({listagemTacosPendentes.length})
                </button>
              </div>
            </div>

            {subAbaFrota === 'preventivas' && (
              <div className="space-y-3">
                {/* BARRA COMPACTA DE IMPORTAÇÃO + BUSCA NA MESMA LINHA */}
                <div className="p-3 bg-[#1C2541] border border-gray-800 rounded-xl flex flex-wrap justify-between items-center gap-3">
                  <div className="flex items-center gap-2 flex-1 min-w-[240px] bg-[#0B1329] px-3 py-2 rounded-lg border border-gray-700">
                    <Search size={15} className="text-gray-400 shrink-0" />
                    <input
                      type="text"
                      value={termoBusca}
                      onChange={(e) => setTermoBusca(e.target.value)}
                      placeholder="Pesquisar preventiva ativa por Placa ou CDV..."
                      className="bg-transparent w-full outline-none text-xs text-white"
                    />
                  </div>
                  <div className="flex flex-wrap items-center gap-2.5">
                    <label className="flex items-center gap-1.5 cursor-pointer text-xs text-gray-300 font-bold">
                      <input
                        type="checkbox"
                        checked={substituirFrota}
                        onChange={(e) => setSubstituirFrota(e.target.checked)}
                        className="accent-orange-500"
                      />
                      Substituir atual
                    </label>
                    <button
                      onClick={limparTodaTabelaPreventivas}
                      className="bg-red-600/10 text-red-400 hover:bg-red-600 hover:text-white border border-red-500/20 text-xs font-black px-3 py-2 rounded-lg transition-all flex items-center gap-1"
                    >
                      <Trash2 size={13} /> Limpar
                    </button>
                    <label className="bg-orange-500 hover:bg-orange-600 text-white font-black text-xs px-3.5 py-2 rounded-lg cursor-pointer flex items-center gap-1.5 transition-all">
                      <Upload size={13} />
                      Carregar Excel Preventivas
                      <input
                        type="file"
                        accept=".xlsx, .xls"
                        onChange={lidarComImportacaoPreventivas}
                        className="hidden"
                      />
                    </label>
                  </div>
                </div>

                <div className="rounded-xl border border-gray-800 bg-[#1C2541] overflow-hidden">
                  <div className="max-h-[calc(100vh-190px)] overflow-y-auto">
                    <table className="w-full text-left text-xs table-auto">
                      <thead className="sticky top-0 bg-[#111A33] font-bold text-gray-400 uppercase text-[10px] border-b border-gray-800">
                        <tr>
                          <th className="p-3">Placa</th>
                          <th className="p-3">Filial</th>
                          <th className="p-3">CDV</th>
                          <th className="p-3">KM</th>
                          <th className="p-3 text-orange-400">Mês Meta</th>
                          <th className="p-3">Oficina</th>
                          <th className="p-3">Obs 1</th>
                          <th className="p-3">Obs 2</th>
                          <th className="p-3 text-right">Ação</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gray-800">
                        {listagemPrevsPendentes
                          .filter(
                            (p) =>
                              !termoBusca ||
                              String(p.placa || '').includes(
                                termoBusca.toUpperCase()
                              ) ||
                              String(p.cdv || '').includes(
                                termoBusca.toUpperCase()
                              )
                          )
                          .map((p, idx) => (
                            <tr
                              key={p.id || idx}
                              className="hover:bg-gray-800/20"
                            >
                              <td className="p-3 font-mono font-bold">
                                {p.placa}
                              </td>
                              <td className="p-3 text-blue-400 uppercase">
                                {p.regional}
                              </td>
                              <td className="p-3 uppercase">{p.cdv}</td>
                              <td className="p-3">{p.km || '--'}</td>
                              <td className="p-3 text-orange-400 font-black uppercase">
                                {p.mes_referencia || 'SEM MÊS'}
                              </td>
                              <td className="p-3">{p.oficina || '--'}</td>
                              <td className="p-3 text-gray-400 italic">
                                {p.observacao || '--'}
                              </td>
                              <td className="p-3 text-gray-400 italic">
                                {p.observacao2 || '--'}
                              </td>
                              <td className="p-3 text-right">
                                <button
                                  onClick={() => concluirPreventiva(p.id)}
                                  className="bg-emerald-600/15 text-emerald-400 hover:bg-emerald-600 hover:text-white text-[10px] px-2.5 py-1 rounded font-bold transition-all"
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
              <div className="space-y-3">
                <div className="p-3 bg-[#1C2541] border border-gray-800 rounded-xl flex flex-wrap justify-between items-center gap-3">
                  <div className="flex items-center gap-2 flex-1 min-w-[240px] bg-[#0B1329] px-3 py-2 rounded-lg border border-gray-700">
                    <Search size={15} className="text-gray-400 shrink-0" />
                    <input
                      type="text"
                      value={termoBuscaTaco}
                      onChange={(e) => setTermoBuscaTaco(e.target.value)}
                      placeholder="Pesquisar laudo ativo por Placa ou CDV..."
                      className="bg-transparent w-full outline-none text-xs text-white"
                    />
                  </div>
                  <div className="flex flex-wrap items-center gap-2.5">
                    <label className="flex items-center gap-1.5 cursor-pointer text-xs text-gray-300 font-bold">
                      <input
                        type="checkbox"
                        checked={substituirTaco}
                        onChange={(e) => setSubstituirTaco(e.target.checked)}
                        className="accent-blue-500"
                      />
                      Substituir atual
                    </label>
                    <button
                      onClick={limparTodaTabelaTacografos}
                      className="bg-red-600/10 text-red-400 hover:bg-red-600 hover:text-white border border-red-500/20 text-xs font-black px-3 py-2 rounded-lg transition-all flex items-center gap-1"
                    >
                      <Trash2 size={13} /> Limpar
                    </button>
                    <label className="bg-blue-600 hover:bg-blue-500 text-white font-black text-xs px-3.5 py-2 rounded-lg cursor-pointer flex items-center gap-1.5 transition-all">
                      <Upload size={13} />
                      Carregar Excel Tacógrafos
                      <input
                        type="file"
                        accept=".xlsx, .xls"
                        onChange={lidarComImportacaoTacografos}
                        className="hidden"
                      />
                    </label>
                  </div>
                </div>

                <div className="rounded-xl border border-gray-800 bg-[#1C2541] overflow-hidden">
                  <div className="max-h-[calc(100vh-190px)] overflow-y-auto">
                    <table className="w-full text-left text-xs table-auto">
                      <thead className="sticky top-0 bg-[#111A33] font-bold text-gray-400 uppercase text-[10px] border-b border-gray-800">
                        <tr>
                          <th className="p-3">Placa</th>
                          <th className="p-3">Modelo</th>
                          <th className="p-3">CDV Logístico</th>
                          <th className="p-3 font-bold text-blue-400">
                            Filial / Sede
                          </th>
                          <th className="p-3">Mês / Ano Meta</th>
                          <th className="p-3">Validade Legal</th>
                          <th className="p-3 text-right">Ação</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gray-800">
                        {listagemTacosPendentes
                          .filter(
                            (t) =>
                              !termoBuscaTaco ||
                              String(t.placa || '').includes(
                                termoBuscaTaco.toUpperCase()
                              ) ||
                              String(t.cdv || '').includes(
                                termoBuscaTaco.toUpperCase()
                              )
                          )
                          .map((t, idx) => (
                            <tr
                              key={t.id || idx}
                              className="hover:bg-gray-800/20"
                            >
                              <td className="p-3 font-mono font-bold">
                                {t.placa}
                              </td>
                              <td className="p-3 text-gray-400 uppercase">
                                {t.modelo || 'CAMINHÃO'}
                              </td>
                              <td className="p-3 uppercase">{t.cdv}</td>
                              <td className="p-3 uppercase text-blue-400">
                                {t.regional}
                              </td>
                              <td className="p-3 text-blue-400 font-black uppercase">
                                {t.mes_referencia || 'SEM MÊS'}
                              </td>
                              <td className="p-3 font-bold">
                                {t.validade
                                  ? new Date(
                                      t.validade + 'T00:00:00'
                                    ).toLocaleDateString('pt-BR')
                                  : '--'}
                              </td>
                              <td className="p-3 text-right">
                                <button
                                  onClick={() => concluirTacografo(t.id)}
                                  className="bg-emerald-600/15 text-emerald-400 hover:bg-emerald-600 hover:text-white text-[10px] px-2.5 py-1 rounded font-bold transition-all"
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

        {/* ============================================================================ */}
        {/* --- TELA 2: STATUS PREVENTIVAS (SPLIT-SCREEN: GRÁFICO ESQ + TABELA DIR) --- */}
        {/* ============================================================================ */}
        {activeTab === 'historico_mensal' && (
          <div className="w-full animate-fadeIn space-y-3">
            <div className="flex flex-wrap justify-between items-center gap-2">
              <div className="flex items-center gap-3">
                <h2 className="text-xl font-black tracking-tight text-white">
                  Cronograma de Preventivas
                </h2>
                <span className="text-[10px] bg-orange-500/15 text-orange-300 border border-orange-500/30 px-2.5 py-0.5 rounded-full font-bold">
                  🕒 {ultimaAtuPrev}
                </span>
              </div>
              <div className="flex items-center gap-2 px-3 py-1.5 rounded-xl border border-gray-800 bg-[#1C2541] text-xs font-bold text-white">
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

            {/* FAIXA COMPACTA DE 5 CARDS */}
            <div className="grid grid-cols-2 sm:grid-cols-5 gap-3">
              {[
                {
                  id: 'EM DIA',
                  label: 'EM DIA',
                  pct: `${pctEmDia}%`,
                  qtd: qtdEmDiaComData,
                  grad: 'from-green-500 to-emerald-600 text-white',
                  ring: 'ring-green-400',
                },
                {
                  id: 'EM EXECUÇÃO',
                  label: 'EM EXECUÇÃO',
                  pct: `${pctExecucao}%`,
                  qtd: qtdExecucaoComData,
                  grad: 'from-orange-500 to-amber-600 text-white',
                  ring: 'ring-orange-400',
                },
                {
                  id: 'A VENCER',
                  label: 'A VENCER',
                  pct: `${pctAVencer}%`,
                  qtd: qtdAVencerComData,
                  grad: 'from-yellow-400 to-yellow-500 text-gray-900',
                  ring: 'ring-yellow-300',
                },
                {
                  id: 'VENCIDAS',
                  label: 'VENCIDAS',
                  pct: `${pctVencidas}%`,
                  qtd: qtdVencidasComData,
                  grad: 'from-red-500 to-rose-600 text-white',
                  ring: 'ring-red-400',
                },
                {
                  id: 'SEM DATA',
                  label: 'SEM MÊS',
                  pct: 'PREENCHER',
                  qtd: preventivasSemDataFilaGeral.length,
                  grad: 'from-gray-600 to-gray-700 text-white',
                  ring: 'ring-gray-400',
                },
              ].map((c) => (
                <div
                  key={c.id}
                  onClick={() =>
                    setStatusFiltroCard(
                      statusFiltroCard === c.id ? 'TODOS' : c.id
                    )
                  }
                  className={`p-3 rounded-xl shadow-sm border border-gray-800 cursor-pointer transition-all ${
                    statusFiltroCard === c.id ? `ring-2 ${c.ring}` : ''
                  } bg-gradient-to-br ${c.grad}`}
                >
                  <div className="flex justify-between items-center">
                    <span className="text-[10px] font-black">{c.label}</span>
                    <span className="text-[9px] bg-black/20 px-1.5 py-0.5 rounded font-bold">
                      {c.pct}
                    </span>
                  </div>
                  <h3 className="text-2xl font-black mt-1">
                    {String(c.qtd).padStart(2, '0')}{' '}
                    <span className="text-[10px] font-normal opacity-80">
                      Placas
                    </span>
                  </h3>
                </div>
              ))}
            </div>

            {/* SPLIT-SCREEN LADO A LADO: GRÁFICO (5 COLUNAS) + TABELA (7 COLUNAS) */}
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-4 items-stretch">
              {/* COLUNA ESQUERDA: GRÁFICO ANUAL */}
              <div className="lg:col-span-5 p-4 rounded-2xl border border-gray-800/80 bg-[#131B35] shadow-xl flex flex-col justify-between">
                <div>
                  <h3 className="text-sm font-black text-white uppercase flex items-center justify-between">
                    <span>📊 Frota por Mês (2026)</span>
                    <span className="text-[10px] text-orange-400 font-bold">
                      Clique na barra p/ filtrar
                    </span>
                  </h3>
                </div>

                <div className="relative pt-6 pb-2 my-auto">
                  <div className="h-56 flex items-end justify-between gap-1 px-1 relative z-10 border-b border-gray-700 pb-2">
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
                            className={`text-[9px] font-black mb-1 ${
                              estaSelecionado
                                ? 'text-orange-400 scale-110'
                                : 'text-gray-400'
                            }`}
                          >
                            {item.qtd > 0 ? item.qtd : ''}
                          </span>
                          <div
                            className={`w-full max-w-[26px] rounded-t-md transition-all ${
                              estaSelecionado
                                ? 'bg-gradient-to-t from-rose-600 via-orange-500 to-amber-300 ring-2 ring-orange-400'
                                : 'bg-gradient-to-t from-rose-900/60 via-orange-600/70 to-amber-500/80 group-hover:brightness-125'
                            }`}
                            style={{ height: `${alturaPct}%` }}
                          ></div>
                          <span
                            className={`mt-2 text-[9px] font-bold uppercase ${
                              estaSelecionado
                                ? 'text-orange-400 font-black'
                                : 'text-gray-400'
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
                      <span className="text-[9px] font-black text-gray-400 mb-1">
                        {preventivasSemDataFilaGeral.length || ''}
                      </span>
                      <div
                        className="w-full max-w-[26px] rounded-t-md bg-gradient-to-t from-gray-700 to-gray-400"
                        style={{
                          height: `${Math.max(
                            (preventivasSemDataFilaGeral.length /
                              maxQtdMesesPrev) *
                              100,
                            4
                          )}%`,
                        }}
                      ></div>
                      <span className="mt-2 text-[9px] font-bold uppercase text-gray-400">
                        s/m
                      </span>
                    </div>
                  </div>
                </div>

                <div className="text-[11px] text-gray-400 flex justify-between pt-2">
                  <span>Total no Mês Selecionado:</span>
                  <span className="font-black text-white">
                    {preventivasDoMesHistorico.length} veículos
                  </span>
                </div>
              </div>

              {/* COLUNA DIREITA: TABELA DE PLACAS COM BUSCA INTEGRADA */}
              <div className="lg:col-span-7 rounded-2xl border border-gray-800 bg-[#1C2541] overflow-hidden shadow-xl flex flex-col">
                <div className="p-3 border-b border-gray-800 flex flex-wrap justify-between items-center gap-2 bg-[#111A33]">
                  <div className="font-bold text-xs uppercase text-gray-300">
                    Filtro:{' '}
                    <span className="text-orange-400 font-black">
                      {statusFiltroCard}
                    </span>{' '}
                    ({listagemFiltradaCardsChefe.length})
                  </div>
                  <div className="flex items-center gap-2 bg-[#0B1329] px-3 py-1.5 rounded-lg border border-gray-700 w-60">
                    <Search size={14} className="text-gray-400" />
                    <input
                      type="text"
                      value={termoBuscaGerencial}
                      onChange={(e) => setTermoBuscaGerencial(e.target.value)}
                      placeholder="Placa ou CDV..."
                      className="bg-transparent w-full outline-none text-xs text-white"
                    />
                  </div>
                </div>

                <div className="h-[340px] overflow-y-auto">
                  <table className="w-full text-left text-xs table-auto">
                    <thead className="sticky top-0 bg-[#111A33] font-bold text-gray-400 uppercase text-[10px] border-b border-gray-800">
                      <tr>
                        <th className="p-3">Placa</th>
                        <th className="p-3">Filial / CDV</th>
                        <th className="p-3">Oficina</th>
                        <th className="p-3">Status</th>
                        <th className="p-3">Observação</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-800">
                      {listagemFiltradaCardsChefe.map((p, idx) => (
                        <tr
                          key={`prev-${p.id || idx}`}
                          className="hover:bg-gray-800/20"
                        >
                          <td className="p-3 font-mono font-bold text-white">
                            {p.placa}
                          </td>
                          <td className="p-3 uppercase">
                            <span className="text-blue-400 font-bold block">
                              {p.regional}
                            </span>
                            <span className="text-[10px] text-gray-400">
                              {p.cdv}
                            </span>
                          </td>
                          <td className="p-3 uppercase text-gray-300">
                            {p.oficina || '--'}
                          </td>
                          <td className="p-3">
                            <span
                              className={`text-[10px] font-black px-2 py-0.5 rounded ${
                                p.status === 'Concluído' ||
                                p.status === 'EM DIA'
                                  ? 'bg-green-500/15 text-green-400'
                                  : p.status === 'EM EXECUÇÃO' ||
                                    p.status === 'OFICINA'
                                  ? 'bg-orange-500/15 text-orange-400'
                                  : p.status === 'A VENCER'
                                  ? 'bg-yellow-500/15 text-yellow-400'
                                  : 'bg-red-500/15 text-red-400'
                              }`}
                            >
                              {p.status === 'Concluído' || p.status === 'EM DIA'
                                ? 'EM DIA'
                                : p.status}
                            </span>
                          </td>
                          <td className="p-3 text-gray-400 italic truncate max-w-[160px]">
                            {p.observacao || p.observacao2 || '--'}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ============================================================================ */}
        {/* --- TELA 3: STATUS TACÓGRAFOS (SPLIT-SCREEN: GRÁFICO ESQ + TABELA DIR) --- */}
        {/* ============================================================================ */}
        {activeTab === 'status_tacografos' && (
          <div className="w-full animate-fadeIn space-y-3">
            <div className="flex flex-wrap justify-between items-center gap-2">
              <div className="flex items-center gap-3">
                <h2 className="text-xl font-black tracking-tight text-white">
                  Cronograma de Tacógrafos
                </h2>
                <span className="text-[10px] bg-amber-500/15 text-amber-300 border border-amber-500/30 px-2.5 py-0.5 rounded-full font-bold">
                  🕒 {ultimaAtuTaco} ({tacografos.length} veículos)
                </span>
              </div>
              <div className="flex items-center gap-2 px-3 py-1.5 rounded-xl border border-gray-800 bg-[#1C2541] text-xs font-bold text-white">
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

            {/* FAIXA COMPACTA DE 6 CARDS */}
            <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-2.5">
              {[
                {
                  id: 'EM DIA',
                  label: 'EM DIA',
                  pct: `${pctTacoEmDia}%`,
                  qtd: qtdTacoEmDiaComData,
                  grad: 'from-green-500 to-emerald-600 text-white',
                },
                {
                  id: 'EM EXECUÇÃO',
                  label: 'EM EXEC.',
                  pct: `${pctTacoExecucao}%`,
                  qtd: qtdTacoExecucaoComData,
                  grad: 'from-orange-500 to-amber-600 text-white',
                },
                {
                  id: 'A VENCER',
                  label: 'A VENCER',
                  pct: `${pctTacoAVencer}%`,
                  qtd: qtdTacoAVencerComData,
                  grad: 'from-yellow-400 to-yellow-500 text-gray-900',
                },
                {
                  id: 'VENCIDAS',
                  label: 'VENCIDAS',
                  pct: `${pctTacoVencidas}%`,
                  qtd: qtdTacoVencidasComData,
                  grad: 'from-red-500 to-rose-600 text-white',
                },
                {
                  id: 'FUTUROS',
                  label: '2027+ FUTUROS',
                  pct: 'VALIDADE',
                  qtd: tacosAnosFuturos.length,
                  grad: 'from-blue-600 to-indigo-700 text-white',
                },
                {
                  id: 'SEM DATA',
                  label: 'SEM MÊS',
                  pct: 'PREENCHER',
                  qtd: tacosSemDataFilaGeral.length,
                  grad: 'from-gray-600 to-gray-700 text-white',
                },
              ].map((c) => (
                <div
                  key={c.id}
                  onClick={() =>
                    setStatusFiltroCardTaco(
                      statusFiltroCardTaco === c.id ? 'TODOS' : c.id
                    )
                  }
                  className={`p-3 rounded-xl shadow-sm border border-gray-800 cursor-pointer transition-all ${
                    statusFiltroCardTaco === c.id ? 'ring-2 ring-white' : ''
                  } bg-gradient-to-br ${c.grad}`}
                >
                  <div className="flex justify-between items-center">
                    <span className="text-[10px] font-black">{c.label}</span>
                    <span className="text-[8px] bg-black/20 px-1 py-0.5 rounded font-bold">
                      {c.pct}
                    </span>
                  </div>
                  <h3 className="text-2xl font-black mt-1">
                    {String(c.qtd).padStart(2, '0')}{' '}
                    <span className="text-[9px] font-normal opacity-80">
                      Placas
                    </span>
                  </h3>
                </div>
              ))}
            </div>

            {/* SPLIT-SCREEN: GRÁFICO (5 COLUNAS) + TABELA (7 COLUNAS) */}
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-4 items-stretch">
              <div className="lg:col-span-5 p-4 rounded-2xl border border-gray-800/80 bg-[#131B35] shadow-xl flex flex-col justify-between">
                <h3 className="text-sm font-black text-white uppercase flex items-center justify-between">
                  <span>📊 Cronograma Anual (Tacógrafos)</span>
                  <span className="text-[10px] text-amber-400 font-bold">
                    Visão 2026
                  </span>
                </h3>

                <div className="relative pt-6 pb-2 my-auto">
                  <div className="h-56 flex items-end justify-between gap-1 px-1 relative z-10 border-b border-gray-700 pb-2">
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
                            className={`text-[9px] font-black mb-1 ${
                              estaSelecionado
                                ? 'text-amber-400 scale-110'
                                : 'text-gray-400'
                            }`}
                          >
                            {item.qtd > 0 ? item.qtd : ''}
                          </span>
                          <div
                            className={`w-full max-w-[24px] rounded-t-md transition-all ${
                              estaSelecionado
                                ? 'bg-gradient-to-t from-rose-600 via-orange-500 to-amber-300 ring-2 ring-amber-400'
                                : 'bg-gradient-to-t from-rose-900/60 via-orange-600/70 to-amber-500/80'
                            }`}
                            style={{ height: `${alturaPct}%` }}
                          ></div>
                          <span
                            className={`mt-2 text-[9px] font-bold uppercase ${
                              estaSelecionado
                                ? 'text-amber-400 font-black'
                                : 'text-gray-400'
                            }`}
                          >
                            {item.sigla}
                          </span>
                        </div>
                      );
                    })}

                    <div
                      onClick={() => setStatusFiltroCardTaco('FUTUROS')}
                      className="flex-1 flex flex-col items-center h-full justify-end cursor-pointer border-l border-gray-800 pl-1"
                    >
                      <span className="text-[9px] font-black text-blue-400 mb-1">
                        {tacosAnosFuturos.length || ''}
                      </span>
                      <div
                        className="w-full max-w-[24px] rounded-t-md bg-gradient-to-t from-blue-900 to-cyan-400"
                        style={{
                          height: `${Math.max(
                            (tacosAnosFuturos.length / maxQtdMesesTaco) * 100,
                            4
                          )}%`,
                        }}
                      ></div>
                      <span className="mt-2 text-[9px] font-black text-blue-400">
                        27+
                      </span>
                    </div>

                    <div
                      onClick={() => setStatusFiltroCardTaco('SEM DATA')}
                      className="flex-1 flex flex-col items-center h-full justify-end cursor-pointer border-l border-gray-800 pl-1"
                    >
                      <span className="text-[9px] font-black text-gray-400 mb-1">
                        {tacosSemDataFilaGeral.length || ''}
                      </span>
                      <div
                        className="w-full max-w-[24px] rounded-t-md bg-gradient-to-t from-gray-700 to-gray-400"
                        style={{
                          height: `${Math.max(
                            (tacosSemDataFilaGeral.length / maxQtdMesesTaco) *
                              100,
                            4
                          )}%`,
                        }}
                      ></div>
                      <span className="mt-2 text-[9px] font-bold text-gray-400">
                        s/m
                      </span>
                    </div>
                  </div>
                </div>

                <div className="text-[11px] text-gray-400 flex justify-between pt-2">
                  <span>Total Mapeado em Tacógrafos:</span>
                  <span className="font-black text-white">
                    {tacografos.length} veículos
                  </span>
                </div>
              </div>

              <div className="lg:col-span-7 rounded-2xl border border-gray-800 bg-[#1C2541] overflow-hidden shadow-xl flex flex-col">
                <div className="p-3 border-b border-gray-800 flex flex-wrap justify-between items-center gap-2 bg-[#111A33]">
                  <div className="font-bold text-xs uppercase text-gray-300">
                    Lista:{' '}
                    <span className="text-amber-400 font-black">
                      {statusFiltroCardTaco}
                    </span>{' '}
                    ({listagemFiltradaTacosChefe.length})
                  </div>
                  <div className="flex items-center gap-2 bg-[#0B1329] px-3 py-1.5 rounded-lg border border-gray-700 w-60">
                    <Search size={14} className="text-gray-400" />
                    <input
                      type="text"
                      value={termoBuscaGerencialTaco}
                      onChange={(e) =>
                        setTermoBuscaGerencialTaco(e.target.value)
                      }
                      placeholder="Placa ou CDV..."
                      className="bg-transparent w-full outline-none text-xs text-white"
                    />
                  </div>
                </div>

                <div className="h-[340px] overflow-y-auto">
                  <table className="w-full text-left text-xs table-auto">
                    <thead className="sticky top-0 bg-[#111A33] font-bold text-gray-400 uppercase text-[10px] border-b border-gray-800">
                      <tr>
                        <th className="p-3">Placa</th>
                        <th className="p-3">Modelo</th>
                        <th className="p-3">CDV / Filial</th>
                        <th className="p-3">Validade</th>
                        <th className="p-3">Status</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-800">
                      {listagemFiltradaTacosChefe.map((t, idx) => (
                        <tr
                          key={`taco-ger-${t.id || idx}`}
                          className="hover:bg-gray-800/20"
                        >
                          <td className="p-3 font-mono font-bold text-white">
                            {t.placa}
                          </td>
                          <td className="p-3 text-gray-300 uppercase">
                            {t.modelo || 'CAMINHÃO'}
                          </td>
                          <td className="p-3 uppercase">
                            <span className="text-blue-400 font-bold block">
                              {t.regional}
                            </span>
                            <span className="text-[10px] text-gray-400">
                              {t.cdv}
                            </span>
                          </td>
                          <td className="p-3 font-bold">
                            {t.validade
                              ? new Date(
                                  t.validade + 'T00:00:00'
                                ).toLocaleDateString('pt-BR')
                              : '--'}
                          </td>
                          <td className="p-3">
                            <span
                              className={`text-[10px] font-black px-2 py-0.5 rounded ${
                                t.status === 'Concluído' ||
                                t.status === 'EM DIA'
                                  ? 'bg-green-500/15 text-green-400'
                                  : t.status === 'EM EXECUÇÃO' ||
                                    t.status === 'OFICINA'
                                  ? 'bg-orange-500/15 text-orange-400'
                                  : t.status === 'A VENCER'
                                  ? 'bg-yellow-500/15 text-yellow-400'
                                  : 'bg-red-500/15 text-red-400'
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
          </div>
        )}

        {/* ============================================================================ */}
        {/* --- TELA 4: STATUS FINANCEIRO (COCKPIT COMPACTO COM GAVETA & ABAS) --- */}
        {/* ============================================================================ */}
        {activeTab === 'status_financeiro' && (
          <div className="w-full animate-fadeIn space-y-3">
            {/* CABEÇALHO COMPACTO COM BOTÃO DE GAVETA DE UPLOAD E ALERTA CTA */}
            <div className="flex flex-wrap justify-between items-center gap-2 border-b border-gray-800 pb-2.5">
              <div className="flex flex-wrap items-center gap-2.5">
                <DollarSign className="text-rose-400" size={22} />
                <h2 className="text-xl font-black tracking-tight text-white">
                  Status Financeiro & Rateio
                </h2>
                <span className="text-[10px] bg-rose-500/15 text-rose-300 border border-rose-500/30 px-2.5 py-0.5 rounded-full font-bold">
                  🕒 {ultimaAtuFin}
                </span>
              </div>

              <div className="flex flex-wrap items-center gap-2">
                {finCta.length > 0 && (
                  <button
                    onClick={() => setMostrarCtaExpandido(!mostrarCtaExpandido)}
                    className="px-3 py-1.5 rounded-xl bg-red-600/20 hover:bg-red-600/30 border border-red-500/50 text-red-300 text-xs font-black flex items-center gap-1.5 animate-pulse"
                  >
                    <AlertTriangle size={14} className="text-red-400" />
                    <span>
                      🚨 {finCta.length} CTA Críticos (
                      {formatarMoeda(
                        finCta.reduce(
                          (acc, c) => acc + (Number(c.valor) || 0),
                          0
                        )
                      )}
                      ) {mostrarCtaExpandido ? '▴' : '▾'}
                    </span>
                  </button>
                )}

                <button
                  onClick={() =>
                    setMostrarGavetaUploadFin(!mostrarGavetaUploadFin)
                  }
                  className="px-3 py-1.5 rounded-xl bg-[#1C2541] hover:bg-gray-800 border border-gray-700 text-rose-300 text-xs font-black flex items-center gap-1.5"
                >
                  <Upload size={13} />
                  <span>Importar / Banco {mostrarGavetaUploadFin ? '▴' : '▾'}</span>
                </button>

                <div className="flex items-center gap-1.5 px-3 py-1.5 rounded-xl border border-gray-800 bg-[#1C2541] text-xs font-bold text-white">
                  <Clock size={13} className="text-rose-400" />
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
                        {m.nome} {mesOficialRateio === m.num ? '(Oficial)' : ''}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
            </div>

            {/* GAVETA RETRÁTIL DE IMPORTAÇÃO FINANCEIRA */}
            {mostrarGavetaUploadFin && (
              <div className="p-3.5 bg-[#1C2541] border border-rose-500/30 rounded-xl flex flex-wrap justify-between items-center gap-3 animate-fadeIn">
                <div className="flex items-center gap-3">
                  <span className="text-xs font-black uppercase text-rose-400">
                    💸 Planilha OS's - SUL2 - 2026:
                  </span>
                  <label className="flex items-center gap-1.5 cursor-pointer text-xs text-gray-300">
                    <input
                      type="checkbox"
                      checked={substituirFinanceiro}
                      onChange={(e) =>
                        setSubstituirFinanceiro(e.target.checked)
                      }
                      className="accent-rose-500"
                    />
                    Substituir base atual
                  </label>
                </div>
                <div className="flex gap-2">
                  <button
                    onClick={limparTodaTabelaFinanceiro}
                    className="bg-red-600/10 text-red-400 hover:bg-red-600 hover:text-white border border-red-500/20 text-xs font-black px-3 py-1.5 rounded-lg flex items-center gap-1"
                  >
                    <Trash2 size={13} /> Limpar Banco
                  </button>
                  <label className="bg-gradient-to-r from-rose-600 to-orange-500 text-white font-black text-xs px-3.5 py-1.5 rounded-lg cursor-pointer flex items-center gap-1.5">
                    <Upload size={13} />
                    Selecionar Excel Financeiro
                    <input
                      type="file"
                      accept=".xlsx, .xls"
                      onChange={lidarComImportacaoFinanceiro}
                      className="hidden"
                    />
                  </label>
                </div>
              </div>
            )}

            {/* GAVETA RETRÁTIL DE ALERTAS CTA */}
            {mostrarCtaExpandido && finCta.length > 0 && (
              <div className="p-3.5 rounded-xl border border-red-500/50 bg-[#1C2541] grid grid-cols-1 md:grid-cols-3 gap-2.5 animate-fadeIn">
                {finCta.map((item, idx) => (
                  <div
                    key={item.id || idx}
                    onClick={() => {
                      setTermoBuscaFinanceiro(item.placa);
                      setSubVisaoFinanceiro('AUDITORIA_OS');
                    }}
                    className="p-2.5 rounded-lg bg-[#0B1329] border border-red-500/30 hover:border-red-400 cursor-pointer flex justify-between items-center gap-2"
                  >
                    <div>
                      <div className="flex items-center gap-1.5">
                        <span className="font-mono font-black text-xs text-white bg-red-600/30 px-1.5 py-0.5 rounded">
                          {item.placa}
                        </span>
                        <span className="text-[10px] font-black text-amber-400 uppercase">
                          {item.servico}
                        </span>
                      </div>
                      <p className="text-[10px] text-gray-400 truncate max-w-[160px] mt-0.5">
                        {item.oficina}
                      </p>
                    </div>
                    <div className="text-right">
                      <span className="text-[9px] font-black text-red-400 block">
                        {item.status}
                      </span>
                      <span className="text-xs font-black text-white">
                        {formatarMoeda(item.valor)}
                      </span>
                    </div>
                  </div>
                ))}
              </div>
            )}

            {/* 4 CARDS EXECUTIVOS COMPACTOS */}
            <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
              <div
                onClick={() => {
                  setContaFiltroFinanceiro('TODAS');
                  setNfFiltroFinanceiro('TODAS');
                  setCdvFiltroFinanceiro('TODOS');
                }}
                className="p-3 rounded-xl border border-gray-800 cursor-pointer bg-gradient-to-br from-blue-600 to-indigo-700 text-white"
              >
                <div className="flex justify-between items-center">
                  <span className="text-[10px] font-black uppercase">
                    ORÇAMENTO (AOP + PROV)
                  </span>
                  <span className="text-[9px] bg-white/20 px-1.5 py-0.5 rounded font-bold">
                    MÊS {mesSelecionadoFinanceiro}
                  </span>
                </div>
                <h3 className="text-xl font-black mt-1">
                  {formatarMoeda(totalOrcamentoMes)}
                </h3>
              </div>

              <div
                onClick={() =>
                  setNfFiltroFinanceiro(
                    nfFiltroFinanceiro === 'COBRADAS' ? 'TODAS' : 'COBRADAS'
                  )
                }
                className={`p-3 rounded-xl border border-gray-800 cursor-pointer ${
                  nfFiltroFinanceiro === 'COBRADAS'
                    ? 'ring-2 ring-emerald-400'
                    : ''
                } bg-gradient-to-br from-emerald-600 to-teal-700 text-white`}
              >
                <div className="flex justify-between items-center">
                  <span className="text-[10px] font-black uppercase">
                    FATURADO (NFs)
                  </span>
                  <span className="text-[9px] bg-white/20 px-1.5 py-0.5 rounded font-bold">
                    {totalOrcamentoMes > 0
                      ? (
                          (totalFaturadoNFsMes / totalOrcamentoMes) *
                          100
                        ).toFixed(1)
                      : '0.0'}
                    %
                  </span>
                </div>
                <h3 className="text-xl font-black mt-1">
                  {formatarMoeda(totalFaturadoNFsMes)}
                </h3>
              </div>

              <div
                onClick={() =>
                  setNfFiltroFinanceiro(
                    nfFiltroFinanceiro === 'PEND' ? 'TODAS' : 'PEND'
                  )
                }
                className={`p-3 rounded-xl border border-gray-800 cursor-pointer ${
                  nfFiltroFinanceiro === 'PEND' ? 'ring-2 ring-amber-400' : ''
                } bg-gradient-to-br from-orange-500 to-amber-600 text-white`}
              >
                <div className="flex justify-between items-center">
                  <span className="text-[10px] font-black uppercase">
                    PENDENTE (PEND)
                  </span>
                  <span className="text-[9px] bg-white/20 px-1.5 py-0.5 rounded font-bold">
                    {totalOrcamentoMes > 0
                      ? ((totalPendenteMes / totalOrcamentoMes) * 100).toFixed(
                          1
                        )
                      : '0.0'}
                    %
                  </span>
                </div>
                <h3 className="text-xl font-black mt-1">
                  {formatarMoeda(totalPendenteMes)}
                </h3>
              </div>

              <div
                className={`p-3 rounded-xl border border-gray-800 ${
                  saldoFinalMes >= 0
                    ? 'bg-gradient-to-br from-green-600 to-emerald-700'
                    : 'bg-gradient-to-br from-red-600 to-rose-700'
                } text-white`}
              >
                <div className="flex justify-between items-center">
                  <span className="text-[10px] font-black uppercase">
                    SALDO FINAL
                  </span>
                  <span className="text-[8px] bg-black/20 px-1.5 py-0.5 rounded font-black">
                    {saldoFinalMes >= 0 ? 'DENTRO DO AOP' : 'ESTOURO'}
                  </span>
                </div>
                <h3 className="text-xl font-black mt-1">
                  {formatarMoeda(saldoFinalMes)}
                </h3>
              </div>
            </div>

{/* SELETOR DE 3 ABAS FINANCEIRAS WIDESCREEN (ZERO-SCROLL COM TABELAS COMPLETAS!) */}
<div className="flex flex-wrap items-center justify-between gap-2 bg-[#111A33] p-2 rounded-xl border border-gray-800">
              <div className="flex flex-wrap items-center gap-1.5">
                <button
                  onClick={() => setSubVisaoFinanceiro('MATRIZES')}
                  className={`px-3.5 py-1.5 rounded-lg text-xs font-black uppercase transition-all ${
                    subVisaoFinanceiro === 'MATRIZES'
                      ? 'bg-rose-500 text-white shadow-md'
                      : 'text-gray-400 hover:text-white'
                  }`}
                >
                  🏢 Ranking por CDV ({analisePorCdv.length} Unidades)
                </button>
                <button
                  onClick={() => setSubVisaoFinanceiro('RATEIO_CONTAS')}
                  className={`px-3.5 py-1.5 rounded-lg text-xs font-black uppercase transition-all ${
                    subVisaoFinanceiro === 'RATEIO_CONTAS'
                      ? 'bg-rose-500 text-white shadow-md'
                      : 'text-gray-400 hover:text-white'
                  }`}
                >
                  📋 Matriz Divisão de Rateio (Mês {mesSelecionadoFinanceiro})
                </button>
                <button
                  onClick={() => setSubVisaoFinanceiro('AUDITORIA_OS')}
                  className={`px-3.5 py-1.5 rounded-lg text-xs font-black uppercase transition-all ${
                    subVisaoFinanceiro === 'AUDITORIA_OS'
                      ? 'bg-rose-500 text-white shadow-md'
                      : 'text-gray-400 hover:text-white'
                  }`}
                >
                  🔍 Auditoria Detalhada de OSs ({listagemOsFinanceiroFiltrada.length})
                </button>
              </div>

              {(contaFiltroFinanceiro !== 'TODAS' ||
                cdvFiltroFinanceiro !== 'TODOS' ||
                nfFiltroFinanceiro !== 'TODAS') && (
                <button
                  onClick={() => {
                    setContaFiltroFinanceiro('TODAS');
                    setCdvFiltroFinanceiro('TODOS');
                    setNfFiltroFinanceiro('TODAS');
                  }}
                  className="text-[10px] font-black bg-rose-500/20 text-rose-300 border border-rose-500/30 px-2.5 py-1 rounded-lg"
                >
                  Limpar Filtros ({contaFiltroFinanceiro} / {cdvFiltroFinanceiro}) ✕
                </button>
              )}
            </div>

            {/* ==================================================================== */}
            {/* ABA 1: DESEMPENHO ORÇAMENTÁRIO & RANKING POR CDV (COMPLETO COM BARRAS!) */}
            {/* ==================================================================== */}
            {subVisaoFinanceiro === 'MATRIZES' && (
              <div className="rounded-2xl border border-gray-800 bg-[#1C2541] overflow-hidden shadow-xl animate-fadeIn">
                <div className="p-3.5 border-b border-gray-800 bg-[#111A33] flex flex-wrap justify-between items-center gap-2">
                  <div>
                    <h3 className="text-xs font-black uppercase text-white flex items-center gap-2">
                      <span>
                        🏢 Desempenho Orçamentário & Ranking por CDV ({totalPlacasSul2} Placas)
                      </span>
                    </h3>
                    <p className="text-[11px] text-gray-400">
                      Comparativo entre o Orçamento Proporcional por frota vs. Gasto Real (BASE SUL II). Clique num CDV para ver suas OSs.
                    </p>
                  </div>
                </div>

                <div className="max-h-[calc(100vh-290px)] overflow-y-auto">
                  <table className="w-full text-left text-xs table-auto">
                    <thead className="sticky top-0 bg-[#0B1329] font-black text-gray-400 uppercase text-[10px] border-b border-gray-800 z-10">
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
                        <th className="p-3.5 w-52">
                          Peso no Gasto Geral
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
                            onClick={() => {
                              setCdvFiltroFinanceiro(
                                selecionado ? 'TODOS' : item.cdv
                              );
                              setSubVisaoFinanceiro('AUDITORIA_OS');
                            }}
                            className={`transition-all cursor-pointer ${
                              selecionado
                                ? 'bg-blue-500/20 border-l-4 border-blue-400'
                                : 'hover:bg-gray-800/30'
                            }`}
                          >
                            <td className="p-3.5 font-black text-white flex items-center gap-2.5">
                              <span
                                className={`text-[10px] font-mono px-2 py-0.5 rounded border ${
                                  idx === 0
                                    ? 'bg-amber-500/20 text-amber-300 border-amber-500/40'
                                    : idx === 1
                                    ? 'bg-slate-400/20 text-slate-200 border-slate-400/40'
                                    : idx === 2
                                    ? 'bg-orange-700/20 text-orange-300 border-orange-600/40'
                                    : 'bg-[#0B1329] text-gray-400 border-gray-700'
                                }`}
                              >
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
                                  className={`text-[11px] font-mono font-black px-2.5 py-1 rounded-md ${
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
                                  <span className="text-red-400 inline-flex items-center gap-1.5">
                                    <span>+{item.pctEstouro.toFixed(2)}% acima</span>
                                    <span className="w-2.5 h-2.5 rounded-full bg-red-500 shadow-[0_0_8px_#ef4444]" />
                                  </span>
                                ) : (
                                  <span className="text-emerald-400 inline-flex items-center gap-1.5">
                                    <span>{item.pctEstouro.toFixed(2)}% (Econ.)</span>
                                    <span className="w-2.5 h-2.5 rounded-full bg-emerald-400 shadow-[0_0_8px_#34d399]" />
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
                              <div className="flex items-center gap-2.5">
                                <div className="flex-1 h-2.5 bg-[#0B1329] rounded-full overflow-hidden border border-gray-800">
                                  <div
                                    className={`h-full rounded-full ${
                                      estourou
                                        ? 'bg-gradient-to-r from-orange-500 to-red-500'
                                        : 'bg-gradient-to-r from-blue-500 to-emerald-400'
                                    }`}
                                    style={{
                                      width: `${Math.min(
                                        item.shareGastoTotal * 2.5,
                                        100
                                      )}%`,
                                    }}
                                  ></div>
                                </div>
                                <span className="font-mono text-[11px] font-black text-amber-400 w-12 text-right">
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
            )}

            {/* ==================================================================== */}
            {/* ABA 2: MATRIZ OFICIAL DIVISÃO DE RATEIO (TODAS AS NFs + TERMÔMETRO) */}
            {/* ==================================================================== */}
            {subVisaoFinanceiro === 'RATEIO_CONTAS' && (
              <div className="rounded-2xl border border-gray-800 bg-[#1C2541] overflow-hidden shadow-xl animate-fadeIn">
                <div className="p-3.5 border-b border-gray-800 bg-[#111A33] flex flex-wrap justify-between items-center gap-2">
                  <div>
                    <h3 className="text-xs font-black uppercase text-white flex items-center gap-2">
                      <span>
                        📋 Matriz Oficial Divisão de Rateio (Mês {mesSelecionadoFinanceiro})
                      </span>
                      {usandoRateioDireto && (
                        <span className="text-[10px] bg-emerald-500/20 text-emerald-400 border border-emerald-500/30 px-2 py-0.5 rounded-full">
                          Leitura Direta da Aba DIVISÃO RATEIO
                        </span>
                      )}
                    </h3>
                  </div>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-left text-xs table-auto">
                    <thead className="bg-[#0B1329] font-black text-gray-400 uppercase text-[10px] border-b border-gray-800">
                      <tr>
                        <th className="p-3.5">Conta Correção</th>
                        <th className="p-3.5 text-right text-blue-400">AOP (Meta)</th>
                        <th className="p-3.5 text-right text-indigo-400">PROV</th>
                        <th className="p-3.5 text-right text-orange-400">PEND</th>
                        <th className="p-3.5 text-right text-emerald-400">{labelNf1}</th>
                        <th className="p-3.5 text-right text-emerald-400">{labelNf2}</th>
                        <th className="p-3.5 text-right text-emerald-400">{labelNf3}</th>
                        <th className="p-3.5 text-right text-emerald-400">{labelNf4}</th>
                        <th className="p-3.5 text-right text-white">SALDO</th>
                        <th className="p-3.5 w-44">Consumo do AOP</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-800">
                      {linhasMatrizRateio.map((linha, index) => {
                        const isTotal = String(linha.conta)
                          .toUpperCase()
                          .includes('TOTAL GERAL');
                        const tetoConta =
                          (Number(linha.aop) || 0) + (Number(linha.prov) || 0);
                        const gastoConta =
                          (Number(linha.pend) || 0) +
                          (Number(linha.nf1) || 0) +
                          (Number(linha.nf2) || 0) +
                          (Number(linha.nf3) || 0) +
                          (Number(linha.nf4) || 0);
                        const pctConsumo =
                          tetoConta > 0 ? (gastoConta / tetoConta) * 100 : 0;
                        const estourouConta = pctConsumo > 100;

                        return (
                          <tr
                            key={index}
                            onClick={() => {
                              if (!isTotal) {
                                setContaFiltroFinanceiro(
                                  String(linha.conta).toUpperCase()
                                );
                                setSubVisaoFinanceiro('AUDITORIA_OS');
                              }
                            }}
                            className={`transition-all ${
                              isTotal
                                ? 'bg-[#111A33] font-black text-white border-t-2 border-gray-700'
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
                            <td className="p-3.5">
                              <div className="flex items-center gap-2">
                                <div className="flex-1 h-2 bg-[#0B1329] rounded-full overflow-hidden border border-gray-800">
                                  <div
                                    className={`h-full rounded-full ${
                                      estourouConta
                                        ? 'bg-gradient-to-r from-orange-500 to-red-500'
                                        : 'bg-gradient-to-r from-emerald-500 to-green-400'
                                    }`}
                                    style={{
                                      width: `${Math.min(pctConsumo, 100)}%`,
                                    }}
                                  ></div>
                                </div>
                                <span
                                  className={`font-mono text-[10px] font-black w-12 text-right ${
                                    estourouConta
                                      ? 'text-red-400'
                                      : 'text-emerald-400'
                                  }`}
                                >
                                  {pctConsumo.toFixed(0)}%
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
            )}

            {/* ==================================================================== */}
            {/* ABA 3: AUDITORIA DETALHADA DE ORDENS DE SERVIÇO (BASE SUL II) */}
            {/* ==================================================================== */}
            {subVisaoFinanceiro === 'AUDITORIA_OS' && (
              <div className="rounded-2xl border border-gray-800 bg-[#1C2541] overflow-hidden shadow-xl animate-fadeIn">
                <div className="p-3 border-b border-gray-800 flex flex-wrap justify-between items-center gap-2 bg-[#111A33]">
                  <div className="flex items-center gap-2 flex-1 min-w-[240px] bg-[#0B1329] px-3 py-1.5 rounded-lg border border-gray-700">
                    <Search size={14} className="text-gray-400" />
                    <input
                      type="text"
                      value={termoBuscaFinanceiro}
                      onChange={(e) => setTermoBuscaFinanceiro(e.target.value)}
                      placeholder="Pesquisar OS por Placa, Número da OS, CDV, Oficina ou NF..."
                      className="bg-transparent w-full outline-none text-xs text-white"
                    />
                  </div>
                  <div className="text-xs font-black text-emerald-400">
                    Soma ({listagemOsFinanceiroFiltrada.length} OSs):{' '}
                    {formatarMoeda(
                      listagemOsFinanceiroFiltrada.reduce(
                        (acc, b) => acc + (Number(b.valor_total) || 0),
                        0
                      )
                    )}
                  </div>
                </div>
                <div className="max-h-[calc(100vh-290px)] overflow-y-auto">
                  <table className="w-full text-left text-xs table-auto">
                    <thead className="sticky top-0 bg-[#111A33] font-bold text-gray-400 uppercase text-[10px] border-b border-gray-800">
                      <tr>
                        <th className="p-3">OS</th>
                        <th className="p-3">Data</th>
                        <th className="p-3">Placa</th>
                        <th className="p-3">CDV</th>
                        <th className="p-3">Oficina</th>
                        <th className="p-3">Conta</th>
                        <th className="p-3">NF / Lote</th>
                        <th className="p-3 text-right">Valor Total</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-800">
                      {listagemOsFinanceiroFiltrada
                        .slice(0, 300)
                        .map((b, idx) => (
                          <tr
                            key={b.id || idx}
                            className="hover:bg-gray-800/20"
                          >
                            <td className="p-3 font-mono text-gray-300 font-bold">
                              {b.os}
                            </td>
                            <td className="p-3 text-gray-400">
                              {b.data_os
                                ? new Date(
                                    b.data_os + 'T00:00:00'
                                  ).toLocaleDateString('pt-BR')
                                : '--'}
                            </td>
                            <td className="p-3 font-mono font-black text-white">
                              {b.placa}
                            </td>
                            <td className="p-3 uppercase text-blue-400 font-semibold">
                              {b.cdv}
                            </td>
                            <td className="p-3 text-gray-300 truncate max-w-[180px]">
                              {b.estabelecimento}
                            </td>
                            <td className="p-3">
                              <span className="text-[10px] font-black px-2 py-0.5 rounded bg-gray-800 text-amber-400 uppercase">
                                {b.conta_correcao}
                              </span>
                            </td>
                            <td className="p-3">
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
                            <td className="p-3 text-right font-mono font-black text-white">
                              {formatarMoeda(b.valor_total)}
                            </td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
        )}

        {/* ============================================================================ */}
        {/* --- TELA 5: COCKPIT VEÍCULOS PARADOS 3D/2D LADO A LADO (ZERO-SCROLL!) --- */}
        {/* ============================================================================ */}
        {activeTab === 'veiculos_parados' && (
          <div className="w-full animate-fadeIn space-y-3">
            {/* LINHA 1: CABEÇALHO ULTRA-COMPACTO + FILTROS + UPLOAD NA MESMA BARRA */}
            <div className="flex flex-wrap justify-between items-center gap-2 border-b border-gray-800 pb-2.5">
              <div className="flex flex-wrap items-center gap-2.5">
                <Truck className="text-cyan-400" size={22} />
                <h2 className="text-xl font-black tracking-tight text-white">
                  Cockpit de Disponibilidade & Vitrine 3D/2D
                </h2>
                <span className="text-[10px] bg-cyan-500/15 text-cyan-300 border border-cyan-500/30 px-2.5 py-0.5 rounded-full font-bold">
                  🕒 {ultimaAtuParados}
                </span>
              </div>

              <div className="flex flex-wrap items-center gap-2">
                {/* Busca Rápida Placa */}
                <div className="flex items-center gap-1.5 bg-[#1C2541] border border-gray-700 focus-within:border-cyan-400 px-2.5 py-1.5 rounded-xl w-44">
                  <Search size={13} className="text-cyan-400 shrink-0" />
                  <input
                    type="text"
                    value={buscaPlaca3DTopo}
                    onChange={(e) => executarBuscaRapidaPlaca3D(e.target.value)}
                    placeholder="Buscar Placa 3D..."
                    className="bg-transparent w-full outline-none text-xs font-mono font-bold text-white uppercase"
                  />
                </div>

                {/* Dropdown de Oficinas */}
                <div className="flex items-center gap-1.5 bg-[#1C2541] border border-gray-700 px-2.5 py-1.5 rounded-xl">
                  <Building2 size={13} className="text-amber-400 shrink-0" />
                  <select
                    value={filtroOficinaParados}
                    onChange={(e) => setFiltroOficinaParados(e.target.value)}
                    className="bg-transparent text-xs font-black text-white uppercase outline-none cursor-pointer"
                  >
                    <option value="TODAS" className="bg-[#0B1329] text-white">
                      Oficinas ({oficinasUnicasParados.length})
                    </option>
                    {oficinasUnicasParados.map((ofic) => (
                      <option
                        key={ofic}
                        value={ofic}
                        className="bg-[#0B1329] text-white"
                      >
                        🏢 {ofic}
                      </option>
                    ))}
                  </select>
                </div>

                <button
                  onClick={limparTodaTabelaParados}
                  className="bg-red-600/10 text-red-400 hover:bg-red-600 hover:text-white border border-red-500/20 text-xs font-black px-2.5 py-1.5 rounded-xl transition-all flex items-center gap-1"
                  title="Limpar Banco de Parados"
                >
                  <Trash2 size={13} />
                </button>

                <label className="bg-gradient-to-r from-cyan-600 to-blue-600 hover:from-cyan-500 hover:to-blue-500 text-white font-black text-xs px-3 py-1.5 rounded-xl cursor-pointer flex items-center gap-1.5 shadow-md">
                  <Upload size={13} />
                  <span>Carregar Excel</span>
                  <input
                    type="file"
                    accept=".xlsx, .xls"
                    onChange={lidarComImportacaoVeiculosParados}
                    className="hidden"
                  />
                </label>
              </div>
            </div>

            {/* LINHA 2: 4 CARDS COMPACTOS DE DISPONIBILIDADE */}
            <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
              <div
                onClick={() => setFiltroDispoParados('TODOS')}
                className="p-2.5 px-3.5 rounded-xl border border-gray-800 cursor-pointer bg-gradient-to-br from-blue-600 to-indigo-700 text-white flex justify-between items-center"
              >
                <div>
                  <span className="text-[10px] font-black uppercase block opacity-80">
                    FROTA TOTAL MAPEADA
                  </span>
                  <h3 className="text-2xl font-black">{totalFrotaResumo}</h3>
                </div>
                <span className="text-[9px] bg-white/20 px-2 py-0.5 rounded font-bold">
                  POSIÇÃO NACIONAL
                </span>
              </div>

              <div
                onClick={() =>
                  setFiltroDispoParados(
                    filtroDispoParados === 'DISPON' ? 'TODOS' : 'DISPON'
                  )
                }
                className={`p-2.5 px-3.5 rounded-xl border border-gray-800 cursor-pointer flex justify-between items-center ${
                  filtroDispoParados === 'DISPON'
                    ? 'ring-2 ring-emerald-400'
                    : ''
                } bg-gradient-to-br from-emerald-600 to-green-700 text-white`}
              >
                <div>
                  <span className="text-[10px] font-black uppercase block opacity-80">
                    DISPONÍVEIS RODANDO
                  </span>
                  <h3 className="text-2xl font-black">
                    {totalDisponivelResumo}
                  </h3>
                </div>
                <span className="text-xs bg-white/20 px-2 py-0.5 rounded font-black">
                  {pctDispoGeral}%
                </span>
              </div>

              <div
                onClick={() =>
                  setFiltroDispoParados(
                    filtroDispoParados === 'OFICINA' ? 'TODOS' : 'OFICINA'
                  )
                }
                className={`p-2.5 px-3.5 rounded-xl border border-gray-800 cursor-pointer flex justify-between items-center ${
                  filtroDispoParados === 'OFICINA' ? 'ring-2 ring-red-400' : ''
                } bg-gradient-to-br from-rose-600 to-red-700 text-white`}
              >
                <div>
                  <span className="text-[10px] font-black uppercase block opacity-80">
                    PARADOS EM OFICINA
                  </span>
                  <h3 className="text-2xl font-black">{totalOficinaResumo}</h3>
                </div>
                <span className="text-xs bg-white/20 px-2 py-0.5 rounded font-bold">
                  {totalFrotaResumo > 0
                    ? ((totalOficinaResumo / totalFrotaResumo) * 100).toFixed(1)
                    : '0.0'}
                  %
                </span>
              </div>

              <div
                onClick={() =>
                  setFiltroDispoParados(
                    filtroDispoParados === 'LIBERA' ? 'TODOS' : 'LIBERA'
                  )
                }
                className={`p-2.5 px-3.5 rounded-xl border border-gray-800 cursor-pointer flex justify-between items-center ${
                  filtroDispoParados === 'LIBERA' ? 'ring-2 ring-amber-400' : ''
                } bg-gradient-to-br from-amber-500 to-orange-600 text-white`}
              >
                <div>
                  <span className="text-[10px] font-black uppercase block opacity-80">
                    P/ LIBERAÇÃO IMEDIATA
                  </span>
                  <h3 className="text-2xl font-black">
                    {totalLiberacaoResumo}
                  </h3>
                </div>
                <span className="text-[9px] bg-black/20 px-2 py-0.5 rounded font-black">
                  PRONTOS
                </span>
              </div>
            </div>

            {/* LINHA 3: COCKPIT SPLIT-SCREEN (ESQUERDA: CAMINHÃO 3D/2D | DIREITA: PLACAS / CDVs / LAUDO) */}
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-4 items-stretch">
              {/* COLUNA ESQUERDA (7 COLUNAS): VITRINE 3D / ESQUEMA 2D */}
              <div className="lg:col-span-7 flex flex-col">
                <VitrineCaminhao3D
                  veiculo={veiculo3dSelecionado}
                  zonaAtiva={zona3dAtiva}
                  filtroZona3D={filtroZona3D}
                  setFiltroZona3D={setFiltroZona3D}
                />
              </div>

              {/* COLUNA DIREITA (5 COLUNAS): CENTRAL INTERATIVA DE PLACAS, RESUMO CDV E LAUDO */}
              <div className="lg:col-span-5 rounded-2xl border border-gray-800 bg-[#1C2541] overflow-hidden shadow-2xl flex flex-col justify-between">
                {/* ABAS DO PAINEL DIREITO */}
                <div className="p-2.5 bg-[#111A33] border-b border-gray-800 flex items-center justify-between gap-1">
                  <div className="flex items-center gap-1">
                    <button
                      onClick={() => setAbaLateralParados('PLACAS')}
                      className={`px-2.5 py-1.5 rounded-lg text-[10px] font-black uppercase transition-all ${
                        abaLateralParados === 'PLACAS'
                          ? 'bg-cyan-500 text-gray-950 shadow'
                          : 'text-gray-400 hover:text-white'
                      }`}
                    >
                      🚛 Placas ({listaVeiculosParadosFiltrada.length})
                    </button>
                    <button
                      onClick={() => setAbaLateralParados('RESUMO_CDV')}
                      className={`px-2.5 py-1.5 rounded-lg text-[10px] font-black uppercase transition-all ${
                        abaLateralParados === 'RESUMO_CDV'
                          ? 'bg-cyan-500 text-gray-950 shadow'
                          : 'text-gray-400 hover:text-white'
                      }`}
                    >
                      🏢 Resumo CDV ({resumoDispoCalculado.length})
                    </button>
                    <button
                      onClick={() => setAbaLateralParados('LAUDO')}
                      className={`px-2.5 py-1.5 rounded-lg text-[10px] font-black uppercase transition-all ${
                        abaLateralParados === 'LAUDO'
                          ? 'bg-cyan-500 text-gray-950 shadow'
                          : 'text-gray-400 hover:text-white'
                      }`}
                    >
                      🔬 Telemetria ({veiculo3dSelecionado?.placa || '--'})
                    </button>
                  </div>

                  {(filtroCdvParados !== 'TODOS' ||
                    filtroOficinaParados !== 'TODAS' ||
                    filtroZona3D !== 'TODAS') && (
                    <button
                      onClick={() => {
                        setFiltroCdvParados('TODOS');
                        setFiltroOficinaParados('TODAS');
                        setFiltroZona3D('TODAS');
                        setFiltroDispoParados('TODOS');
                      }}
                      className="text-[9px] font-black bg-rose-500/20 text-rose-300 px-2 py-1 rounded"
                    >
                      Limpar ✕
                    </button>
                  )}
                </div>

                {/* CONTEÚDO DA ABA SELECIONADA NO COCKPIT DIREITO */}
                {abaLateralParados === 'PLACAS' && (
                  <div className="flex-1 flex flex-col">
                    <div className="p-2.5 border-b border-gray-800 bg-[#0B1329]/60 flex items-center gap-2">
                      <Search size={14} className="text-gray-400" />
                      <input
                        type="text"
                        value={termoBuscaParados}
                        onChange={(e) => setTermoBuscaParados(e.target.value)}
                        placeholder="Filtrar placa, falha (ex: Motor, Câmbio), oficina ou CDV..."
                        className="bg-transparent w-full outline-none text-xs text-white"
                      />
                    </div>

                    <div className="h-[385px] overflow-y-auto">
                      <table className="w-full text-left text-[11px] table-auto">
                        <thead className="sticky top-0 bg-[#111A33] font-bold text-gray-400 uppercase text-[9px] border-b border-gray-800">
                          <tr>
                            <th className="p-2.5">Placa</th>
                            <th className="p-2.5">CDV / Oficina</th>
                            <th className="p-2.5">Peça 3D</th>
                            <th className="p-2.5">Diagnóstico</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-800">
                          {listaVeiculosParadosFiltrada.length === 0 ? (
                            <tr>
                              <td
                                colSpan={4}
                                className="p-6 text-center text-gray-400 font-bold"
                              >
                                Nenhum veículo. Suba a planilha em "Carregar Excel".
                              </td>
                            </tr>
                          ) : (
                            listaVeiculosParadosFiltrada.map((v, index) => {
                              const zona = identificarZonaDefeito3D(v.status);
                              const isSelected =
                                veiculo3dSelecionado?.placa === v.placa;
                              return (
                                <tr
                                  key={`${v.placa}-${index}`}
                                  onClick={() => setVeiculo3dSelecionado(v)}
                                  className={`cursor-pointer transition-all ${
                                    isSelected
                                      ? 'bg-cyan-500/20 border-l-4 border-cyan-400'
                                      : 'hover:bg-gray-800/30'
                                  }`}
                                >
                                  <td className="p-2.5 font-mono font-black text-white">
                                    {v.placa}
                                    <span
                                      className={`block text-[8px] font-bold uppercase ${
                                        String(v.disponibilidade)
                                          .toUpperCase()
                                          .includes('DISPON')
                                          ? 'text-emerald-400'
                                          : String(v.disponibilidade)
                                              .toUpperCase()
                                              .includes('LIBERA')
                                          ? 'text-amber-400'
                                          : 'text-rose-400'
                                      }`}
                                    >
                                      {v.disponibilidade}
                                    </span>
                                  </td>
                                  <td className="p-2.5 uppercase">
                                    <span className="text-blue-400 font-bold block truncate max-w-[110px]">
                                      {v.cdv}
                                    </span>
                                    <span className="text-[9px] text-gray-400 truncate block max-w-[110px]">
                                      {v.oficina}
                                    </span>
                                  </td>
                                  <td className="p-2.5">
                                    <span
                                      className="text-[9px] font-black px-1.5 py-0.5 rounded border inline-flex items-center gap-1"
                                      style={{
                                        backgroundColor: `${zona.cor}20`,
                                        borderColor: `${zona.cor}50`,
                                        color: '#fff',
                                      }}
                                    >
                                      <span
                                        className="w-1.5 h-1.5 rounded-full"
                                        style={{ backgroundColor: zona.cor }}
                                      />
                                      <span>{zona.id}</span>
                                    </span>
                                  </td>
                                  <td className="p-2.5 font-bold text-amber-300 text-[10px]">
                                    {v.status || 'Em diagnóstico'}
                                  </td>
                                </tr>
                              );
                            })
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                {abaLateralParados === 'RESUMO_CDV' && (
                  <div className="h-[425px] overflow-y-auto">
                    <table className="w-full text-left text-[11px] table-auto">
                      <thead className="sticky top-0 bg-[#0B1329] font-black text-gray-400 uppercase text-[9px] border-b border-gray-800">
                        <tr>
                          <th className="p-2.5">CDV</th>
                          <th className="p-2.5 text-center text-emerald-400">
                            Disp
                          </th>
                          <th className="p-2.5 text-center text-rose-400">
                            Ofic
                          </th>
                          <th className="p-2.5 text-center text-white">Tot</th>
                          <th className="p-2.5 w-36">% Dispo (Farol)</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gray-800">
                        {resumoDispoCalculado.map((r, i) => {
                          const selecionado = filtroCdvParados === r.cdv;
                          const pct = Number(r.pct_dispo) || 0;

                          // REGRA DAS 4 FAIXAS DE CORES:
                          // 100% = Azul | 95% a 99.9% = Verde | 90% a 94.9% = Amarelo | < 90% = Vermelho
                          let corBarra =
                            'bg-gradient-to-r from-red-600 to-rose-500';
                          let corTexto = 'text-red-400';
                          let corLuz = 'bg-red-500 shadow-[0_0_8px_#ef4444]';

                          if (pct >= 100) {
                            corBarra =
                              'bg-gradient-to-r from-blue-500 to-cyan-400';
                            corTexto = 'text-cyan-400';
                            corLuz = 'bg-cyan-400 shadow-[0_0_8px_#22d3ee]';
                          } else if (pct >= 95) {
                            corBarra =
                              'bg-gradient-to-r from-emerald-500 to-green-400';
                            corTexto = 'text-emerald-400';
                            corLuz = 'bg-emerald-400 shadow-[0_0_8px_#34d399]';
                          } else if (pct >= 90) {
                            corBarra =
                              'bg-gradient-to-r from-amber-500 to-yellow-400';
                            corTexto = 'text-yellow-400';
                            corLuz = 'bg-yellow-400 shadow-[0_0_8px_#facc15]';
                          }

                          return (
                            <tr
                              key={i}
                              onClick={() => {
                                setFiltroCdvParados(
                                  selecionado ? 'TODOS' : r.cdv
                                );
                                setAbaLateralParados('PLACAS');
                              }}
                              className={`cursor-pointer transition-all ${
                                selecionado
                                  ? 'bg-cyan-500/20 border-l-4 border-cyan-400'
                                  : 'hover:bg-gray-800/30'
                              }`}
                            >
                              <td className="p-2.5 font-black text-white uppercase">
                                {r.cdv}
                              </td>
                              <td className="p-2.5 text-center font-mono font-bold text-emerald-400">
                                {r.disponivel}
                              </td>
                              <td className="p-2.5 text-center font-mono font-black text-rose-400">
                                {r.em_oficina || 0}
                              </td>
                              <td className="p-2.5 text-center font-mono font-black text-white">
                                {r.total_geral}
                              </td>
                              <td className="p-2.5">
                                <div className="flex items-center gap-2">
                                  <span
                                    className={`w-2.5 h-2.5 rounded-full shrink-0 ${corLuz}`}
                                  />
                                  <div className="flex-1 h-2 bg-[#0B1329] rounded-full overflow-hidden border border-gray-800">
                                    <div
                                      className={`h-full rounded-full ${corBarra}`}
                                      style={{
                                        width: `${Math.min(pct, 100)}%`,
                                      }}
                                    ></div>
                                  </div>
                                  <span
                                    className={`font-mono text-[11px] font-black w-12 text-right ${corTexto}`}
                                  >
                                    {pct.toFixed(1)}%
                                  </span>
                                </div>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}

                {abaLateralParados === 'LAUDO' && (
                  <div className="p-4 space-y-3 flex-1 flex flex-col justify-between">
                    <div className="space-y-3">
                      <div className="flex items-center justify-between border-b border-gray-800 pb-2.5">
                        <div>
                          <span className="text-[10px] font-black text-gray-400 uppercase">
                            Laudo de Telemetria Mecânica
                          </span>
                          <h4 className="text-2xl font-mono font-black text-white">
                            {veiculo3dSelecionado?.placa || 'SEM PLACA'}
                          </h4>
                        </div>
                        {veiculo3dSelecionado && (
                          <span className="text-[10px] font-black px-2.5 py-1 rounded-full uppercase bg-cyan-500/20 text-cyan-300 border border-cyan-500/30">
                            {veiculo3dSelecionado.disponibilidade}
                          </span>
                        )}
                      </div>

                      <div
                        className="p-3 rounded-xl border"
                        style={{
                          backgroundColor: `${zona3dAtiva.cor}15`,
                          borderColor: `${zona3dAtiva.cor}50`,
                        }}
                      >
                        <div className="flex items-center gap-2">
                          <span
                            className="w-6 h-6 rounded-full flex items-center justify-center text-xs font-black text-white shrink-0"
                            style={{ backgroundColor: zona3dAtiva.cor }}
                          >
                            {zona3dAtiva.simbolo}
                          </span>
                          <h5 className="text-xs font-black text-white">
                            {zona3dAtiva.nome}
                          </h5>
                        </div>
                        <p className="text-[11px] text-gray-300 mt-1.5">
                          {zona3dAtiva.descricao}
                        </p>
                      </div>

                      <div className="space-y-2 text-xs">
                        <div className="flex justify-between p-2.5 rounded-lg bg-[#0B1329]">
                          <span className="text-gray-400 font-bold">Falha:</span>
                          <span className="font-black text-amber-400 text-right">
                            {veiculo3dSelecionado?.status || '--'}
                          </span>
                        </div>
                        <div className="flex justify-between p-2.5 rounded-lg bg-[#0B1329]">
                          <span className="text-gray-400 font-bold">
                            CDV / Filial:
                          </span>
                          <span className="font-black text-blue-400">
                            {veiculo3dSelecionado?.cdv} (
                            {veiculo3dSelecionado?.filial})
                          </span>
                        </div>
                        <div className="flex justify-between p-2.5 rounded-lg bg-[#0B1329]">
                          <span className="text-gray-400 font-bold">
                            Oficina:
                          </span>
                          <span className="font-bold text-white">
                            {veiculo3dSelecionado?.oficina || '--'}
                          </span>
                        </div>
                        <div className="flex justify-between p-2.5 rounded-lg bg-[#0B1329]">
                          <span className="text-gray-400 font-bold">
                            OS / Valor:
                          </span>
                          <span className="font-mono font-black text-rose-400">
                            {veiculo3dSelecionado?.os || '--'} •{' '}
                            {veiculo3dSelecionado?.valor > 0
                              ? formatarMoeda(veiculo3dSelecionado.valor)
                              : 'Ag. Orçamento'}
                          </span>
                        </div>
                      </div>
                    </div>
                  </div>
                )}

                {/* RODAPÉ RÁPIDO DO PAINEL DIREITO COM MINI-RESUMO DA PLACA ATIVA */}
                {veiculo3dSelecionado && abaLateralParados !== 'LAUDO' && (
                  <div
                    onClick={() => setAbaLateralParados('LAUDO')}
                    className="p-2.5 bg-[#0B1329] border-t border-gray-800 flex items-center justify-between cursor-pointer hover:bg-gray-900 transition-all"
                  >
                    <div className="flex items-center gap-2">
                      <span
                        className="w-2.5 h-2.5 rounded-full shrink-0"
                        style={{ backgroundColor: zona3dAtiva.cor }}
                      />
                      <span className="font-mono font-black text-xs text-yellow-400">
                        {veiculo3dSelecionado.placa}
                      </span>
                      <span className="text-[10px] font-bold text-gray-300 truncate max-w-[170px]">
                        • {veiculo3dSelecionado.status}
                      </span>
                    </div>
                    <span className="text-[10px] font-black text-cyan-400">
                      Ver Laudo ▸
                    </span>
                  </div>
                )}
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}

export default App;
