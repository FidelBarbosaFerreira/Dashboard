import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  ShieldCheck,
  FileSpreadsheet,
  Upload,
  Search,
  TrendingUp,
  TrendingDown,
  UserCheck,
  RotateCcw,
  LayoutDashboard,
  ClipboardCheck,
  ArrowLeft,
  ArrowRight,
  Sparkles,
  RefreshCw,
  BarChart3,
  FileText,
  CheckCircle2,
  ChevronDown,
  ChevronUp,
  FilterX,
  AlertTriangle,
  Award,
  CreditCard,
  Database,
  Car,
  Download,
  Users,
  MapPin,
} from 'lucide-react';

// 🔗 LINKS DIRETO DAS PLANILHAS NO SEU ONEDRIVE (PEPSICO)
const URL_SINISTROS_NUVEM =
  'https://pepsico-my.sharepoint.com/personal/fidel_barbosa_contractor_pepsico_com/Documents/PROJETO_CONTROL/Sinistros.xlsx?download=1';

const URL_FORMS_NUVEM =
  'https://pepsico-my.sharepoint.com/personal/fidel_barbosa_contractor_pepsico_com/Documents/PROJETO_CONTROL/CERTIFICA%C3%87%C3%83O%20DIRE%C3%87%C3%83O%20SEGURA.xlsx?download=1';

const URL_POSICAO_NUVEM =
  'https://pepsico-my.sharepoint.com/personal/fidel_barbosa_contractor_pepsico_com/Documents/PROJETO_CONTROL/Posicao_Nacional.xlsx?download=1';

export default function App() {
  // --- NAVEGAÇÃO DE TELAS: 'menu' | 'sinistros' | 'forms' | 'frota_leves' ---
  const [telaAtiva, setTelaAtiva] = useState('menu');

  // --- ESTADOS MÓDULO 1: SINISTROS ---
  const [dadosBaseSinistros, setDadosBaseSinistros] = useState([]);
  const [mapaPosicaoNacional, setMapaPosicaoNacional] = useState({});
  const [qtdPlacasPosicao, setQtdPlacasPosicao] = useState(0);
  const [carregadoSinistros, setCarregadoSinistros] = useState(false);
  const [carregandoNuvemSinistros, setCarregandoNuvemSinistros] =
    useState(false);
  const [anoFocal, setAnoFocal] = useState(2026);
  const [filtroFrota, setFiltroFrota] = useState('TODAS');
  const [filtroCategoriaVeiculo, setFiltroCategoriaVeiculo] =
    useState('PESADO');
  const [filtroCulpabilidade, setFiltroCulpabilidade] = useState('TODAS');
  const [filtroSetor, setFiltroSetor] = useState('TODOS');
  const [apenasCDVsNoRanking, setApenasCDVsNoRanking] = useState(true);
  const [termoBuscaSinistros, setTermoBuscaSinistros] = useState('');
  const [mesFiltroKpi, setMesFiltroKpi] = useState(new Date().getMonth() + 1);

  // --- ESTADOS MÓDULO 2: FORMS (CERTIFICAÇÃO DIREÇÃO SEGURA) ---
  const [dadosBaseForms, setDadosBaseForms] = useState([]);
  const [colunasQuestoesForms, setColunasQuestoesForms] = useState([]);
  const [carregadoForms, setCarregadoForms] = useState(false);
  const [carregandoNuvemForms, setCarregandoNuvemForms] = useState(false);
  const [filtroAnoForms, setFiltroAnoForms] = useState('TODOS');
  const [filtroMesForms, setFiltroMesForms] = useState('TODOS');
  const [filtroCanalForms, setFiltroCanalForms] = useState('TODOS');
  const [filtroDivisaoForms, setFiltroDivisaoForms] = useState('TODAS');
  const [filtroTipoCertForms, setFiltroTipoCertForms] = useState('TODOS');
  const [filtroAuditoriaForms, setFiltroAuditoriaForms] = useState('TODOS');
  const [questaoDesvioSelecionada, setQuestaoDesvioSelecionada] =
    useState(null);
  const [termoBuscaForms, setTermoBuscaForms] = useState('');
  const [expandirPerguntasForms, setExpandirPerguntasForms] = useState(false);

  // --- ESTADOS MÓDULO 3: FROTA LEVE ATIVA (DTS LEVE vs OT) ---
  const [dadosFrotaAtiva, setDadosFrotaAtiva] = useState([]);
  const [carregadoFrotaAtiva, setCarregadoFrotaAtiva] = useState(false);
  const [carregandoNuvemPosicao, setCarregandoNuvemPosicao] = useState(false);
  const [filtroSegmentoLeve, setFiltroSegmentoLeve] = useState('TODOS'); // 'TODOS' | 'OT' | 'DTS LEVE' | 'SUPORTE'
  const [filtroAreaLeve, setFiltroAreaLeve] = useState('TODAS');
  const [filtroGerenteLeve, setFiltroGerenteLeve] = useState('TODOS');
  const [filtroTipoFrotaLeve, setFiltroTipoFrotaLeve] = useState('OPERACIONAL'); // 'OPERACIONAL' (783) | 'COM_EXECUTIVO' (1087)
  const [incluirTranspWhsNoDts, setIncluirTranspWhsNoDts] = useState(false);
  const [termoBuscaLeves, setTermoBuscaLeves] = useState('');

  const anoComparacao = anoFocal === 2026 ? 2025 : 2026;

  // 🎯 CARREGA OS DADOS AUTOMATICAMENTE AO ABRIR OS MÓDULOS
  useEffect(() => {
    if (telaAtiva === 'forms' && !carregadoForms) {
      carregarFormsNuvem();
    }
    if (telaAtiva === 'sinistros' && !carregadoSinistros) {
      carregarSinistrosNuvem();
    }
    if (telaAtiva === 'frota_leves' && !carregadoFrotaAtiva) {
      carregarPosicaoNuvem();
    }
  }, [telaAtiva]);

  // 🔤 FUNÇÃO DE NORMALIZAÇÃO (REMOVE ACENTOS E ESPAÇOS DUPLOS)
  function normalizarTexto(str) {
    return String(str || '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/\u00a0/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toUpperCase();
  }

  function padronizarFilial(nomeFilial) {
    let f = String(nomeFilial || 'SEM INFORMAÇÃO')
      .replace(/\u00a0/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toUpperCase();

    if (f === 'CDV MEGA RIO') return 'CDV MEGA RJ';
    if (f === '-' || f === 'NAN' || f === 'SEM INFORMACAO' || f === '')
      return 'SEM INFORMAÇÃO';
    return f;
  }

  function extrairDataCompletaExcel(raw) {
    if (!raw || String(raw).trim() === '')
      return { ano: null, mes: null, formatada: 'N/I', dateObj: null };

    if (typeof raw === 'number') {
      const UTC_DAYS = Math.floor(raw - 25569);
      const dataObj = new Date(UTC_DAYS * 86400 * 1000);
      if (!isNaN(dataObj.getTime())) {
        const dia = String(dataObj.getUTCDate()).padStart(2, '0');
        const mes = dataObj.getUTCMonth() + 1;
        const ano = dataObj.getUTCFullYear();
        return {
          ano,
          mes,
          formatada: `${dia}/${String(mes).padStart(2, '0')}/${ano}`,
          dateObj: dataObj,
        };
      }
    }

    const strData = String(raw).trim();
    if (strData.includes('-')) {
      const parteData = strData.split(' ')[0].split('T')[0];
      const partes = parteData.split('-');
      if (partes.length === 3) {
        const ano = parseInt(partes[0], 10);
        const mes = parseInt(partes[1], 10);
        const dia = parseInt(partes[2], 10);
        if (ano > 1990 && mes >= 1 && mes <= 12) {
          return {
            ano,
            mes,
            formatada: `${String(dia).padStart(2, '0')}/${String(mes).padStart(
              2,
              '0'
            )}/${ano}`,
            dateObj: new Date(ano, mes - 1, dia),
          };
        }
      }
    }

    if (strData.includes('/')) {
      const parteData = strData.split(' ')[0];
      const partes = parteData.split('/');
      if (partes.length === 3) {
        const p1 = parseInt(partes[0], 10);
        const p2 = parseInt(partes[1], 10);
        let ano = parseInt(partes[2], 10);
        if (ano < 100) ano += 2000;
        let dia = p1;
        let mes = p2;
        if (p1 <= 12 && p2 > 12) {
          mes = p1;
          dia = p2;
        }
        if (ano > 1990 && mes >= 1 && mes <= 12) {
          return {
            ano,
            mes,
            formatada: `${String(dia).padStart(2, '0')}/${String(mes).padStart(
              2,
              '0'
            )}/${ano}`,
            dateObj: new Date(ano, mes - 1, dia),
          };
        }
      }
    }

    const dObj = new Date(strData);
    if (!isNaN(dObj.getTime())) {
      const dia = String(dObj.getDate()).padStart(2, '0');
      const mes = dObj.getMonth() + 1;
      const ano = dObj.getFullYear();
      return {
        ano,
        mes,
        formatada: `${dia}/${String(mes).padStart(2, '0')}/${ano}`,
        dateObj: dObj,
      };
    }

    return { ano: null, mes: null, formatada: strData, dateObj: null };
  }

  // 🧠 REGRA OFICIAL PARA SEPARAR "DTS LEVE" vs "OT" vs "SUPORTE" NA POSIÇÃO NACIONAL
  function classificarSegmentoLeve(
    diretoriaRaw,
    descCcustoRaw,
    frotaRaw,
    incluirTranspWhs = false
  ) {
    const dir = normalizarTexto(diretoriaRaw);
    const cc = normalizarTexto(descCcustoRaw);
    const frota = normalizarTexto(frotaRaw);

    if (frota === 'EXECUTIVO' || frota === 'EXECUTIVA') {
      return 'EXECUTIVO';
    }

    // 1. Diretorias que pertencem 100% ao DTS LEVE (Vendas DTS / Entregadores Leves / Apoio CDV)
    if (
      dir === 'DTS' ||
      dir.includes('HYPER TRANSFORMATION') ||
      dir.includes('TRANSF. & EXECUTION') ||
      dir.includes('TRANSF & EXECUTION')
    ) {
      return 'DTS LEVE';
    }

    // 2. Diretorias Regionais de Vendas OT (Organized Trade / Key Accounts / Merchandising)
    if (
      dir === 'SULESTE' ||
      dir === 'SP + CO' ||
      dir === 'NENO' ||
      dir === 'NATIONAL ONE' ||
      dir === 'SUL + CO' ||
      dir === 'SP ESTADO' ||
      dir === 'LESTE + KA' ||
      dir === 'LESTE' ||
      dir.includes('GROWTH DIGITAL') ||
      dir.includes('VP VENDAS') ||
      dir.includes('COMMERCIAL')
    ) {
      return 'OT';
    }

    // 3. Caso opcional: se a gestão quiser somar os leves de Transportes/Warehouse dos CDVs dentro de DTS Leve
    if (
      incluirTranspWhs &&
      (dir === 'TRANSPORTES' || dir === 'WAREHOUSE')
    ) {
      return 'DTS LEVE';
    }

    // 4. Verificação secundária pelo Centro de Custo caso a Diretoria esteja genérica
    if (
      cc.includes(' OT') ||
      cc.startsWith('OT ') ||
      cc.includes('MERCHAN') ||
      cc.includes('KAN ') ||
      cc.includes('KAM ') ||
      cc.includes('WAL MART') ||
      cc.includes('CARREFOUR') ||
      cc.includes('CASH & CARRY') ||
      cc.includes('ATACADO')
    ) {
      return 'OT';
    }

    if (cc.includes('DTS') || cc.includes('ENTREGADORES')) {
      return 'DTS LEVE';
    }

    // 5. Áreas de Suporte (EHS, Agro, Security, Qualidade, Transportes, Warehouse, etc.)
    return 'SUPORTE / OUTROS';
  }

  // 🔄 CARREGAMENTO AUTOMÁTICO DO ONEDRIVE - SINISTROS, FORMS E POSIÇÃO NACIONAL
  const carregarSinistrosNuvem = async () => {
    setCarregandoNuvemSinistros(true);
    try {
      const resposta = await fetch(URL_SINISTROS_NUVEM);
      if (!resposta.ok) throw new Error('Falha no download da nuvem');

      const dadosArrayBuffer = await resposta.arrayBuffer();
      const livro = XLSX.read(dadosArrayBuffer, { type: 'array' });
      processarDadosLivroSinistros(livro);
    } catch (err) {
      console.log(
        'Leitura automática da nuvem (Sinistros) aguardando upload local:',
        err
      );
    } finally {
      setCarregandoNuvemSinistros(false);
    }
  };

  const carregarFormsNuvem = async () => {
    setCarregandoNuvemForms(true);
    try {
      const resposta = await fetch(URL_FORMS_NUVEM);
      if (!resposta.ok)
        throw new Error('Falha no download da planilha Forms do OneDrive');

      const dadosArrayBuffer = await resposta.arrayBuffer();
      const livro = XLSX.read(dadosArrayBuffer, { type: 'array' });
      const primeiraAba = livro.Sheets[livro.SheetNames[0]];
      const registrosBrutos = XLSX.utils.sheet_to_json(primeiraAba, {
        defval: '',
      });

      if (registrosBrutos && registrosBrutos.length > 0) {
        processarDadosFormsBrutos(registrosBrutos);
      }
    } catch (err) {
      console.log(
        'Leitura automática do OneDrive Forms aguardando upload local:',
        err
      );
    } finally {
      setCarregandoNuvemForms(false);
    }
  };

  const carregarPosicaoNuvem = async () => {
    setCarregandoNuvemPosicao(true);
    try {
      const resposta = await fetch(URL_POSICAO_NUVEM);
      if (!resposta.ok)
        throw new Error('Falha no download da Posição Nacional na nuvem');

      const dadosArrayBuffer = await resposta.arrayBuffer();
      const livro = XLSX.read(dadosArrayBuffer, { type: 'array' });
      processarLivroPosicaoNacional(livro);
    } catch (err) {
      console.log(
        'Posição Nacional aguardando upload manual (.xlsx):',
        err
      );
    } finally {
      setCarregandoNuvemPosicao(false);
    }
  };

  // 🔗 PROCESSAMENTO COMPLETO DA PLANILHA "POSIÇÃO NACIONAL FROTA" (ABA ATIVOS + CRUZAMENTO)
  const processarLivroPosicaoNacional = (livro) => {
    const novoMapa = {};
    const listaAtivosLeves = [];

    livro.SheetNames.forEach((nomeAba) => {
      const nUp = normalizarTexto(nomeAba);
      const ehAbaAtivos = nUp === 'ATIVOS';

      if (
        ehAbaAtivos ||
        nUp.includes('INATIVO') ||
        nUp.includes('BAIXADO') ||
        nUp.includes('POSICAO')
      ) {
        const folha = livro.Sheets[nomeAba];
        const matriz = XLSX.utils.sheet_to_json(folha, {
          header: 1,
          defval: '',
        });

        let linhaCabecalho = -1;
        for (let i = 0; i < Math.min(matriz.length, 20); i++) {
          const rowStr = matriz[i].map((c) => normalizarTexto(c));
          if (rowStr.includes('PLACA') && rowStr.includes('FROTA')) {
            linhaCabecalho = i;
            break;
          }
        }

        if (linhaCabecalho !== -1) {
          const cabecalhos = matriz[linhaCabecalho].map((c) =>
            normalizarTexto(c)
          );
          const idxCodigo = cabecalhos.findIndex((c) => c.includes('CODIGO'));
          const idxPlaca = cabecalhos.indexOf('PLACA');
          const idxModelo = cabecalhos.indexOf('MODELO');
          const idxFrota = cabecalhos.indexOf('FROTA');
          const idxLocadora = cabecalhos.indexOf('LOCADORA');
          const idxUf = cabecalhos.indexOf('UF');
          const idxAno = cabecalhos.indexOf('MOD');
          const idxCcustoSap = cabecalhos.findIndex((c) =>
            c.includes('CCUSTO SAP')
          );
          const idxDiret = cabecalhos.findIndex((c) => c.includes('DIRET'));
          const idxDescCcusto = cabecalhos.findIndex((c) =>
            c.includes('DESC. CCUSTO') || c.includes('DESC CCUSTO')
          );
          const idxCdv = cabecalhos.findIndex((c) => c.includes('DESC. CDV'));
          const idxArea = cabecalhos.findIndex((c) => c === 'AREA');
          const idxGerente = cabecalhos.findIndex(
            (c) => c.includes('GERENTE') || c.includes('COORDENADOR')
          );
          const idxGpid = cabecalhos.indexOf('GPID');
          const idxUsuario = cabecalhos.findIndex((c) =>
            c.includes('USUARIO')
          );
          const idxStatus = cabecalhos.indexOf('STATUS');

          for (let r = linhaCabecalho + 1; r < matriz.length; r++) {
            const row = matriz[r];
            const placa = String(row[idxPlaca] || '')
              .trim()
              .toUpperCase();
            if (!placa || placa.length < 5 || placa === '-') continue;

            const codigo =
              idxCodigo !== -1 ? String(row[idxCodigo] || '').trim() : '';
            let modelo =
              idxModelo !== -1
                ? String(row[idxModelo] || '')
                    .replace(/\s+/g, ' ')
                    .trim()
                    .toUpperCase()
                : '';

            // Corrige erro conhecido da planilha onde a placa FXH7F49 foi digitada no lugar do modelo
            let alertaErroCadastro = '';
            if (modelo === placa || /^[A-Z]{3}\d[A-Z0-9]\d{2}$/.test(modelo)) {
              alertaErroCadastro = `Modelo preenchido com placa (${modelo})`;
              modelo = 'VEÍCULO LEVE (MODELO N/I)';
            }

            const frota =
              idxFrota !== -1
                ? String(row[idxFrota] || '')
                    .replace(/\s+/g, ' ')
                    .trim()
                    .toUpperCase()
                : '';
            const diretoria =
              idxDiret !== -1
                ? String(row[idxDiret] || '')
                    .replace(/\s+/g, ' ')
                    .trim()
                    .toUpperCase()
                : '';
            const descCcusto =
              idxDescCcusto !== -1
                ? String(row[idxDescCcusto] || '')
                    .replace(/\s+/g, ' ')
                    .trim()
                    .toUpperCase()
                : '';
            const cdv =
              idxCdv !== -1
                ? padronizarFilial(row[idxCdv])
                : 'SEM INFORMAÇÃO';
            const area =
              idxArea !== -1
                ? String(row[idxArea] || '')
                    .trim()
                    .toUpperCase()
                : 'N/I';
            const gerente =
              idxGerente !== -1
                ? String(row[idxGerente] || '')
                    .replace(/\s+/g, ' ')
                    .trim()
                    .toUpperCase()
                : 'N/I';
            const locadora =
              idxLocadora !== -1
                ? String(row[idxLocadora] || '')
                    .trim()
                    .toUpperCase()
                : 'N/I';
            const uf =
              idxUf !== -1 ? String(row[idxUf] || '').trim().toUpperCase() : '';
            const anoMod =
              idxAno !== -1 ? String(row[idxAno] || '').trim() : '';
            const ccustoSap =
              idxCcustoSap !== -1
                ? String(row[idxCcustoSap] || '').trim()
                : '';
            const gpid =
              idxGpid !== -1 ? String(row[idxGpid] || '').trim() : '';
            const usuario =
              idxUsuario !== -1
                ? String(row[idxUsuario] || '')
                    .replace(/\s+/g, ' ')
                    .trim()
                : '';
            const status =
              idxStatus !== -1
                ? String(row[idxStatus] || '')
                    .trim()
                    .toUpperCase()
                : '-';

            if (!novoMapa[placa]) {
              novoMapa[placa] = {
                diretoria,
                frota,
                modelo,
                area,
                cdv,
                descCcusto,
              };
            }

            // Se for da aba ATIVOS e for da Frota Leve (Locados, ALD Flex, Temporários ou Executivo)
            if (ehAbaAtivos) {
              const frotaNorm = normalizarTexto(frota);
              const ehLeveOperacional =
                frotaNorm.includes('LOCADOS') ||
                frotaNorm.includes('ALD FLEX') ||
                frotaNorm.includes('TEMPORARIO');
              const ehExecutivo = frotaNorm.includes('EXECUTIV');

              if (ehLeveOperacional || ehExecutivo) {
                listaAtivosLeves.push({
                  id: r,
                  codigo,
                  placa,
                  modelo,
                  frota,
                  ehExecutivo,
                  locadora,
                  uf,
                  anoMod,
                  ccustoSap,
                  diretoriaOriginal: diretoria || 'NÃO INFORMADA',
                  descCcusto: descCcusto || 'NÃO INFORMADO',
                  cdv,
                  area,
                  gerente,
                  gpid: gpid === '-' ? '' : gpid,
                  usuario: usuario === '-' ? 'SEM CONDUTOR FIXO' : usuario,
                  status: status === '-' ? 'ATIVO EM OPERAÇÃO' : status,
                  alertaErroCadastro,
                });
              }
            }
          }
        }
      }
    });

    const totalMapeadas = Object.keys(novoMapa).length;
    if (totalMapeadas > 0) {
      setMapaPosicaoNacional(novoMapa);
      setQtdPlacasPosicao(totalMapeadas);
      setDadosFrotaAtiva(listaAtivosLeves);
      setCarregadoFrotaAtiva(true);

      // Atualiza também o módulo de Sinistros caso já esteja carregado
      setDadosBaseSinistros((prev) =>
        prev.map((item) => {
          const infoPos = novoMapa[item.placa];
          if (!infoPos) return item;
          const veicFinal = item.veiculo || infoPos.modelo;
          const catFinal = classificarTipoVeiculo(veicFinal);
          const frotaFinal = infoPos.frota
            ? infoPos.frota.toUpperCase()
            : item.tipoFrota;
          const setorFinal = deduzirSetorOperacional(
            frotaFinal,
            item.divisao,
            item.filial,
            catFinal,
            infoPos.diretoria
          );
          return {
            ...item,
            veiculo: veicFinal,
            categoriaVeiculo: catFinal,
            tipoFrota: frotaFinal,
            setor: setorFinal,
          };
        })
      );
    } else {
      alert('Nenhuma placa encontrada nas abas da Posição Nacional.');
    }
  };

  const lidarComUploadPosicaoNacional = (e) => {
    const arquivo = e.target.files[0];
    if (!arquivo) return;
    const leitor = new FileReader();
    leitor.onload = (evento) => {
      try {
        const dadosBinarios = evento.target.result;
        const livro = XLSX.read(dadosBinarios, { type: 'binary' });
        processarLivroPosicaoNacional(livro);
      } catch (err) {
        alert('Erro ao ler Posição Nacional: ' + err.message);
      }
    };
    leitor.readAsBinaryString(arquivo);
  };

  // --- LÓGICA E CÁLCULOS DO MÓDULO 3: FROTA LEVE ATIVA (DTS LEVE vs OT) ---
  const frotaLeveProcessada = useMemo(() => {
    return dadosFrotaAtiva
      .filter((v) =>
        filtroTipoFrotaLeve === 'OPERACIONAL' ? !v.ehExecutivo : true
      )
      .map((v) => ({
        ...v,
        segmentoOficial: classificarSegmentoLeve(
          v.diretoriaOriginal,
          v.descCcusto,
          v.frota,
          incluirTranspWhsNoDts
        ),
      }));
  }, [dadosFrotaAtiva, filtroTipoFrotaLeve, incluirTranspWhsNoDts]);

  const opcoesAreaLeve = useMemo(() => {
    const s = new Set(
      frotaLeveProcessada.map((v) => v.area).filter((a) => a && a !== 'N/I')
    );
    return Array.from(s).sort();
  }, [frotaLeveProcessada]);

  const opcoesGerenteLeve = useMemo(() => {
    const s = new Set(
      frotaLeveProcessada
        .map((v) => v.gerente)
        .filter((g) => g && g !== 'N/I')
    );
    return Array.from(s).sort();
  }, [frotaLeveProcessada]);

  // Base filtrada por Área, Gerente e Busca (para manter os contadores de OT e DTS Leve visíveis nos cards)
  const baseLevesSemFiltroSegmento = useMemo(() => {
    return frotaLeveProcessada.filter((v) => {
      if (filtroAreaLeve !== 'TODAS' && v.area !== filtroAreaLeve) return false;
      if (filtroGerenteLeve !== 'TODOS' && v.gerente !== filtroGerenteLeve)
        return false;
      if (termoBuscaLeves) {
        const b = normalizarTexto(termoBuscaLeves);
        return (
          normalizarTexto(v.placa).includes(b) ||
          normalizarTexto(v.codigo).includes(b) ||
          normalizarTexto(v.modelo).includes(b) ||
          normalizarTexto(v.diretoriaOriginal).includes(b) ||
          normalizarTexto(v.descCcusto).includes(b) ||
          normalizarTexto(v.cdv).includes(b) ||
          normalizarTexto(v.usuario).includes(b) ||
          normalizarTexto(v.gpid).includes(b)
        );
      }
      return true;
    });
  }, [
    frotaLeveProcessada,
    filtroAreaLeve,
    filtroGerenteLeve,
    termoBuscaLeves,
  ]);

  const levesFiltradosFinal = useMemo(() => {
    return baseLevesSemFiltroSegmento.filter((v) => {
      if (filtroSegmentoLeve === 'TODOS') return true;
      return v.segmentoOficial === filtroSegmentoLeve;
    });
  }, [baseLevesSemFiltroSegmento, filtroSegmentoLeve]);

  // Contadores exatos de OT vs DTS Leve vs Suporte
  const resumoSegmentosLeves = useMemo(() => {
    const c = {
      OT: 0,
      'DTS LEVE': 0,
      'SUPORTE / OUTROS': 0,
      EXECUTIVO: 0,
    };
    baseLevesSemFiltroSegmento.forEach((v) => {
      if (c[v.segmentoOficial] !== undefined) c[v.segmentoOficial]++;
      else c['SUPORTE / OUTROS']++;
    });
    return c;
  }, [baseLevesSemFiltroSegmento]);

  // Comparativo por Regional / Área (DTS Leve vs OT)
  const comparativoRegionalLeves = useMemo(() => {
    const mapa = {};
    baseLevesSemFiltroSegmento.forEach((v) => {
      const ar = v.area || 'OUTROS';
      if (!mapa[ar]) {
        mapa[ar] = { area: ar, dtsLeve: 0, ot: 0, outros: 0, total: 0 };
      }
      if (v.segmentoOficial === 'DTS LEVE') mapa[ar].dtsLeve++;
      else if (v.segmentoOficial === 'OT') mapa[ar].ot++;
      else mapa[ar].outros++;
      mapa[ar].total++;
    });
    return Object.values(mapa).sort((a, b) => b.total - a.total);
  }, [baseLevesSemFiltroSegmento]);

  // Comparativo por Gerente Regional
  const comparativoGerentesLeves = useMemo(() => {
    const mapa = {};
    baseLevesSemFiltroSegmento.forEach((v) => {
      const ger = v.gerente || 'NÃO INFORMADO';
      if (!mapa[ger]) {
        mapa[ger] = { gerente: ger, dtsLeve: 0, ot: 0, outros: 0, total: 0 };
      }
      if (v.segmentoOficial === 'DTS LEVE') mapa[ger].dtsLeve++;
      else if (v.segmentoOficial === 'OT') mapa[ger].ot++;
      else mapa[ger].outros++;
      mapa[ger].total++;
    });
    return Object.values(mapa).sort((a, b) => b.total - a.total);
  }, [baseLevesSemFiltroSegmento]);

  // 📥 FUNÇÃO PARA EXPORTAR A BASE HIGIENIZADA DE LEVES (DTS LEVE E OT) PARA EXCEL (.XLSX)
  const exportarBaseLevesExcel = () => {
    if (levesFiltradosFinal.length === 0) {
      alert('Não há veículos filtrados para exportar.');
      return;
    }

    const linhasExportacao = levesFiltradosFinal.map((v) => ({
      'CLASSIFICAÇÃO OFICIAL': v.segmentoOficial,
      CÓDIGO: v.codigo,
      PLACA: v.placa,
      MODELO: v.modelo,
      'TIPO DE FROTA': v.frota,
      LOCADORA: v.locadora,
      UF: v.uf,
      'ANO MODELO': v.anoMod,
      'DIRETORIA (ORIGINAL)': v.diretoriaOriginal,
      'CCUSTO SAP': v.ccustoSap,
      'DESCRIÇÃO CENTRO DE CUSTO': v.descCcusto,
      'CDV / UNIDADE': v.cdv,
      'ÁREA / REGIONAL': v.area,
      'GERENTE REGIONAL': v.gerente,
      'GPID CONDUTOR': v.gpid,
      'USUÁRIO / CONDUTOR': v.usuario,
      STATUS: v.status,
      'AUDITORIA CADASTRO': v.alertaErroCadastro || 'OK',
    }));

    const novaFolha = XLSX.utils.json_to_sheet(linhasExportacao);
    const novoLivro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(novoLivro, novaFolha, 'Base Leves DTS e OT');
    XLSX.writeFile(
      novoLivro,
      `Base_Leves_Ativos_${filtroSegmentoLeve.replace(/\s+/g, '_')}_2026.xlsx`
    );
  };

  // --- PROCESSAMENTO DO MÓDULO 2: FORMS ---
  const processarDadosFormsBrutos = (registros) => {
    if (!registros || registros.length === 0) return;

    const todasChaves = Object.keys(registros[0]);

    const chavesPerguntas = todasChaves.filter((k) => {
      const limpo = k.replace(/\u00a0/g, ' ').trim();
      return /^\d+\s*[-–.]/.test(limpo);
    });
    setColunasQuestoesForms(chavesPerguntas);

    const acharValor = (obj, termos) => {
      const chave = Object.keys(obj).find((k) => {
        const kUp = k
          .replace(/\u00a0/g, ' ')
          .trim()
          .toUpperCase();
        return termos.some((t) => kUp === t || kUp.includes(t));
      });
      return chave ? obj[chave] : '';
    };

    const hoje = new Date();

    const processados = registros.map((r, idx) => {
      const linhaLimpa = {};
      Object.keys(r).forEach((k) => {
        const val = r[k];
        if (
          val === null ||
          val === undefined ||
          String(val).trim() === '' ||
          String(val).toUpperCase() === 'NULL'
        ) {
          linhaLimpa[k] = 'NÃO RESPONDIDO';
        } else {
          linhaLimpa[k] = String(val)
            .replace(/\u00a0/g, ' ')
            .trim();
        }
      });

      const tipoCert = String(
        acharValor(r, ['TIPO DE CERTIFICAÇÃO', 'TIPO DE CERTIFICACAO']) ||
          'NÃO INFORMADO'
      )
        .replace(/\u00a0/g, ' ')
        .trim()
        .toUpperCase();

      const gpidAvaliador = String(
        acharValor(r, ['GPID DO AVALIADOR']) || ''
      ).trim();

      const condutor = String(acharValor(r, ['CONDUTOR']) || 'NÃO INFORMADO')
        .replace(/\u00a0/g, ' ')
        .trim();

      const gpidCondutor = String(
        acharValor(r, ['GPID DO CONDUTOR']) || ''
      ).trim();

      const canal = String(acharValor(r, ['CANAL']) || 'OUTROS')
        .replace(/\u00a0/g, ' ')
        .trim()
        .toUpperCase();

      const divDTS = String(acharValor(r, ['DIVISÃO DTS', 'DIVISAO DTS']) || '')
        .replace(/\u00a0/g, ' ')
        .trim()
        .toUpperCase();

      const divOT = String(acharValor(r, ['DIVISÃO OT', 'DIVISAO OT']) || '')
        .replace(/\u00a0/g, ' ')
        .trim()
        .toUpperCase();

      let divisaoUnificada = 'NÃO INFORMADA';
      if (canal === 'DTS' && divDTS) divisaoUnificada = divDTS;
      else if (canal === 'OT' && divOT) divisaoUnificada = divOT;
      else if (divOT) divisaoUnificada = divOT;
      else if (divDTS) divisaoUnificada = divDTS;

      const rawDataCert = acharValor(r, [
        'DATA DA CERTIFICAÇÃO',
        'DATA DA CERTIFICACAO',
      ]);
      const infoDataCert = extrairDataCompletaExcel(rawDataCert);

      const rawDataCNH = acharValor(r, ['DATA VALIDADE CNH']);
      const infoDataCNH = extrairDataCompletaExcel(rawDataCNH);

      let cnhAlerta = false;
      if (infoDataCNH.dateObj && infoDataCNH.dateObj < hoje) {
        cnhAlerta = true;
      }

      const comentarios = String(
        acharValor(r, ['COMENTÁRIOS', 'COMENTARIOS']) || ''
      )
        .replace(/\u00a0/g, ' ')
        .trim();

      const listaDesvios = [];
      chavesPerguntas.forEach((qCol) => {
        const resp = String(linhaLimpa[qCol] || '').toUpperCase();
        if (
          resp.includes('NÃO EFICI') ||
          resp.includes('NAO EFICI') ||
          resp.includes('NÃO CONFORME') ||
          resp.includes('NAO CONFORME')
        ) {
          listaDesvios.push(qCol);
        }
      });

      return {
        ...linhaLimpa,
        _id: idx + 1,
        _tipoCert: tipoCert,
        _gpidAvaliador: gpidAvaliador,
        _condutor: condutor,
        _gpidCondutor: gpidCondutor,
        _canal: canal,
        _divisao: divisaoUnificada,
        _divDTS: divDTS || 'N/A',
        _divOT: divOT || 'N/A',
        _dataCertFormatada: infoDataCert.formatada,
        _anoCert: infoDataCert.ano === 2032 ? 2026 : infoDataCert.ano || 2026,
        _mesCert: infoDataCert.mes || 1,
        _dataCNHFormatada: infoDataCNH.formatada,
        _cnhAlerta: cnhAlerta,
        _comentarios: comentarios,
        _qtdDesvios: listaDesvios.length,
        _listaDesvios: listaDesvios,
        _conformeGeral: listaDesvios.length === 0,
      };
    });

    setDadosBaseForms(processados);
    setCarregadoForms(true);
  };

  const lidarComUploadForms = (e) => {
    const arquivo = e.target.files[0];
    if (!arquivo) return;
    const leitor = new FileReader();
    leitor.onload = (evento) => {
      try {
        const dadosBinarios = evento.target.result;
        const livro = XLSX.read(dadosBinarios, { type: 'binary' });
        const primeiraAba = livro.Sheets[livro.SheetNames[0]];
        const registrosBrutos = XLSX.utils.sheet_to_json(primeiraAba, {
          defval: '',
        });

        processarDadosFormsBrutos(registrosBrutos);
      } catch (err) {
        alert('Erro ao processar planilha do Forms: ' + err.message);
      }
    };
    leitor.readAsBinaryString(arquivo);
  };

  const limparFiltrosForms = () => {
    setFiltroAnoForms('TODOS');
    setFiltroMesForms('TODOS');
    setFiltroCanalForms('TODOS');
    setFiltroDivisaoForms('TODAS');
    setFiltroTipoCertForms('TODOS');
    setFiltroAuditoriaForms('TODOS');
    setQuestaoDesvioSelecionada(null);
    setTermoBuscaForms('');
  };

  const opcoesDivisaoForms = useMemo(() => {
    const base =
      filtroCanalForms === 'TODOS'
        ? dadosBaseForms
        : dadosBaseForms.filter((d) => d._canal === filtroCanalForms);
    const setDivs = new Set(
      base.map((d) => d._divisao).filter((v) => v && v !== 'NÃO INFORMADA')
    );
    return Array.from(setDivs).sort();
  }, [dadosBaseForms, filtroCanalForms]);

  const opcoesTipoCertForms = useMemo(() => {
    const setTipos = new Set(
      dadosBaseForms
        .map((d) => d._tipoCert)
        .filter((v) => v && v !== 'NÃO INFORMADO')
    );
    return Array.from(setTipos).sort();
  }, [dadosBaseForms]);

  const formsFiltrados = useMemo(() => {
    return dadosBaseForms.filter((item) => {
      if (
        filtroAnoForms !== 'TODOS' &&
        item._anoCert !== Number(filtroAnoForms)
      )
        return false;
      if (
        filtroMesForms !== 'TODOS' &&
        item._mesCert !== Number(filtroMesForms)
      )
        return false;
      if (filtroCanalForms !== 'TODOS' && item._canal !== filtroCanalForms)
        return false;
      if (
        filtroDivisaoForms !== 'TODAS' &&
        item._divisao !== filtroDivisaoForms
      )
        return false;
      if (
        filtroTipoCertForms !== 'TODOS' &&
        item._tipoCert !== filtroTipoCertForms
      )
        return false;

      if (filtroAuditoriaForms === 'COM_DESVIO' && item._qtdDesvios === 0)
        return false;
      if (filtroAuditoriaForms === 'SEM_DESVIO' && item._qtdDesvios > 0)
        return false;
      if (filtroAuditoriaForms === 'CNH_ALERTA' && !item._cnhAlerta)
        return false;

      if (
        questaoDesvioSelecionada &&
        !item._listaDesvios.includes(questaoDesvioSelecionada)
      ) {
        return false;
      }

      if (termoBuscaForms) {
        const b = termoBuscaForms.toUpperCase();
        const bateCondutor = item._condutor.toUpperCase().includes(b);
        const bateGpidCond = item._gpidCondutor.toUpperCase().includes(b);
        const bateGpidAval = item._gpidAvaliador.toUpperCase().includes(b);
        const bateDiv = item._divisao.toUpperCase().includes(b);
        const bateComent = item._comentarios.toUpperCase().includes(b);
        return (
          bateCondutor || bateGpidCond || bateGpidAval || bateDiv || bateComent
        );
      }

      return true;
    });
  }, [
    dadosBaseForms,
    filtroAnoForms,
    filtroMesForms,
    filtroCanalForms,
    filtroDivisaoForms,
    filtroTipoCertForms,
    filtroAuditoriaForms,
    questaoDesvioSelecionada,
    termoBuscaForms,
  ]);

  const calcularGraficoCampo = (extrator, ignorarNaoInformado = false) => {
    if (formsFiltrados.length === 0) return [];

    const contagem = {};
    let totalValido = 0;

    formsFiltrados.forEach((row) => {
      const valor = String(extrator(row) || 'NÃO RESPONDIDO')
        .replace(/\u00a0/g, ' ')
        .trim()
        .toUpperCase();

      if (
        ignorarNaoInformado &&
        (valor === 'NÃO RESPONDIDO' ||
          valor === 'N/A' ||
          valor === 'NÃO INFORMADA')
      ) {
        return;
      }
      contagem[valor] = (contagem[valor] || 0) + 1;
      totalValido++;
    });

    const divisor = totalValido || 1;

    return Object.entries(contagem)
      .map(([opcao, qtd]) => ({
        opcao,
        qtd,
        porcentagem: ((qtd / divisor) * 100).toFixed(1),
      }))
      .sort((a, b) => b.qtd - a.qtd);
  };

  const rankingDesviosQuestoes = useMemo(() => {
    if (colunasQuestoesForms.length === 0 || formsFiltrados.length === 0)
      return [];

    return colunasQuestoesForms.map((col, idx) => {
      let eficientes = 0;
      let desvios = 0;
      let naoSeAplica = 0;

      formsFiltrados.forEach((r) => {
        const resp = String(r[col] || '').toUpperCase();
        if (
          resp.includes('NÃO EFICI') ||
          resp.includes('NAO EFICI') ||
          resp.includes('NÃO CONFORME') ||
          resp.includes('NAO CONFORME')
        ) {
          desvios++;
        } else if (resp.includes('APLICA')) {
          naoSeAplica++;
        } else if (resp.includes('EFICI') || resp.includes('CONFORME')) {
          eficientes++;
        }
      });

      const avaliadosValidos = eficientes + desvios || 1;
      const taxaConformidade = ((eficientes / avaliadosValidos) * 100).toFixed(
        1
      );

      return {
        id: idx + 1,
        coluna: col,
        eficientes,
        desvios,
        naoSeAplica,
        taxaConformidade,
      };
    });
  }, [colunasQuestoesForms, formsFiltrados]);

  const topQuestoesComFalha = useMemo(() => {
    return [...rankingDesviosQuestoes]
      .filter((q) => q.desvios > 0)
      .sort((a, b) => b.desvios - a.desvios);
  }, [rankingDesviosQuestoes]);

  const totalFormsExibidos = formsFiltrados.length;
  const totalSemDesvio = formsFiltrados.filter((d) => d._conformeGeral).length;
  const totalComDesvio = formsFiltrados.filter((d) => !d._conformeGeral).length;
  const somaApontamentosDesvio = formsFiltrados.reduce(
    (acc, cur) => acc + cur._qtdDesvios,
    0
  );
  const totalAlertasCNH = formsFiltrados.filter((d) => d._cnhAlerta).length;
  const taxaAprovacaoDireta =
    totalFormsExibidos > 0
      ? ((totalSemDesvio / totalFormsExibidos) * 100).toFixed(1)
      : '0.0';

  // --- LÓGICA DO MÓDULO 1: SINISTROS ---
  const limparFiltrosSinistros = () => {
    setAnoFocal(2026);
    setFiltroSetor('TODOS');
    setFiltroFrota('TODAS');
    setFiltroCategoriaVeiculo('PESADO');
    setFiltroCulpabilidade('TODAS');
    setTermoBuscaSinistros('');
    setMesFiltroKpi(new Date().getMonth() + 1);
  };

  function padronizarCulpabilidade(texto) {
    const t = normalizarTexto(texto);
    if (t.includes('PEPSICO')) return 'MOTORISTA PEPSICO';
    if (t.includes('TERC. IDENT') || t.includes('TERCEIRO IDENTIFICADO'))
      return 'TERCEIRO IDENTIFICADO';
    if (t.includes('NAO IDENT')) return 'TERCEIRO NÃO IDENTIFICADO';
    return 'OUTROS';
  }

  function classificarTipoVeiculo(nomeVeiculo) {
    const v = normalizarTexto(nomeVeiculo);

    if (
      v.includes('REBOQUE') ||
      v.includes('SR FACCHINI') ||
      v.includes('SR RANDON') ||
      v.includes('BITREM') ||
      v.includes('VANDERLEIA') ||
      v.includes('BAU 03 EIXOS') ||
      v.includes('SEMI REBOQUE') ||
      v.includes('REB ')
    ) {
      return 'REBOQUE';
    }

    const palavrasLeves = [
      'ONIX',
      'KA',
      'CRETA',
      'TRACKER',
      'HB20',
      'YARIS',
      'GOL',
      'COROLLA',
      'CIVIC',
      'T CROSS',
      'HRV',
      'HR-V',
      'TIGGO',
      'RENEGADE',
      'COMPASS',
      'BRONCO',
      'NIVUS',
      'TORO',
      '2008',
      'FIORINO',
      'PULSE',
      'FASTBACK',
      'TAOS',
      'MOBI',
      'ARGO',
      'KWID',
      'OROCH',
      'VOYAGE',
      'POLO',
      'STRADA',
    ];

    if (v.includes('DAILY') || v.includes('CARGO') || v.includes('ACCELO')) {
      return 'PESADO';
    }

    const ehLeve = palavrasLeves.some((termo) => v.includes(termo));
    return ehLeve ? 'LEVE' : 'PESADO';
  }

  function deduzirSetorOperacional(
    tipoFrotaRaw,
    divisaoRaw,
    filialRaw,
    categoriaVeiculo,
    diretoriaPosicao = ''
  ) {
    const dirPos = normalizarTexto(diretoriaPosicao);
    const frota = normalizarTexto(tipoFrotaRaw);
    const div = normalizarTexto(divisaoRaw);
    const fil = normalizarTexto(filialRaw);

    if (dirPos) {
      if (dirPos === 'DTS') return 'DTS';
      if (dirPos.includes('HYPER TRANSFORMATION') || dirPos.includes('TRANSF'))
        return 'DTS LEVE';
      if (dirPos.includes('TRANSPORTES')) return 'TRANSPORTES';
      if (
        dirPos.includes('SULESTE') ||
        dirPos.includes('SP + CO') ||
        dirPos.includes('NENO') ||
        dirPos.includes('SUL + CO') ||
        dirPos.includes('LESTE') ||
        dirPos.includes('NATIONAL ONE') ||
        dirPos.includes('SP ESTADO')
      ) {
        return 'OT';
      }
      if (
        dirPos.includes('SACCAF') ||
        dirPos.includes('PEPSI') ||
        dirPos.includes('RH') ||
        dirPos.includes('MANUFATURA') ||
        dirPos.includes('FINANCE') ||
        dirPos.includes('JURIDICO') ||
        dirPos.includes('PROCUREMENT')
      ) {
        return 'EXECUTIVA';
      }
    }

    if (
      frota.includes('PRIMARIA') ||
      div === 'CARRETAS' ||
      div === 'ITU' ||
      div === '7L' ||
      div === 'CIC' ||
      fil.includes('FROTA PRIM') ||
      fil.includes('GERENCIA FRETE') ||
      fil.includes('FRETE PRIM') ||
      fil.includes('FRETE CD')
    ) {
      return 'TRANSPORTES';
    }

    if (
      frota.includes('EXECUTIV') ||
      div.includes('EXECUTIV') ||
      div.includes('ESCRITORIO') ||
      fil.includes('ESCRITORIO')
    ) {
      return 'EXECUTIVA';
    }

    if (
      frota.includes('LOCADOS') ||
      frota.includes('ALD') ||
      frota.includes('LEASING') ||
      categoriaVeiculo === 'LEVE'
    ) {
      return 'OT / DTS LEVE';
    }

    if (frota.includes('SECUNDARIA') || div.includes('DTS')) {
      return 'DTS';
    }

    return 'OUTROS';
  }

  const processarDadosLivroSinistros = (livro) => {
    let abaAlvo = null;
    let registrosBrutos = [];

    for (const nomeAba of livro.SheetNames) {
      if (nomeAba.toLowerCase().includes('base')) {
        abaAlvo = nomeAba;
        break;
      }
    }

    if (!abaAlvo) {
      for (const nomeAba of livro.SheetNames) {
        const folha = livro.Sheets[nomeAba];
        const jsonTemp = XLSX.utils.sheet_to_json(folha, { defval: '' });
        if (jsonTemp.length > 0) {
          const chaves = Object.keys(jsonTemp[0]).map((k) => k.toUpperCase());
          const possuiColunasSinistro = chaves.some(
            (k) =>
              k.includes('PLACA') ||
              k.includes('SINISTRO') ||
              k.includes('GPID')
          );
          if (possuiColunasSinistro) {
            abaAlvo = nomeAba;
            break;
          }
        }
      }
    }

    if (!abaAlvo) {
      abaAlvo = livro.SheetNames[2] || livro.SheetNames[0];
    }

    registrosBrutos = XLSX.utils.sheet_to_json(livro.Sheets[abaAlvo], {
      defval: '',
    });

    const listaSinistros = [];
    registrosBrutos.forEach((linha, index) => {
      const l = {};
      Object.keys(linha).forEach((k) => {
        l[k.trim().toUpperCase()] = linha[k];
      });

      const chavePlaca =
        Object.keys(l).find((k) => k.includes('PLACA')) || 'PLACA';
      const placa = String(l[chavePlaca] || '')
        .toUpperCase()
        .trim();

      const chaveData =
        Object.keys(l).find(
          (k) =>
            k.includes('DATA') ||
            k.includes('OCORRÊNCIA') ||
            k.includes('OCORRENCIA')
        ) || 'DATA';
      const dataBruta = l[chaveData] || '';

      let { ano, mes, formatada } = extrairDataCompletaExcel(dataBruta);

      if (!ano) {
        const strD = String(dataBruta);
        if (strD.includes('2025') || strD.includes('/25')) ano = 2025;
        else if (strD.includes('2026') || strD.includes('/26')) ano = 2026;
        else ano = 2026;
      }

      if (!mes) mes = 1;

      if (placa && placa !== 'N/A' && placa !== 'NULL') {
        const chaveCulp =
          Object.keys(l).find((k) => k.includes('CULP')) || 'CULPABILIDADE';

        const chaveDivisao =
          Object.keys(l).find(
            (k) => k.includes('DIVISÃO') || k.includes('DIVISAO')
          ) || 'DIVISÃO';

        const chaveSetor =
          Object.keys(l).find(
            (k) =>
              k.includes('DIRETORIA') ||
              k.includes('DIRET.') ||
              k.includes('SETOR')
          ) || '';

        const chaveVeiculo =
          Object.keys(l).find(
            (k) => k.includes('VEÍCUL') || k.includes('VEICULO')
          ) || 'VEÍCULO';
        const nomeVeiculo = String(l[chaveVeiculo] || '').trim();

        const chaveFrota =
          Object.keys(l).find((k) => k.includes('FROTA')) || 'TIPO DE FROTA';
        const chaveAcidente =
          Object.keys(l).find(
            (k) =>
              k.includes('ACIDENTE') ||
              k.includes('CAUSA') ||
              k.includes('MOTIVO')
          ) || 'TIPO ACIDENTE';

        const filialPadronizada = padronizarFilial(
          l['FILIAL'] || l['CDV'] || 'SEM INFORMAÇÃO'
        );
        const divisaoBruta = String(l[chaveDivisao] || '').trim();
        const tipoFrotaBruto = String(l[chaveFrota] || 'OUTROS')
          .toUpperCase()
          .trim();
        const catVeiculo = classificarTipoVeiculo(nomeVeiculo);

        const infoPos = mapaPosicaoNacional[placa];
        const diretoriaExplicit = chaveSetor ? String(l[chaveSetor] || '') : '';

        const setorDeduzido = deduzirSetorOperacional(
          infoPos?.frota || tipoFrotaBruto,
          divisaoBruta,
          filialPadronizada,
          catVeiculo,
          infoPos?.diretoria || diretoriaExplicit
        );

        listaSinistros.push({
          id: index,
          placa: placa,
          numSinistro: String(
            l['Nº DO SINISITRO'] ||
              l['Nº DO SINISTRO'] ||
              l['NUMERO SINISTRO'] ||
              l['SINISTRO'] ||
              index + 1
          ).trim(),
          filial: filialPadronizada,
          divisao: divisaoBruta.toUpperCase() || 'N/I',
          data: formatada || String(dataBruta) || 'N/I',
          ano: ano,
          mes: mes,
          gpid: String(l['GPID'] || l['RE'] || '').trim(),
          motorista: String(
            l['MOTORISTA'] || l['CONDUTOR'] || l['NOME'] || ''
          ).trim(),
          culpabilidade: padronizarCulpabilidade(l[chaveCulp] || ''),
          setor: setorDeduzido,
          veiculo: nomeVeiculo,
          categoriaVeiculo: catVeiculo,
          tipoFrota: tipoFrotaBruto,
          tipoAcidente: String(l[chaveAcidente] || 'OUTROS')
            .toUpperCase()
            .trim(),
        });
      }
    });

    if (listaSinistros.length > 0) {
      setDadosBaseSinistros(listaSinistros);
      setCarregadoSinistros(true);
    }
  };

  const lidarComUploadSinistros = (e) => {
    const arquivo = e.target.files[0];
    if (!arquivo) return;
    const leitor = new FileReader();
    leitor.onload = (evento) => {
      try {
        const dadosBinarios = evento.target.result;
        const livro = XLSX.read(dadosBinarios, { type: 'binary' });
        processarDadosLivroSinistros(livro);
      } catch (err) {
        alert('Erro ao processar planilha de sinistros: ' + err.message);
      }
    };
    leitor.readAsBinaryString(arquivo);
  };

  const itemPassaFiltrosBase = (item, anoAlvo) => {
    if (item.ano !== anoAlvo) return false;

    const frotaNorm = normalizarTexto(item.tipoFrota);
    let passaFrota = true;
    if (filtroFrota === 'PRIMARIA') {
      passaFrota = frotaNorm.includes('PRIMARIA');
    } else if (filtroFrota === 'SECUNDARIA_VENDAS') {
      passaFrota = frotaNorm.includes('VENDAS') || frotaNorm === 'SECUNDARIA';
    } else if (filtroFrota === 'SECUNDARIA_LOCADOS') {
      passaFrota =
        frotaNorm.includes('LOCADOS') ||
        frotaNorm.includes('ALD') ||
        frotaNorm.includes('LEASING');
    } else if (filtroFrota === 'SECUNDARIA_TODAS') {
      passaFrota =
        frotaNorm.includes('SECUNDARIA') || frotaNorm.includes('ALD');
    } else if (filtroFrota === 'EXECUTIVA') {
      passaFrota = frotaNorm.includes('EXECUTIV');
    }

    let passaSetor =
      filtroSetor === 'TODOS' ||
      normalizarTexto(item.setor).includes(normalizarTexto(filtroSetor));

    let passaCategoria = true;
    if (filtroCategoriaVeiculo === 'PESADO') {
      passaCategoria = item.categoriaVeiculo === 'PESADO';
    } else if (filtroCategoriaVeiculo === 'PESADO_E_REBOQUE') {
      passaCategoria =
        item.categoriaVeiculo === 'PESADO' ||
        item.categoriaVeiculo === 'REBOQUE';
    } else if (filtroCategoriaVeiculo === 'LEVE') {
      passaCategoria = item.categoriaVeiculo === 'LEVE';
    } else if (filtroCategoriaVeiculo === 'REBOQUE') {
      passaCategoria = item.categoriaVeiculo === 'REBOQUE';
    }

    return passaFrota && passaSetor && passaCategoria;
  };

  const baseSemCulpabilidade = useMemo(() => {
    return dadosBaseSinistros.filter((item) => {
      if (!itemPassaFiltrosBase(item, anoFocal)) return false;

      if (termoBuscaSinistros) {
        const b = normalizarTexto(termoBuscaSinistros);
        return (
          normalizarTexto(item.placa).includes(b) ||
          normalizarTexto(item.filial).includes(b) ||
          normalizarTexto(item.veiculo).includes(b) ||
          normalizarTexto(item.motorista).includes(b) ||
          normalizarTexto(item.setor).includes(b)
        );
      }
      return true;
    });
  }, [
    dadosBaseSinistros,
    anoFocal,
    filtroFrota,
    filtroSetor,
    filtroCategoriaVeiculo,
    termoBuscaSinistros,
  ]);

  const sinistrosFiltrados = useMemo(() => {
    return baseSemCulpabilidade.filter((item) => {
      return (
        filtroCulpabilidade === 'TODAS' ||
        item.culpabilidade === filtroCulpabilidade
      );
    });
  }, [baseSemCulpabilidade, filtroCulpabilidade]);

  const contagemCulpabilidade = useMemo(() => {
    const contagem = {
      'MOTORISTA PEPSICO': 0,
      'TERCEIRO IDENTIFICADO': 0,
      'TERCEIRO NÃO IDENTIFICADO': 0,
      OUTROS: 0,
    };
    baseSemCulpabilidade.forEach((d) => {
      if (contagem[d.culpabilidade] !== undefined) contagem[d.culpabilidade]++;
      else contagem['OUTROS']++;
    });
    return contagem;
  }, [baseSemCulpabilidade]);

  const totalCulpGeral = baseSemCulpabilidade.length || 1;

  const sinistrosMesFocal = sinistrosFiltrados.filter(
    (d) => d.ano === anoFocal && d.mes === Number(mesFiltroKpi)
  );
  const sinistrosMesComp = dadosBaseSinistros.filter((d) => {
    if (!itemPassaFiltrosBase(d, anoComparacao)) return false;
    let passaCulp =
      filtroCulpabilidade === 'TODAS' ||
      d.culpabilidade === filtroCulpabilidade;
    return d.mes === Number(mesFiltroKpi) && passaCulp;
  });

  const totalMesFocal = sinistrosMesFocal.length;
  const totalMesComp = sinistrosMesComp.length;
  const diffAbsoluta = totalMesFocal - totalMesComp;

  const sinistrosAnoFocal = sinistrosFiltrados;
  const sinistrosAnoComp = dadosBaseSinistros.filter((d) => {
    if (!itemPassaFiltrosBase(d, anoComparacao)) return false;
    let passaCulp =
      filtroCulpabilidade === 'TODAS' ||
      d.culpabilidade === filtroCulpabilidade;
    return passaCulp;
  });

  const mesesNomes = [
    'Jan',
    'Fev',
    'Mar',
    'Abr',
    'Mai',
    'Jun',
    'Jul',
    'Ago',
    'Set',
    'Out',
    'Nov',
    'Dez',
  ];
  const dadosEvolucaoMensal = mesesNomes.map((nomeMes, index) => {
    const numMes = index + 1;
    const qtdFocal = sinistrosAnoFocal.filter((d) => d.mes === numMes).length;
    const qtdComp = sinistrosAnoComp.filter((d) => d.mes === numMes).length;
    return { mes: nomeMes, numMes, qtdFocal, qtdComp };
  });

  const piorMesObj = [...dadosEvolucaoMensal].sort(
    (a, b) => b.qtdFocal - a.qtdFocal
  )[0];

  const contagemCausas = {};
  sinistrosFiltrados.forEach((d) => {
    contagemCausas[d.tipoAcidente] = (contagemCausas[d.tipoAcidente] || 0) + 1;
  });
  const rankingCausas = Object.entries(contagemCausas)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);

  const rankingFiliais = useMemo(() => {
    const contagem = {};
    sinistrosFiltrados.forEach((d) => {
      const f = d.filial;
      if (apenasCDVsNoRanking) {
        if (
          f.includes('FROTA PRIM') ||
          f.includes('FRETE PRIM') ||
          f.includes('GERENCIA FRETE') ||
          f.includes('FRETE CD') ||
          f === 'SEM INFORMAÇÃO'
        ) {
          return;
        }
      }
      contagem[f] = (contagem[f] || 0) + 1;
    });
    return Object.entries(contagem)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5);
  }, [sinistrosFiltrados, apenasCDVsNoRanking]);

  const maxSinistrosFilial = rankingFiliais[0] ? rankingFiliais[0][1] : 1;

  const selecionarCulpabilidadeCard = (tipo) => {
    if (filtroCulpabilidade === tipo) {
      setFiltroCulpabilidade('TODAS');
    } else {
      setFiltroCulpabilidade(tipo);
    }
  };

  const paletaFormsHex = [
    '#5B73E8',
    '#E8117F',
    '#00A896',
    '#7D52B5',
    '#107C41',
    '#D83B01',
    '#A4262C',
    '#0078D4',
  ];

  const graficosPrincipaisForms = useMemo(() => {
    return [
      {
        titulo: '1. TIPO DE CERTIFICAÇÃO',
        tipo: 'donut',
        dados: calcularGraficoCampo((r) => r._tipoCert),
      },
      {
        titulo: '2. CANAL OPERACIONAL (OT vs DTS)',
        tipo: 'donut',
        dados: calcularGraficoCampo((r) => r._canal),
      },
      {
        titulo: '3. DISTRIBUIÇÃO POR DIVISÃO / REGIONAL',
        tipo: 'barras',
        dados: calcularGraficoCampo((r) => r._divisao, true),
      },
      {
        titulo: '4. CONFORMIDADE GERAL DOS AVALIADOS',
        tipo: 'donut',
        dados: calcularGraficoCampo((r) =>
          r._conformeGeral ? '100% CONFORME / EFICIENTE' : 'COM DESVIO / FALHA'
        ),
      },
    ];
  }, [formsFiltrados]);

  return (
    <div className="min-h-screen bg-[#0B1329] text-gray-100 font-sans">
      {/* 🏠 TELA 1: MENU PRINCIPAL DE MÓDULOS (AGORA COM 3 MÓDULOS) */}
      {telaAtiva === 'menu' && (
        <div className="min-h-screen flex flex-col justify-between p-8 relative overflow-hidden bg-gradient-to-b from-[#0B1329] via-[#101B3B] to-[#0B1329]">
          <div className="absolute top-0 left-1/2 -translate-x-1/2 w-[800px] h-[350px] bg-orange-500/10 blur-[120px] rounded-full pointer-events-none"></div>
          <div className="absolute bottom-0 right-0 w-[500px] h-[500px] bg-blue-600/10 blur-[150px] rounded-full pointer-events-none"></div>

          <header className="flex justify-between items-center z-10 max-w-6xl mx-auto w-full border-b border-gray-800/80 pb-6">
            <div className="flex items-center gap-3">
              <div className="bg-gradient-to-tr from-orange-500 to-amber-400 p-2.5 rounded-2xl shadow-lg shadow-orange-500/20">
                <ShieldCheck className="text-white" size={26} />
              </div>
              <div>
                <h1 className="text-xl font-black text-white tracking-wide">
                  PORTAL CONTROL TOWER
                </h1>
                <p className="text-xs text-gray-400">
                  Operação Logística PepsiCo / Opentech
                </p>
              </div>
            </div>

            <div className="bg-[#1C2541] border border-gray-800 px-4 py-1.5 rounded-full flex items-center gap-2 text-xs font-bold text-gray-300">
              <Sparkles size={14} className="text-amber-400" />
              <span>Sistema Unificado 2026</span>
            </div>
          </header>

          <main className="z-10 max-w-6xl mx-auto w-full my-auto py-12">
            <div className="text-center mb-12">
              <h2 className="text-4xl font-black text-white tracking-tight">
                Selecione o Módulo de Análise
              </h2>
              <p className="text-sm text-gray-400 mt-2 max-w-lg mx-auto">
                Escolha o painel operacional para gerenciar os dados
                estratégicos de sinistros, certificações e controle de frota
                leve (DTS Leve & OT).
              </p>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
              {/* CARD 1: SINISTROS */}
              <div
                onClick={() => setTelaAtiva('sinistros')}
                className="group relative bg-[#1C2541] border border-gray-800 hover:border-orange-500/80 p-7 rounded-3xl cursor-pointer transition-all duration-300 hover:-translate-y-2 hover:shadow-2xl hover:shadow-orange-500/10 flex flex-col justify-between overflow-hidden"
              >
                <div className="absolute top-0 right-0 w-32 h-32 bg-orange-500/5 rounded-bl-full group-hover:bg-orange-500/10 transition-all"></div>

                <div>
                  <div className="w-14 h-14 bg-gradient-to-tr from-orange-500 to-amber-500 rounded-2xl flex items-center justify-center text-white shadow-lg mb-6 group-hover:scale-110 transition-all">
                    <LayoutDashboard size={28} />
                  </div>

                  <span className="text-[10px] font-black tracking-widest text-orange-400 uppercase bg-orange-500/10 px-3 py-1 rounded-full border border-orange-500/20">
                    Módulo 1 • Riscos
                  </span>

                  <h3 className="text-xl font-black text-white mt-3 group-hover:text-orange-400 transition-colors">
                    Control Tower (Sinistros)
                  </h3>

                  <p className="text-xs text-gray-400 mt-2 leading-relaxed">
                    Painel executivo de acidentes e sinistros. Compare 2025 vs
                    2026, separe Caminhões de Reboques/Carretas e filtre por
                    DTS, OT e Primária.
                  </p>
                </div>

                <div className="mt-8 flex items-center justify-between border-t border-gray-800/80 pt-4 text-xs font-bold text-gray-300 group-hover:text-white">
                  <span>
                    {carregadoSinistros
                      ? `✅ ${dadosBaseSinistros.length} registros`
                      : 'Sincronização em tempo real'}
                  </span>
                  <div className="w-8 h-8 rounded-full bg-[#0B1329] flex items-center justify-center group-hover:bg-orange-500 group-hover:text-white transition-all">
                    <ArrowRight size={16} />
                  </div>
                </div>
              </div>

              {/* CARD 2: FORMS CERTIFICAÇÃO */}
              <div
                onClick={() => setTelaAtiva('forms')}
                className="group relative bg-[#1C2541] border border-gray-800 hover:border-blue-500/80 p-7 rounded-3xl cursor-pointer transition-all duration-300 hover:-translate-y-2 hover:shadow-2xl hover:shadow-blue-500/10 flex flex-col justify-between overflow-hidden"
              >
                <div className="absolute top-0 right-0 w-32 h-32 bg-blue-500/5 rounded-bl-full group-hover:bg-blue-500/10 transition-all"></div>

                <div>
                  <div className="w-14 h-14 bg-gradient-to-tr from-blue-600 to-indigo-500 rounded-2xl flex items-center justify-center text-white shadow-lg mb-6 group-hover:scale-110 transition-all">
                    <ClipboardCheck size={28} />
                  </div>

                  <span className="text-[10px] font-black tracking-widest text-blue-400 uppercase bg-blue-500/10 px-3 py-1 rounded-full border border-blue-500/20">
                    Módulo 2 • Certificação SELL
                  </span>

                  <h3 className="text-xl font-black text-white mt-3 group-hover:text-blue-400 transition-colors">
                    Certificação Direção Segura
                  </h3>

                  <p className="text-xs text-gray-400 mt-2 leading-relaxed">
                    Auditoria inteligente das 24 regras de segurança e Método
                    Smith. Filtre por Canal (OT/DTS), Divisão, Alertas de CNH e
                    Ranking de Desvios.
                  </p>
                </div>

                <div className="mt-8 flex items-center justify-between border-t border-gray-800/80 pt-4 text-xs font-bold text-gray-300 group-hover:text-white">
                  <span>
                    {carregadoForms
                      ? `⚡ ${dadosBaseForms.length} avaliações`
                      : 'Sincronização em tempo real'}
                  </span>
                  <div className="w-8 h-8 rounded-full bg-[#0B1329] flex items-center justify-center group-hover:bg-blue-600 group-hover:text-white transition-all">
                    <ArrowRight size={16} />
                  </div>
                </div>
              </div>

              {/* CARD 3: BASE DE LEVES ATIVOS (DTS LEVE vs OT) */}
              <div
                onClick={() => setTelaAtiva('frota_leves')}
                className="group relative bg-[#1C2541] border border-gray-800 hover:border-emerald-500/80 p-7 rounded-3xl cursor-pointer transition-all duration-300 hover:-translate-y-2 hover:shadow-2xl hover:shadow-emerald-500/10 flex flex-col justify-between overflow-hidden"
              >
                <div className="absolute top-0 right-0 w-32 h-32 bg-emerald-500/5 rounded-bl-full group-hover:bg-emerald-500/10 transition-all"></div>

                <div>
                  <div className="w-14 h-14 bg-gradient-to-tr from-emerald-500 to-teal-400 rounded-2xl flex items-center justify-center text-white shadow-lg mb-6 group-hover:scale-110 transition-all">
                    <Car size={28} />
                  </div>

                  <span className="text-[10px] font-black tracking-widest text-emerald-400 uppercase bg-emerald-500/10 px-3 py-1 rounded-full border border-emerald-500/20">
                    Módulo 3 • Posição Nacional
                  </span>

                  <h3 className="text-xl font-black text-white mt-3 group-hover:text-emerald-400 transition-colors">
                    Frota Leve Ativa (DTS Leve & OT)
                  </h3>

                  <p className="text-xs text-gray-400 mt-2 leading-relaxed">
                    Higienização e separação exata dos veículos leves ativos
                    entre <strong>DTS Leve</strong> e <strong>OT</strong> por
                    Regional, Gerente e Centro de Custo + Exportação Excel.
                  </p>
                </div>

                <div className="mt-8 flex items-center justify-between border-t border-gray-800/80 pt-4 text-xs font-bold text-gray-300 group-hover:text-white">
                  <span>
                    {carregadoFrotaAtiva
                      ? `🚗 ${dadosFrotaAtiva.length} leves mapeados`
                      : 'Controle por Diretoria & CCusto'}
                  </span>
                  <div className="w-8 h-8 rounded-full bg-[#0B1329] flex items-center justify-center group-hover:bg-emerald-500 group-hover:text-white transition-all">
                    <ArrowRight size={16} />
                  </div>
                </div>
              </div>
            </div>
          </main>

          <footer className="z-10 text-center text-xs text-gray-500 border-t border-gray-800/80 pt-6 max-w-6xl mx-auto w-full flex justify-between items-center">
            <p>© 2026 PepsiCo / Opentech Operations</p>
            <p className="font-bold text-gray-400">Desenvolvido por Fidel</p>
          </footer>
        </div>
      )}

      {/* 🛡️ TELA 2: CONTROL TOWER DE SINISTROS */}
      {telaAtiva === 'sinistros' && (
        <div className="p-8">
          <div className="mb-6 flex items-center justify-between bg-[#1C2541] border border-gray-800 p-3 rounded-2xl">
            <button
              onClick={() => setTelaAtiva('menu')}
              className="bg-[#0B1329] hover:bg-gray-800 text-xs font-black text-orange-400 px-4 py-2 rounded-xl flex items-center gap-2 border border-gray-700 transition-all cursor-pointer"
            >
              <ArrowLeft size={16} /> Voltar ao Menu Principal
            </button>
            <div className="flex items-center gap-3">
              {qtdPlacasPosicao > 0 && (
                <span className="text-[11px] bg-cyan-500/20 text-cyan-400 border border-cyan-500/30 font-black px-3 py-1 rounded-full">
                  🔗 Posição Nacional Ativa ({qtdPlacasPosicao} placas)
                </span>
              )}
              <span className="text-xs font-black text-gray-400 uppercase tracking-widest hidden md:inline">
                Módulo: Control Tower Sinistros
              </span>
            </div>
          </div>

          <header className="sticky top-0 z-50 backdrop-blur-md bg-[#0B1329]/95 py-4 mb-8 border-b border-gray-800/80 shadow-2xl flex flex-col gap-4 transition-all">
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
              <div>
                <h1 className="text-2xl font-black text-white tracking-tight flex items-center gap-2">
                  <ShieldCheck className="text-orange-500" />
                  Control Tower | Operação Sinistros & Riscos
                </h1>
                <p className="text-xs text-gray-400 mt-0.5">
                  Análise Comparativa Exclusiva ({anoFocal} vs {anoComparacao}) •
                  PepsiCo / Opentech
                </p>
              </div>

              <div className="flex flex-wrap items-center gap-2.5">
                <button
                  onClick={carregarSinistrosNuvem}
                  className="bg-[#1C2541] hover:bg-gray-800 text-orange-400 border border-gray-700 text-xs font-black px-3.5 py-2 rounded-xl flex items-center gap-2 transition-all cursor-pointer shadow-sm"
                >
                  <RefreshCw
                    size={14}
                    className={carregandoNuvemSinistros ? 'animate-spin' : ''}
                  />
                  Sincronizar OneDrive
                </button>

                <div className="bg-[#1C2541] border border-orange-500/50 p-1 rounded-xl flex text-xs font-bold shadow-md">
                  <button
                    onClick={() => setAnoFocal(2026)}
                    className={`px-3 py-1.5 rounded-lg transition-all cursor-pointer ${
                      anoFocal === 2026
                        ? 'bg-orange-500 text-white font-black'
                        : 'text-gray-400 hover:text-white'
                    }`}
                  >
                    Ano Focal: 2026
                  </button>
                  <button
                    onClick={() => setAnoFocal(2025)}
                    className={`px-3 py-1.5 rounded-lg transition-all cursor-pointer ${
                      anoFocal === 2025
                        ? 'bg-blue-600 text-white font-black'
                        : 'text-gray-400 hover:text-white'
                    }`}
                  >
                    Ano Focal: 2025
                  </button>
                </div>

                <button
                  onClick={limparFiltrosSinistros}
                  className="bg-[#1C2541] hover:bg-gray-800 border border-gray-700 text-gray-300 hover:text-white text-xs font-bold px-3 py-2 rounded-xl flex items-center gap-1.5 transition-all cursor-pointer shadow-sm"
                >
                  <RotateCcw size={14} className="text-orange-400" /> Limpar
                  Filtros
                </button>

                <label
                  className="bg-cyan-600 hover:bg-cyan-700 text-white text-xs font-black px-3.5 py-2 rounded-xl cursor-pointer flex items-center gap-1.5 shadow-lg transition-all"
                  title="Carregue a planilha Posição Nacional Frota para cruzar Diretoria (DTS/OT) por placa"
                >
                  <Database size={14} />{' '}
                  {qtdPlacasPosicao > 0
                    ? 'Posição Nacional OK'
                    : 'Cruzar Posição Nacional'}
                  <input
                    type="file"
                    accept=".xlsx, .xls"
                    onChange={lidarComUploadPosicaoNacional}
                    className="hidden"
                  />
                </label>

                <label className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-black px-4 py-2 rounded-xl cursor-pointer flex items-center gap-1.5 shadow-lg transition-all">
                  <Upload size={14} />{' '}
                  {carregadoSinistros ? 'Trocar Sinistros' : 'Carregar Manual'}
                  <input
                    type="file"
                    accept=".xlsx, .xls"
                    onChange={lidarComUploadSinistros}
                    className="hidden"
                  />
                </label>
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-2.5 pt-2 border-t border-gray-800/80">
              <select
                value={filtroCategoriaVeiculo}
                onChange={(e) => setFiltroCategoriaVeiculo(e.target.value)}
                className="bg-[#1C2541] border border-emerald-500/50 text-xs font-black px-3 py-2 rounded-xl text-emerald-400 outline-none cursor-pointer hover:border-emerald-400 transition-all"
              >
                <option value="PESADO">
                  🚛 Apenas Caminhões (Sem Reboques/Carretas)
                </option>
                <option value="TODOS">
                  🚗 Todos os Veículos (Leves + Caminhões + Reboques)
                </option>
                <option value="PESADO_E_REBOQUE">
                  🚚 Caminhões + Reboques/Carretas
                </option>
                <option value="LEVE">
                  🚙 Apenas Carros Leves (OT / DTS Leve / Executiva)
                </option>
                <option value="REBOQUE">
                  🔗 Apenas Semi-Reboques / Carretas
                </option>
              </select>

              <select
                value={filtroSetor}
                onChange={(e) => setFiltroSetor(e.target.value)}
                className="bg-[#1C2541] border border-cyan-500/40 text-xs font-black px-3 py-2 rounded-xl text-cyan-400 outline-none cursor-pointer hover:border-cyan-400 transition-all"
              >
                <option value="TODOS">
                  🏢 Todos os Setores (DTS / OT / Transportes)
                </option>
                <option value="DTS">Setor: DTS (Secundária / DTS Leve)</option>
                <option value="OT">
                  Setor: OT (Vendas Leves / Locados / ALD)
                </option>
                <option value="TRANSPORTES">
                  Setor: Transportes (Frota Primária)
                </option>
                <option value="EXECUTIVA">Setor: Frota Executiva</option>
              </select>

              <select
                value={filtroFrota}
                onChange={(e) => setFiltroFrota(e.target.value)}
                className="bg-[#1C2541] border border-gray-800 text-xs font-bold px-3 py-2 rounded-xl text-orange-400 outline-none cursor-pointer hover:border-gray-700 transition-all"
              >
                <option value="TODAS">
                  Todas as Frotas (Primária / Secundária)
                </option>
                <option value="SECUNDARIA_VENDAS">
                  Secundária - Vendas (DTS)
                </option>
                <option value="SECUNDARIA_LOCADOS">
                  Secundária - Locados / ALD Flex (Leves)
                </option>
                <option value="SECUNDARIA_TODAS">
                  Toda Frota Secundária (Vendas + Locados)
                </option>
                <option value="PRIMARIA">Frota Primária (Transportes)</option>
                <option value="EXECUTIVA">Frota Executiva</option>
              </select>

              <select
                value={filtroCulpabilidade}
                onChange={(e) => setFiltroCulpabilidade(e.target.value)}
                className="bg-[#1C2541] border border-gray-800 text-xs font-bold px-3 py-2 rounded-xl text-orange-400 outline-none cursor-pointer hover:border-gray-700 transition-all"
              >
                <option value="TODAS">Todas Culpabilidades</option>
                <option value="MOTORISTA PEPSICO">Motorista PepsiCo</option>
                <option value="TERCEIRO IDENTIFICADO">
                  Terceiro Identificado
                </option>
                <option value="TERCEIRO NÃO IDENTIFICADO">
                  Terceiro Não Identificado
                </option>
              </select>
            </div>
          </header>

          {!carregadoSinistros && !carregandoNuvemSinistros ? (
            <div className="bg-[#1C2541] border border-dashed border-gray-700 rounded-3xl p-12 text-center max-w-2xl mx-auto my-12">
              <FileSpreadsheet
                className="mx-auto text-orange-500 mb-4 animate-bounce"
                size={48}
              />
              <h2 className="text-xl font-black text-white">
                Sincronizador do OneDrive
              </h2>
              <p className="text-xs text-gray-400 mt-2 max-w-md mx-auto">
                Buscando a planilha `Sinistros.xlsx` da pasta PROJETO_CONTROL no
                seu OneDrive...
              </p>
              <div className="mt-6 flex justify-center gap-3">
                <button
                  onClick={carregarSinistrosNuvem}
                  className="bg-orange-500 hover:bg-orange-600 text-white font-black text-xs px-6 py-3 rounded-xl cursor-pointer shadow-lg transition-all flex items-center gap-2"
                >
                  <RefreshCw size={14} /> Tentar Sincronizar Novamente
                </button>
                <label className="bg-gray-800 hover:bg-gray-700 text-gray-200 font-black text-xs px-6 py-3 rounded-xl cursor-pointer shadow-lg transition-all">
                  Upload Manual (.XLSX)
                  <input
                    type="file"
                    accept=".xlsx, .xls"
                    onChange={lidarComUploadSinistros}
                    className="hidden"
                  />
                </label>
              </div>
            </div>
          ) : (
            <>
              <div className="bg-[#1C2541] border border-gray-800 p-3 rounded-xl mb-6 flex items-center gap-3 shadow-md">
                <Search size={16} className="text-gray-400" />
                <input
                  type="text"
                  value={termoBuscaSinistros}
                  onChange={(e) => setTermoBuscaSinistros(e.target.value)}
                  placeholder="Pesquisar por Placa, Veículo, Filial, Setor (DTS/OT) ou Motorista..."
                  className="bg-transparent w-full text-xs text-white outline-none font-bold"
                />
              </div>

              <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-8">
                <div className="bg-[#1C2541] border border-gray-800 p-5 rounded-2xl flex flex-col justify-between shadow-md">
                  <div className="flex justify-between items-center text-xs font-bold text-gray-400 mb-2">
                    <select
                      value={mesFiltroKpi}
                      onChange={(e) => setMesFiltroKpi(Number(e.target.value))}
                      className="bg-[#0B1329] border border-gray-700 text-orange-400 font-black text-xs px-2 py-1 rounded-lg outline-none cursor-pointer uppercase"
                    >
                      {[
                        { num: 1, nome: 'Janeiro' },
                        { num: 2, nome: 'Fevereiro' },
                        { num: 3, nome: 'Março' },
                        { num: 4, nome: 'Abril' },
                        { num: 5, nome: 'Maio' },
                        { num: 6, nome: 'Junho' },
                        { num: 7, nome: 'Julho' },
                        { num: 8, nome: 'Agosto' },
                        { num: 9, nome: 'Setembro' },
                        { num: 10, nome: 'Outubro' },
                        { num: 11, nome: 'Novembro' },
                        { num: 12, nome: 'Dezembro' },
                      ].map((m) => (
                        <option key={m.num} value={m.num}>
                          {m.nome}
                        </option>
                      ))}
                    </select>

                    <span
                      className={`flex items-center text-[10px] px-2 py-0.5 rounded font-black ${
                        diffAbsoluta <= 0
                          ? 'bg-emerald-500/20 text-emerald-400'
                          : 'bg-red-500/20 text-red-400'
                      }`}
                    >
                      {diffAbsoluta <= 0 ? (
                        <TrendingDown size={12} className="mr-1" />
                      ) : (
                        <TrendingUp size={12} className="mr-1" />
                      )}
                      {diffAbsoluta > 0 ? `+${diffAbsoluta}` : diffAbsoluta}{' '}
                      cas. vs {anoComparacao}
                    </span>
                  </div>

                  <div>
                    <h2 className="text-3xl font-black text-white mt-1">
                      {totalMesFocal}{' '}
                      <span className="text-xs font-normal text-gray-400">
                        Ocorrências ({anoFocal})
                      </span>
                    </h2>
                    <p className="text-[10px] text-gray-400 mt-2">
                      Mesmo mês em {anoComparacao}:{' '}
                      <strong className="text-gray-200">
                        {totalMesComp} ocorrências
                      </strong>
                    </p>
                  </div>
                </div>

                <div className="bg-[#1C2541] border border-gray-800 p-5 rounded-2xl shadow-md">
                  <span className="text-xs font-bold text-gray-400">
                    MÊS DE MAIOR CALOR ({anoFocal})
                  </span>
                  <h2 className="text-xl font-black text-red-400 mt-2">
                    {piorMesObj ? `${piorMesObj.mes}` : '--'}
                  </h2>
                  <p className="text-[10px] text-gray-400 mt-2">
                    Total no mês:{' '}
                    <strong>
                      {piorMesObj ? piorMesObj.qtdFocal : 0} sinistros
                    </strong>
                  </p>
                </div>

                <div className="bg-[#1C2541] border border-gray-800 p-5 rounded-2xl shadow-md">
                  <span className="text-xs font-bold text-gray-400">
                    PRINCIPAL CAUSA RAÍZ ({anoFocal})
                  </span>
                  <h2 className="text-sm font-black text-orange-400 mt-2 truncate">
                    {rankingCausas[0] ? rankingCausas[0][0] : '--'}
                  </h2>
                  <p className="text-[10px] text-gray-400 mt-2">
                    Ocorrências:{' '}
                    <strong>
                      {rankingCausas[0] ? rankingCausas[0][1] : 0} registros
                    </strong>
                  </p>
                </div>

                <div className="bg-[#1C2541] border border-gray-800 p-5 rounded-2xl shadow-md">
                  <span className="text-xs font-bold text-gray-400">
                    TOTAL ANALISADO ({anoFocal})
                  </span>
                  <h2 className="text-3xl font-black text-blue-400 mt-2">
                    {baseSemCulpabilidade.length}
                  </h2>
                  <p className="text-[10px] text-gray-400 mt-2">
                    Setor: <strong>{filtroSetor}</strong> • Cat:{' '}
                    <strong>{filtroCategoriaVeiculo}</strong>
                  </p>
                </div>
              </div>

              {/* 🎯 CARDS INTERATIVOS DE CULPABILIDADE */}
              <div className="bg-[#1C2541] border border-gray-800 p-6 rounded-2xl mb-8 shadow-xl">
                <div className="flex justify-between items-center mb-4 border-b border-gray-800 pb-3">
                  <div className="flex items-center gap-2">
                    <UserCheck className="text-orange-400" size={18} />
                    <h3 className="text-sm font-black text-white uppercase tracking-wider">
                      Distribuição da Culpabilidade dos Sinistros ({anoFocal})
                    </h3>
                  </div>
                  {filtroCulpabilidade !== 'TODAS' && (
                    <button
                      onClick={() => setFiltroCulpabilidade('TODAS')}
                      className="bg-red-500/20 hover:bg-red-500/30 text-red-400 border border-red-500/30 text-xs font-black px-3 py-1 rounded-xl flex items-center gap-1.5 transition-all cursor-pointer"
                    >
                      <FilterX size={13} /> Limpar Filtro ({filtroCulpabilidade}
                      )
                    </button>
                  )}
                </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  <div
                    onClick={() =>
                      selecionarCulpabilidadeCard('MOTORISTA PEPSICO')
                    }
                    className={`p-4 rounded-xl cursor-pointer transition-all border ${
                      filtroCulpabilidade === 'MOTORISTA PEPSICO'
                        ? 'bg-red-500/20 border-red-500 shadow-xl scale-[1.02]'
                        : 'bg-[#0B1329] border-gray-800 hover:border-gray-700 hover:bg-[#0B1329]/90'
                    }`}
                  >
                    <div className="flex justify-between items-center text-xs font-bold mb-2">
                      <span className="text-red-400 font-black">
                        MOTORISTA PEPSICO
                      </span>
                      <span className="text-white font-black text-sm font-mono">
                        {(
                          (contagemCulpabilidade['MOTORISTA PEPSICO'] /
                            totalCulpGeral) *
                          100
                        ).toFixed(1)}
                        %
                      </span>
                    </div>
                    <div className="w-full h-2.5 bg-gray-800 rounded-full overflow-hidden mb-3">
                      <div
                        className="h-full bg-red-500 rounded-full transition-all duration-500"
                        style={{
                          width: `${
                            (contagemCulpabilidade['MOTORISTA PEPSICO'] /
                              totalCulpGeral) *
                            100
                          }%`,
                        }}
                      ></div>
                    </div>
                    <div className="flex justify-between items-center text-[11px] text-gray-400">
                      <span>
                        Total:{' '}
                        <strong className="text-white">
                          {contagemCulpabilidade['MOTORISTA PEPSICO']}
                        </strong>{' '}
                        sinistros
                      </span>
                      <span className="text-[10px] text-red-400 font-bold bg-red-500/10 px-2 py-0.5 rounded">
                        {filtroCulpabilidade === 'MOTORISTA PEPSICO'
                          ? 'Filtrado ✓'
                          : 'Clique p/ Filtrar'}
                      </span>
                    </div>
                  </div>

                  <div
                    onClick={() =>
                      selecionarCulpabilidadeCard('TERCEIRO IDENTIFICADO')
                    }
                    className={`p-4 rounded-xl cursor-pointer transition-all border ${
                      filtroCulpabilidade === 'TERCEIRO IDENTIFICADO'
                        ? 'bg-blue-500/20 border-blue-500 shadow-xl scale-[1.02]'
                        : 'bg-[#0B1329] border-gray-800 hover:border-gray-700 hover:bg-[#0B1329]/90'
                    }`}
                  >
                    <div className="flex justify-between items-center text-xs font-bold mb-2">
                      <span className="text-blue-400 font-black">
                        TERCEIRO IDENTIFICADO
                      </span>
                      <span className="text-white font-black text-sm font-mono">
                        {(
                          (contagemCulpabilidade['TERCEIRO IDENTIFICADO'] /
                            totalCulpGeral) *
                          100
                        ).toFixed(1)}
                        %
                      </span>
                    </div>
                    <div className="w-full h-2.5 bg-gray-800 rounded-full overflow-hidden mb-3">
                      <div
                        className="h-full bg-blue-500 rounded-full transition-all duration-500"
                        style={{
                          width: `${
                            (contagemCulpabilidade['TERCEIRO IDENTIFICADO'] /
                              totalCulpGeral) *
                            100
                          }%`,
                        }}
                      ></div>
                    </div>
                    <div className="flex justify-between items-center text-[11px] text-gray-400">
                      <span>
                        Total:{' '}
                        <strong className="text-white">
                          {contagemCulpabilidade['TERCEIRO IDENTIFICADO']}
                        </strong>{' '}
                        sinistros
                      </span>
                      <span className="text-[10px] text-blue-400 font-bold bg-blue-500/10 px-2 py-0.5 rounded">
                        {filtroCulpabilidade === 'TERCEIRO IDENTIFICADO'
                          ? 'Filtrado ✓'
                          : 'Clique p/ Filtrar'}
                      </span>
                    </div>
                  </div>

                  <div
                    onClick={() =>
                      selecionarCulpabilidadeCard('TERCEIRO NÃO IDENTIFICADO')
                    }
                    className={`p-4 rounded-xl cursor-pointer transition-all border ${
                      filtroCulpabilidade === 'TERCEIRO NÃO IDENTIFICADO'
                        ? 'bg-yellow-500/20 border-yellow-500 shadow-xl scale-[1.02]'
                        : 'bg-[#0B1329] border-gray-800 hover:border-gray-700 hover:bg-[#0B1329]/90'
                    }`}
                  >
                    <div className="flex justify-between items-center text-xs font-bold mb-2">
                      <span className="text-yellow-400 font-black">
                        TERCEIRO NÃO IDENTIFICADO
                      </span>
                      <span className="text-white font-black text-sm font-mono">
                        {(
                          (contagemCulpabilidade['TERCEIRO NÃO IDENTIFICADO'] /
                            totalCulpGeral) *
                          100
                        ).toFixed(1)}
                        %
                      </span>
                    </div>
                    <div className="w-full h-2.5 bg-gray-800 rounded-full overflow-hidden mb-3">
                      <div
                        className="h-full bg-yellow-500 rounded-full transition-all duration-500"
                        style={{
                          width: `${
                            (contagemCulpabilidade[
                              'TERCEIRO NÃO IDENTIFICADO'
                            ] /
                              totalCulpGeral) *
                            100
                          }%`,
                        }}
                      ></div>
                    </div>
                    <div className="flex justify-between items-center text-[11px] text-gray-400">
                      <span>
                        Total:{' '}
                        <strong className="text-white">
                          {contagemCulpabilidade['TERCEIRO NÃO IDENTIFICADO']}
                        </strong>{' '}
                        sinistros
                      </span>
                      <span className="text-[10px] text-yellow-400 font-bold bg-yellow-500/10 px-2 py-0.5 rounded">
                        {filtroCulpabilidade === 'TERCEIRO NÃO IDENTIFICADO'
                          ? 'Filtrado ✓'
                          : 'Clique p/ Filtrar'}
                      </span>
                    </div>
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-8">
                <div className="bg-[#1C2541] border border-gray-800 p-6 rounded-2xl shadow-xl">
                  <div className="flex justify-between items-center mb-6">
                    <div>
                      <h3 className="text-sm font-black text-white uppercase flex items-center gap-2">
                        Evolução Mês a Mês ({anoFocal} vs {anoComparacao})
                      </h3>
                      <p className="text-xs text-gray-400 mt-0.5">
                        {filtroCulpabilidade !== 'TODAS'
                          ? `Mostrando apenas: ${filtroCulpabilidade}`
                          : 'Comparativo direto de volume exato'}
                      </p>
                    </div>
                    <div className="flex items-center gap-3 text-[10px] font-bold">
                      <span className="flex items-center gap-1 text-gray-400">
                        <span className="w-2.5 h-2.5 bg-gray-600 rounded-full"></span>{' '}
                        {anoComparacao}
                      </span>
                      <span className="flex items-center gap-1 text-orange-400">
                        <span className="w-2.5 h-2.5 bg-orange-500 rounded-full"></span>{' '}
                        {anoFocal}
                      </span>
                    </div>
                  </div>

                  <div className="h-60 flex items-end justify-between gap-1 pt-8 border-b border-gray-800 px-1">
                    {dadosEvolucaoMensal.map((item) => {
                      const maxVal = Math.max(
                        ...dadosEvolucaoMensal.map((d) =>
                          Math.max(d.qtdFocal, d.qtdComp)
                        ),
                        1
                      );
                      const hComp = (item.qtdComp / maxVal) * 100;
                      const hFocal = (item.qtdFocal / maxVal) * 100;

                      return (
                        <div
                          key={item.mes}
                          className="flex-1 flex flex-col items-center h-full justify-end group"
                        >
                          <div className="w-full flex items-end justify-center gap-1 h-full">
                            <div
                              className="w-2.5 bg-gray-600 rounded-t flex flex-col items-center justify-start relative transition-all duration-300"
                              style={{ height: `${Math.max(hComp, 6)}%` }}
                            >
                              <span className="absolute -top-4 text-[9px] font-black text-gray-400">
                                {item.qtdComp > 0 ? item.qtdComp : ''}
                              </span>
                            </div>
                            <div
                              className="w-2.5 bg-orange-500 rounded-t flex flex-col items-center justify-start relative group-hover:brightness-125 transition-all duration-300"
                              style={{ height: `${Math.max(hFocal, 6)}%` }}
                            >
                              <span className="absolute -top-4 text-[9px] font-black text-orange-400">
                                {item.qtdFocal > 0 ? item.qtdFocal : ''}
                              </span>
                            </div>
                          </div>
                          <span className="text-[9px] font-bold text-gray-400 mt-3">
                            {item.mes}
                          </span>
                        </div>
                      );
                    })}
                  </div>
                </div>

                <div className="bg-[#1C2541] border border-gray-800 p-6 rounded-2xl shadow-xl">
                  <div className="flex justify-between items-center mb-6">
                    <div>
                      <h3 className="text-sm font-black text-white uppercase">
                        Top 5 Filiais mais Críticas ({anoFocal})
                      </h3>
                      <p className="text-xs text-gray-400">
                        {apenasCDVsNoRanking
                          ? 'Exibindo apenas CDVs (Ocultando Carretas/Frete)'
                          : 'Exibindo Todas Unidades + Frete Carretas'}
                      </p>
                    </div>
                    <button
                      onClick={() =>
                        setApenasCDVsNoRanking(!apenasCDVsNoRanking)
                      }
                      className="text-[10px] font-black px-2.5 py-1 rounded-lg bg-[#0B1329] border border-gray-700 text-orange-400 hover:border-orange-500 cursor-pointer transition-all"
                    >
                      {apenasCDVsNoRanking
                        ? 'Ver Geral (+Frete)'
                        : 'Apenas CDVs'}
                    </button>
                  </div>

                  <div className="space-y-4">
                    {rankingFiliais.length === 0 ? (
                      <p className="text-xs text-gray-400 py-8 text-center">
                        Nenhum registro para este filtro.
                      </p>
                    ) : (
                      rankingFiliais.map(([filialNome, qtd]) => (
                        <div key={filialNome} className="text-xs">
                          <div className="flex justify-between font-bold mb-1">
                            <span className="text-gray-300 truncate">
                              {filialNome}
                            </span>
                            <span className="text-red-400 font-black">
                              {qtd} Ocorrências
                            </span>
                          </div>
                          <div className="w-full h-2.5 bg-gray-800 rounded-full overflow-hidden">
                            <div
                              className="h-full bg-gradient-to-r from-orange-500 to-red-500 rounded-full transition-all duration-500"
                              style={{
                                width: `${(qtd / maxSinistrosFilial) * 100}%`,
                              }}
                            ></div>
                          </div>
                        </div>
                      ))
                    )}
                  </div>
                </div>
              </div>

              <div className="bg-[#1C2541] border border-gray-800 rounded-2xl p-6 mb-8 shadow-xl">
                <h3 className="text-sm font-black text-white uppercase mb-4">
                  Ranking de Causas de Acidentes ({anoFocal})
                </h3>
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-5 gap-3">
                  {rankingCausas.length === 0 ? (
                    <p className="text-xs text-gray-400 col-span-5 text-center py-4">
                      Nenhum caso localizado para este filtro.
                    </p>
                  ) : (
                    rankingCausas.map(([causa, qtd], idx) => (
                      <div
                        key={causa}
                        className="bg-[#0B1329] border border-gray-800 p-3 rounded-xl"
                      >
                        <span className="text-[10px] font-bold text-orange-400">
                          #{idx + 1} TOP CAUSA
                        </span>
                        <p className="text-xs font-black text-white truncate mt-1">
                          {causa}
                        </p>
                        <p className="text-lg font-black text-gray-200 mt-2">
                          {qtd}{' '}
                          <span className="text-[10px] font-normal text-gray-400">
                            casos
                          </span>
                        </p>
                      </div>
                    ))
                  )}
                </div>
              </div>

              <div className="bg-[#1C2541] border border-gray-800 rounded-2xl overflow-hidden shadow-xl">
                <div className="p-4 border-b border-gray-800 flex justify-between items-center">
                  <div>
                    <h3 className="text-sm font-black text-white uppercase">
                      Lista Detalhada de Sinistros ({sinistrosFiltrados.length}{' '}
                      registros)
                    </h3>
                  </div>
                </div>

                <div className="overflow-x-auto max-h-[450px]">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-[#0B1329] text-gray-400 font-bold uppercase border-b border-gray-800 sticky top-0 z-10">
                      <tr>
                        <th className="p-3">Nº Sinistro</th>
                        <th className="p-3">Placa</th>
                        <th className="p-3">Veículo</th>
                        <th className="p-3">Cat.</th>
                        <th className="p-3">Setor (DTS/OT)</th>
                        <th className="p-3">Tipo Frota</th>
                        <th className="p-3">Filial / CDV</th>
                        <th className="p-3">Data</th>
                        <th className="p-3">Culpabilidade</th>
                        <th className="p-3">Tipo Acidente / Causa</th>
                        <th className="p-3">Motorista</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-800 text-gray-200 font-medium">
                      {sinistrosFiltrados.slice(0, 100).map((s) => (
                        <tr
                          key={s.id}
                          className="hover:bg-[#0B1329]/50 transition-all"
                        >
                          <td className="p-3 font-mono font-bold text-orange-400">
                            {s.numSinistro || 'N/A'}
                          </td>
                          <td className="p-3 font-bold">{s.placa}</td>
                          <td className="p-3 text-gray-300">
                            {s.veiculo || 'N/I'}
                          </td>
                          <td className="p-3">
                            <span
                              className={`text-[10px] font-black px-2 py-0.5 rounded ${
                                s.categoriaVeiculo === 'LEVE'
                                  ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/30'
                                  : s.categoriaVeiculo === 'REBOQUE'
                                  ? 'bg-purple-500/20 text-purple-400 border border-purple-500/30'
                                  : 'bg-orange-500/20 text-orange-400 border border-orange-500/30'
                              }`}
                            >
                              {s.categoriaVeiculo}
                            </span>
                          </td>
                          <td className="p-3 font-black text-cyan-400">
                            {s.setor}
                          </td>
                          <td className="p-3 text-gray-300 font-bold">
                            {s.tipoFrota}
                          </td>
                          <td className="p-3">{s.filial}</td>
                          <td className="p-3 text-gray-300">
                            {String(s.data)}
                          </td>
                          <td className="p-3">
                            <span
                              className={`text-[10px] font-black px-2 py-0.5 rounded ${
                                s.culpabilidade === 'MOTORISTA PEPSICO'
                                  ? 'bg-red-500/20 text-red-400 border border-red-500/30'
                                  : s.culpabilidade === 'TERCEIRO IDENTIFICADO'
                                  ? 'bg-blue-500/20 text-blue-400 border border-blue-500/30'
                                  : 'bg-yellow-500/20 text-yellow-400 border border-yellow-500/30'
                              }`}
                            >
                              {s.culpabilidade}
                            </span>
                          </td>
                          <td className="p-3 truncate max-w-xs">
                            {s.tipoAcidente}
                          </td>
                          <td className="p-3 text-gray-400 font-bold">
                            {s.motorista || 'N/I'}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </>
          )}
        </div>
      )}

      {/* 📝 TELA 3: CERTIFICAÇÃO DIREÇÃO SEGURA */}
      {telaAtiva === 'forms' && (
        <div className="p-8">
          <div className="mb-6 flex items-center justify-between bg-[#1C2541] border border-gray-800 p-3 rounded-2xl">
            <button
              onClick={() => setTelaAtiva('menu')}
              className="bg-[#0B1329] hover:bg-gray-800 text-xs font-black text-blue-400 px-4 py-2 rounded-xl flex items-center gap-2 border border-gray-700 transition-all cursor-pointer"
            >
              <ArrowLeft size={16} /> Voltar ao Menu Principal
            </button>
            <span className="text-xs font-black text-gray-400 uppercase tracking-widest hidden md:inline">
              Módulo: Certificação Direção Segura (SELL)
            </span>
          </div>

          <header className="sticky top-0 z-50 backdrop-blur-md bg-[#0B1329]/95 py-4 mb-8 border-b border-gray-800/80 shadow-2xl flex flex-col gap-4 transition-all">
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
              <div>
                <h1 className="text-2xl font-black text-white tracking-tight flex items-center gap-2">
                  <ClipboardCheck className="text-blue-500" />
                  Certificação Direção Segura | Auditoria & Performance
                </h1>
                <p className="text-xs text-gray-400 mt-0.5">
                  Análise Estratégica SELL • {colunasQuestoesForms.length}{' '}
                  Critérios Técnicos Auditados
                </p>
              </div>

              <div className="flex items-center gap-2.5">
                <button
                  onClick={carregarFormsNuvem}
                  className="bg-[#1C2541] hover:bg-gray-800 text-blue-400 border border-gray-700 text-xs font-black px-3.5 py-2 rounded-xl flex items-center gap-2 transition-all cursor-pointer shadow-sm"
                >
                  <RefreshCw
                    size={14}
                    className={carregandoNuvemForms ? 'animate-spin' : ''}
                  />
                  Sincronizar OneDrive
                </button>

                <button
                  onClick={limparFiltrosForms}
                  className="bg-[#1C2541] hover:bg-gray-800 border border-gray-700 text-gray-300 hover:text-white text-xs font-bold px-3 py-2 rounded-xl flex items-center gap-1.5 transition-all cursor-pointer shadow-sm"
                >
                  <RotateCcw size={14} className="text-blue-400" /> Limpar
                  Filtros
                </button>

                <label className="bg-blue-600 hover:bg-blue-700 text-white text-xs font-black px-4 py-2 rounded-xl cursor-pointer flex items-center gap-1.5 shadow-lg transition-all">
                  <Upload size={14} />{' '}
                  {carregadoForms ? 'Trocar Planilha' : 'Carregar Manual'}
                  <input
                    type="file"
                    accept=".xlsx, .xls"
                    onChange={lidarComUploadForms}
                    className="hidden"
                  />
                </label>
              </div>
            </div>

            {carregadoForms && (
              <div className="flex flex-wrap items-center gap-2.5 pt-2 border-t border-gray-800/80">
                <select
                  value={filtroAnoForms}
                  onChange={(e) => setFiltroAnoForms(e.target.value)}
                  className="bg-[#1C2541] border border-blue-500/40 text-xs font-black px-3 py-2 rounded-xl text-blue-400 outline-none cursor-pointer"
                >
                  <option value="TODOS">📅 Todos os Anos (2025 + 2026)</option>
                  <option value="2026">Ano: 2026</option>
                  <option value="2025">Ano: 2025</option>
                </select>

                <select
                  value={filtroMesForms}
                  onChange={(e) => setFiltroMesForms(e.target.value)}
                  className="bg-[#1C2541] border border-gray-800 text-xs font-bold px-3 py-2 rounded-xl text-gray-200 outline-none cursor-pointer"
                >
                  <option value="TODOS">Todos os Meses</option>
                  {[
                    { num: 1, nome: 'Janeiro' },
                    { num: 2, nome: 'Fevereiro' },
                    { num: 3, nome: 'Março' },
                    { num: 4, nome: 'Abril' },
                    { num: 5, nome: 'Maio' },
                    { num: 6, nome: 'Junho' },
                    { num: 7, nome: 'Julho' },
                    { num: 8, nome: 'Agosto' },
                    { num: 9, nome: 'Setembro' },
                    { num: 10, nome: 'Outubro' },
                    { num: 11, nome: 'Novembro' },
                    { num: 12, nome: 'Dezembro' },
                  ].map((m) => (
                    <option key={m.num} value={m.num}>
                      Mês: {m.nome}
                    </option>
                  ))}
                </select>

                <select
                  value={filtroCanalForms}
                  onChange={(e) => {
                    setFiltroCanalForms(e.target.value);
                    setFiltroDivisaoForms('TODAS');
                  }}
                  className="bg-[#1C2541] border border-cyan-500/40 text-xs font-black px-3 py-2 rounded-xl text-cyan-400 outline-none cursor-pointer"
                >
                  <option value="TODOS">🏢 Todos os Canais (OT + DTS)</option>
                  <option value="OT">Canal: OT (Vendas / Leves)</option>
                  <option value="DTS">Canal: DTS (Entrega / Pesados)</option>
                </select>

                <select
                  value={filtroDivisaoForms}
                  onChange={(e) => setFiltroDivisaoForms(e.target.value)}
                  className="bg-[#1C2541] border border-gray-800 text-xs font-bold px-3 py-2 rounded-xl text-amber-400 outline-none cursor-pointer"
                >
                  <option value="TODAS">📍 Todas as Divisões / Regiões</option>
                  {opcoesDivisaoForms.map((div) => (
                    <option key={div} value={div}>
                      Divisão: {div}
                    </option>
                  ))}
                </select>

                <select
                  value={filtroTipoCertForms}
                  onChange={(e) => setFiltroTipoCertForms(e.target.value)}
                  className="bg-[#1C2541] border border-gray-800 text-xs font-bold px-3 py-2 rounded-xl text-purple-400 outline-none cursor-pointer"
                >
                  <option value="TODOS">🎓 Todos Tipos de Certificação</option>
                  {opcoesTipoCertForms.map((t) => (
                    <option key={t} value={t}>
                      {t}
                    </option>
                  ))}
                </select>

                <select
                  value={filtroAuditoriaForms}
                  onChange={(e) => setFiltroAuditoriaForms(e.target.value)}
                  className="bg-[#1C2541] border border-red-500/40 text-xs font-black px-3 py-2 rounded-xl text-red-400 outline-none cursor-pointer"
                >
                  <option value="TODOS">🛡️ Auditoria: Todos Avaliados</option>
                  <option value="COM_DESVIO">
                    ⚠️ Apenas com Desvios / Reprovações
                  </option>
                  <option value="SEM_DESVIO">
                    ✅ Apenas 100% Conformes / Sem Falhas
                  </option>
                  <option value="CNH_ALERTA">
                    🪪 Alerta de CNH (Vencida / Data Inválida)
                  </option>
                </select>
              </div>
            )}
          </header>

          {!carregadoForms &&
          !carregandoNuvemForms &&
          dadosBaseForms.length === 0 ? (
            <div className="bg-[#1C2541] border border-dashed border-gray-700 rounded-3xl p-12 text-center max-w-2xl mx-auto my-12">
              <ClipboardCheck
                className="mx-auto text-blue-500 mb-4 animate-bounce"
                size={48}
              />
              <h2 className="text-xl font-black text-white">
                Sincronizador Dinâmico OneDrive
              </h2>
              <p className="text-xs text-gray-400 mt-2 max-w-md mx-auto">
                Buscando a planilha `CERTIFICAÇÃO DIREÇÃO SEGURA.xlsx` do
                OneDrive...
              </p>
              <div className="mt-6 flex justify-center gap-3">
                <button
                  onClick={carregarFormsNuvem}
                  className="bg-blue-600 hover:bg-blue-700 text-white font-black text-xs px-6 py-3 rounded-xl cursor-pointer shadow-lg transition-all flex items-center gap-2"
                >
                  <RefreshCw size={14} /> Tentar Sincronizar Novamente
                </button>
                <label className="bg-gray-800 hover:bg-gray-700 text-gray-200 font-black text-xs px-6 py-3 rounded-xl cursor-pointer shadow-lg transition-all">
                  Selecionar Arquivo .XLSX Manual
                  <input
                    type="file"
                    accept=".xlsx, .xls"
                    onChange={lidarComUploadForms}
                    className="hidden"
                  />
                </label>
              </div>
            </div>
          ) : (
            <>
              <div className="bg-[#1C2541] border border-gray-800 p-3 rounded-xl mb-6 flex items-center gap-3 shadow-md">
                <Search size={16} className="text-gray-400" />
                <input
                  type="text"
                  value={termoBuscaForms}
                  onChange={(e) => setTermoBuscaForms(e.target.value)}
                  placeholder="Pesquisar por Nome do Condutor, GPID do Condutor, GPID do Avaliador, Divisão ou Comentários..."
                  className="bg-transparent w-full text-xs text-white outline-none font-bold"
                />
              </div>

              <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-8">
                <div
                  onClick={() => setFiltroAuditoriaForms('TODOS')}
                  className="bg-[#1C2541] border border-gray-800 hover:border-blue-500/50 p-5 rounded-2xl shadow-md cursor-pointer transition-all"
                >
                  <span className="text-xs font-bold text-gray-400 uppercase flex items-center gap-1.5">
                    <ClipboardCheck size={14} className="text-blue-400" />{' '}
                    CERTIFICAÇÕES FILTRADAS
                  </span>
                  <h2 className="text-3xl font-black text-blue-400 mt-2">
                    {totalFormsExibidos}
                  </h2>
                  <p className="text-[10px] text-gray-400 mt-1">
                    De {dadosBaseForms.length} registros totais na planilha
                  </p>
                </div>

                <div
                  onClick={() =>
                    setFiltroAuditoriaForms(
                      filtroAuditoriaForms === 'SEM_DESVIO'
                        ? 'TODOS'
                        : 'SEM_DESVIO'
                    )
                  }
                  className={`bg-[#1C2541] border p-5 rounded-2xl shadow-md cursor-pointer transition-all ${
                    filtroAuditoriaForms === 'SEM_DESVIO'
                      ? 'border-emerald-500 bg-emerald-500/10'
                      : 'border-gray-800 hover:border-emerald-500/50'
                  }`}
                >
                  <span className="text-xs font-bold text-gray-400 uppercase flex items-center gap-1.5">
                    <Award size={14} className="text-emerald-400" /> 100%
                    CONFORMES (SEM DESVIO)
                  </span>
                  <h2 className="text-3xl font-black text-emerald-400 mt-2">
                    {taxaAprovacaoDireta}%
                  </h2>
                  <p className="text-[10px] text-gray-400 mt-1">
                    <strong>{totalSemDesvio}</strong> condutores sem nenhuma
                    falha
                  </p>
                </div>

                <div
                  onClick={() =>
                    setFiltroAuditoriaForms(
                      filtroAuditoriaForms === 'COM_DESVIO'
                        ? 'TODOS'
                        : 'COM_DESVIO'
                    )
                  }
                  className={`bg-[#1C2541] border p-5 rounded-2xl shadow-md cursor-pointer transition-all ${
                    filtroAuditoriaForms === 'COM_DESVIO'
                      ? 'border-red-500 bg-red-500/10'
                      : 'border-gray-800 hover:border-red-500/50'
                  }`}
                >
                  <span className="text-xs font-bold text-gray-400 uppercase flex items-center gap-1.5">
                    <AlertTriangle size={14} className="text-red-400" />{' '}
                    CONDUTORES COM DESVIOS
                  </span>
                  <h2 className="text-3xl font-black text-red-400 mt-2">
                    {totalComDesvio}
                  </h2>
                  <p className="text-[10px] text-gray-400 mt-1">
                    Somando <strong>{somaApontamentosDesvio}</strong>{' '}
                    reprovações (Clique p/ filtrar)
                  </p>
                </div>

                <div
                  onClick={() =>
                    setFiltroAuditoriaForms(
                      filtroAuditoriaForms === 'CNH_ALERTA'
                        ? 'TODOS'
                        : 'CNH_ALERTA'
                    )
                  }
                  className={`bg-[#1C2541] border p-5 rounded-2xl shadow-md cursor-pointer transition-all ${
                    filtroAuditoriaForms === 'CNH_ALERTA'
                      ? 'border-amber-500 bg-amber-500/10'
                      : 'border-gray-800 hover:border-amber-500/50'
                  }`}
                >
                  <span className="text-xs font-bold text-gray-400 uppercase flex items-center gap-1.5">
                    <CreditCard size={14} className="text-amber-400" /> ALERTA
                    VALIDADE CNH
                  </span>
                  <h2 className="text-3xl font-black text-amber-400 mt-2">
                    {totalAlertasCNH}
                  </h2>
                  <p className="text-[10px] text-gray-400 mt-1">
                    CNH vencida ou erro de digitação (Clique p/ ver)
                  </p>
                </div>
              </div>

              <div className="bg-[#1C2541] border border-gray-800 rounded-2xl p-6 mb-8 shadow-xl">
                <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-2 mb-4 border-b border-gray-800 pb-3">
                  <div>
                    <h3 className="text-sm font-black text-white uppercase flex items-center gap-2">
                      <AlertTriangle className="text-red-400" size={18} />
                      Top Regras com Desvios Operacionais (Não Eficiente / Não
                      Conforme)
                    </h3>
                    <p className="text-xs text-gray-400 mt-0.5">
                      Clique em qualquer regra abaixo para filtrar na tabela
                      exatamente quais condutores cometeram o desvio.
                    </p>
                  </div>
                  {questaoDesvioSelecionada && (
                    <button
                      onClick={() => setQuestaoDesvioSelecionada(null)}
                      className="bg-red-500/20 hover:bg-red-500/30 text-red-400 border border-red-500/30 text-xs font-black px-3 py-1.5 rounded-xl flex items-center gap-1.5 cursor-pointer"
                    >
                      <FilterX size={14} /> Remover Filtro de Questão
                    </button>
                  )}
                </div>

                {topQuestoesComFalha.length === 0 ? (
                  <p className="text-xs text-emerald-400 font-bold py-4 text-center">
                    🎉 Nenhum desvio ou não conformidade registrado para os
                    filtros selecionados!
                  </p>
                ) : (
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-3">
                    {topQuestoesComFalha.slice(0, 8).map((q, idx) => {
                      const ativa = questaoDesvioSelecionada === q.coluna;
                      return (
                        <div
                          key={q.coluna}
                          onClick={() =>
                            setQuestaoDesvioSelecionada(ativa ? null : q.coluna)
                          }
                          className={`p-3.5 rounded-xl border cursor-pointer transition-all flex flex-col justify-between ${
                            ativa
                              ? 'bg-red-500/20 border-red-500 scale-[1.02] shadow-lg'
                              : 'bg-[#0B1329] border-gray-800 hover:border-red-500/50'
                          }`}
                        >
                          <div>
                            <div className="flex justify-between items-center mb-1.5">
                              <span className="text-[10px] font-black text-red-400 bg-red-500/10 px-2 py-0.5 rounded border border-red-500/20">
                                #{idx + 1} MAIOR DESVIO
                              </span>
                              <span className="text-xs font-mono font-black text-white">
                                {q.desvios}{' '}
                                {q.desvios === 1 ? 'falha' : 'falhas'}
                              </span>
                            </div>
                            <p
                              className="text-xs font-bold text-gray-200 line-clamp-2 mt-1"
                              title={q.coluna}
                            >
                              {q.coluna}
                            </p>
                          </div>
                          <div className="mt-3 pt-2 border-t border-gray-800/80 flex justify-between items-center text-[10px] text-gray-400">
                            <span>Conformidade: {q.taxaConformidade}%</span>
                            <span className="text-red-400 font-bold">
                              {ativa ? 'Filtrado ✓' : 'Filtrar →'}
                            </span>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

              <div className="mb-8">
                <h3 className="text-base font-black text-white uppercase tracking-wide mb-6 flex items-center gap-2">
                  <BarChart3 className="text-blue-400" size={20} /> Indicadores
                  Chave da Certificação (Recalculados pelos Filtros)
                </h3>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  {graficosPrincipaisForms.map((item) => {
                    const dadosGrafico = item.dados;

                    return (
                      <div
                        key={item.titulo}
                        className="bg-[#1C2541] border border-gray-800 p-6 rounded-2xl shadow-xl flex flex-col justify-between relative overflow-hidden"
                      >
                        <div className="flex justify-between items-center mb-6">
                          <h4 className="text-sm font-black text-white uppercase tracking-wide">
                            {item.titulo}
                          </h4>
                          <span className="text-[10px] font-bold text-blue-400 bg-blue-500/10 px-2.5 py-1 rounded-full">
                            {formsFiltrados.length} avaliados
                          </span>
                        </div>

                        <div className="flex flex-col md:flex-row items-center justify-between gap-6">
                          <div className="w-full md:w-1/2 space-y-2.5 max-h-48 overflow-y-auto pr-1">
                            {dadosGrafico.map((g, gIdx) => {
                              const cor =
                                paletaFormsHex[gIdx % paletaFormsHex.length];
                              return (
                                <div
                                  key={g.opcao}
                                  className="flex items-center justify-between text-xs font-bold"
                                >
                                  <div className="flex items-center gap-2 truncate pr-2">
                                    <span
                                      className="w-3 h-3 rounded-full flex-shrink-0"
                                      style={{ backgroundColor: cor }}
                                    ></span>
                                    <span
                                      className="text-gray-300 truncate"
                                      title={g.opcao}
                                    >
                                      {g.opcao}
                                    </span>
                                  </div>
                                  <span className="text-white font-mono font-black">
                                    {g.qtd}{' '}
                                    <span className="text-[10px] text-gray-400 font-normal">
                                      ({g.porcentagem}%)
                                    </span>
                                  </span>
                                </div>
                              );
                            })}
                          </div>

                          <div className="w-full md:w-1/2 flex items-center justify-center">
                            {item.tipo === 'donut' ? (
                              <div className="relative w-40 h-40 flex items-center justify-center">
                                <svg
                                  className="w-full h-full -rotate-90"
                                  viewBox="0 0 36 36"
                                >
                                  {(() => {
                                    let acumulado = 0;
                                    return dadosGrafico.map((g, gIdx) => {
                                      const porc = Number(g.porcentagem);
                                      const cor =
                                        paletaFormsHex[
                                          gIdx % paletaFormsHex.length
                                        ];
                                      const strokeDasharray = `${porc} ${
                                        100 - porc
                                      }`;
                                      const strokeDashoffset = acumulado;
                                      acumulado -= porc;

                                      return (
                                        <circle
                                          key={g.opcao}
                                          cx="18"
                                          cy="18"
                                          r="14"
                                          fill="transparent"
                                          stroke={cor}
                                          strokeWidth="5"
                                          strokeDasharray={strokeDasharray}
                                          strokeDashoffset={strokeDashoffset}
                                        />
                                      );
                                    });
                                  })()}
                                </svg>
                                <div className="absolute text-center">
                                  <span className="text-lg font-black text-white font-mono">
                                    {dadosGrafico[0]?.porcentagem || 0}%
                                  </span>
                                </div>
                              </div>
                            ) : (
                              <div className="w-full space-y-2 max-h-48 overflow-y-auto pr-1">
                                {dadosGrafico.slice(0, 8).map((g, gIdx) => {
                                  const cor =
                                    paletaFormsHex[
                                      gIdx % paletaFormsHex.length
                                    ];
                                  return (
                                    <div key={g.opcao} className="w-full">
                                      <div className="flex justify-between text-[10px] text-gray-400 mb-0.5">
                                        <span className="truncate">
                                          {g.opcao}
                                        </span>
                                        <span>{g.porcentagem}%</span>
                                      </div>
                                      <div className="w-full h-2.5 bg-gray-800 rounded-sm overflow-hidden">
                                        <div
                                          className="h-full rounded-sm transition-all duration-500"
                                          style={{
                                            width: `${g.porcentagem}%`,
                                            backgroundColor: cor,
                                          }}
                                        ></div>
                                      </div>
                                    </div>
                                  );
                                })}
                              </div>
                            )}
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              <div className="bg-[#1C2541] border border-gray-800 rounded-2xl mb-8 overflow-hidden shadow-xl">
                <button
                  onClick={() =>
                    setExpandirPerguntasForms(!expandirPerguntasForms)
                  }
                  className="w-full p-5 flex justify-between items-center text-left hover:bg-[#0B1329]/40 transition-all cursor-pointer"
                >
                  <div className="flex items-center gap-3">
                    <BarChart3 className="text-amber-400" size={20} />
                    <div>
                      <h3 className="text-sm font-black text-white uppercase tracking-wider">
                        Auditoria Detalhada das {colunasQuestoesForms.length}{' '}
                        Perguntas Técnicas (Método Smith & Segurança)
                      </h3>
                      <p className="text-xs text-gray-400 mt-0.5">
                        {expandirPerguntasForms
                          ? 'Clique para recolher os gráficos individuais das 24 perguntas'
                          : 'Clique para expandir e ver a taxa de Eficiência / Conformidade de cada pergunta'}
                      </p>
                    </div>
                  </div>

                  <div className="bg-[#0B1329] p-2 rounded-xl text-amber-400 border border-gray-700">
                    {expandirPerguntasForms ? (
                      <ChevronUp size={20} />
                    ) : (
                      <ChevronDown size={20} />
                    )}
                  </div>
                </button>

                {expandirPerguntasForms && (
                  <div className="p-6 border-t border-gray-800 bg-[#0B1329]/50 grid grid-cols-1 md:grid-cols-2 gap-6">
                    {colunasQuestoesForms.map((col, index) => {
                      const dadosGrafico = calcularGraficoCampo((r) => r[col]);

                      return (
                        <div
                          key={col}
                          className="bg-[#1C2541] border border-gray-800 p-5 rounded-2xl flex flex-col justify-between shadow-lg"
                        >
                          <div>
                            <div className="flex justify-between items-start gap-2 mb-3 border-b border-gray-800/80 pb-3">
                              <span className="text-[10px] font-black uppercase text-blue-400 bg-blue-500/10 px-2.5 py-1 rounded-md border border-blue-500/20">
                                Critério Técnico #{index + 1}
                              </span>
                              <span className="text-[10px] font-bold text-gray-400">
                                {formsFiltrados.length} Avaliados
                              </span>
                            </div>

                            <h4 className="text-xs font-black text-white mb-4 uppercase tracking-wide">
                              {col}
                            </h4>

                            <div className="space-y-3">
                              {dadosGrafico.map((g) => {
                                const ehDesvio =
                                  g.opcao.includes('NÃO EFICI') ||
                                  g.opcao.includes('NÃO CONFORME');
                                const cor = ehDesvio
                                  ? '#EF4444'
                                  : g.opcao.includes('APLICA')
                                  ? '#6B7280'
                                  : '#10B981';

                                return (
                                  <div key={g.opcao} className="text-[11px]">
                                    <div className="flex justify-between text-gray-300 font-bold mb-1">
                                      <span className="truncate pr-2">
                                        {g.opcao}
                                      </span>
                                      <span className="text-white font-mono font-black">
                                        {g.qtd}{' '}
                                        <span className="text-[10px] font-normal text-gray-400">
                                          ({g.porcentagem}%)
                                        </span>
                                      </span>
                                    </div>
                                    <div className="w-full h-2.5 bg-gray-800 rounded-full overflow-hidden p-0.5">
                                      <div
                                        className="h-full rounded-full transition-all duration-500"
                                        style={{
                                          width: `${g.porcentagem}%`,
                                          backgroundColor: cor,
                                        }}
                                      ></div>
                                    </div>
                                  </div>
                                );
                              })}
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

              <div className="bg-[#1C2541] border border-gray-800 rounded-2xl overflow-hidden shadow-xl">
                <div className="p-4 border-b border-gray-800 flex flex-col md:flex-row justify-between items-start md:items-center gap-2">
                  <div>
                    <h3 className="text-sm font-black text-white uppercase">
                      Relatório de Condutores Avaliados ({formsFiltrados.length}{' '}
                      registros)
                    </h3>
                  </div>
                </div>

                <div className="overflow-x-auto max-h-[550px]">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-[#0B1329] text-gray-400 font-bold uppercase border-b border-gray-800 sticky top-0 z-10">
                      <tr>
                        <th className="p-3">Data Cert.</th>
                        <th className="p-3">Condutor</th>
                        <th className="p-3">GPID Condutor</th>
                        <th className="p-3">Canal</th>
                        <th className="p-3">Divisão / Regional</th>
                        <th className="p-3">Tipo de Certificação</th>
                        <th className="p-3">Validade CNH</th>
                        <th className="p-3 text-center">Status Auditoria</th>
                        <th className="p-3">Comentários do Avaliador</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-800 text-gray-200 font-medium">
                      {formsFiltrados.map((linha) => (
                        <tr
                          key={linha._id}
                          className="hover:bg-[#0B1329]/50 transition-all"
                        >
                          <td className="p-3 font-mono text-gray-300 whitespace-nowrap">
                            {linha._dataCertFormatada}
                          </td>
                          <td className="p-3 font-black text-white whitespace-nowrap">
                            {linha._condutor}
                          </td>
                          <td className="p-3 font-mono text-blue-400 font-bold whitespace-nowrap">
                            {linha._gpidCondutor}
                          </td>
                          <td className="p-3 whitespace-nowrap">
                            <span
                              className={`px-2 py-0.5 rounded text-[10px] font-black border ${
                                linha._canal === 'OT'
                                  ? 'bg-cyan-500/20 text-cyan-400 border-cyan-500/30'
                                  : 'bg-orange-500/20 text-orange-400 border-orange-500/30'
                              }`}
                            >
                              {linha._canal}
                            </span>
                          </td>
                          <td className="p-3 font-bold text-amber-400 whitespace-nowrap">
                            {linha._divisao}
                          </td>
                          <td className="p-3 whitespace-nowrap text-gray-300">
                            {linha._tipoCert}
                          </td>
                          <td className="p-3 font-mono whitespace-nowrap">
                            <span
                              className={
                                linha._cnhAlerta
                                  ? 'text-amber-400 font-black bg-amber-500/10 px-2 py-0.5 rounded border border-amber-500/30'
                                  : 'text-gray-300'
                              }
                            >
                              {linha._dataCNHFormatada}{' '}
                              {linha._cnhAlerta ? '⚠️' : ''}
                            </span>
                          </td>
                          <td className="p-3 text-center whitespace-nowrap">
                            {linha._conformeGeral ? (
                              <span className="bg-emerald-500/20 text-emerald-400 border border-emerald-500/30 font-black text-[10px] px-2.5 py-1 rounded-full">
                                ✅ 100% Conforme
                              </span>
                            ) : (
                              <span
                                className="bg-red-500/20 text-red-400 border border-red-500/30 font-black text-[10px] px-2.5 py-1 rounded-full cursor-help"
                                title={linha._listaDesvios.join(' | ')}
                              >
                                🚨 {linha._qtdDesvios}{' '}
                                {linha._qtdDesvios === 1 ? 'Desvio' : 'Desvios'}
                              </span>
                            )}
                          </td>
                          <td
                            className="p-3 text-gray-400 max-w-xs truncate"
                            title={linha._comentarios}
                          >
                            {linha._comentarios || '-'}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </>
          )}
        </div>
      )}

      {/* 🚗 TELA 4: MÓDULO EXCLUSIVO DE FROTA LEVE ATIVA (DTS LEVE vs OT) */}
      {telaAtiva === 'frota_leves' && (
        <div className="p-8">
          <div className="mb-6 flex items-center justify-between bg-[#1C2541] border border-gray-800 p-3 rounded-2xl">
            <button
              onClick={() => setTelaAtiva('menu')}
              className="bg-[#0B1329] hover:bg-gray-800 text-xs font-black text-emerald-400 px-4 py-2 rounded-xl flex items-center gap-2 border border-gray-700 transition-all cursor-pointer"
            >
              <ArrowLeft size={16} /> Voltar ao Menu Principal
            </button>
            <span className="text-xs font-black text-gray-400 uppercase tracking-widest hidden md:inline">
              Módulo 3: Controle de Frota Leve Ativa (DTS Leve & OT)
            </span>
          </div>

          <header className="sticky top-0 z-50 backdrop-blur-md bg-[#0B1329]/95 py-4 mb-8 border-b border-gray-800/80 shadow-2xl flex flex-col gap-4 transition-all">
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
              <div>
                <h1 className="text-2xl font-black text-white tracking-tight flex items-center gap-2">
                  <Car className="text-emerald-400" />
                  Base Oficial de Leves Ativos | Separação DTS Leve vs. OT
                </h1>
                <p className="text-xs text-gray-400 mt-0.5">
                  Higienização automática da Posição Nacional (Aba Ativos) •
                  Mapeamento por Diretoria, Centro de Custo, Regional e Gerente
                </p>
              </div>

              <div className="flex flex-wrap items-center gap-2.5">
                {carregadoFrotaAtiva && (
                  <button
                    onClick={exportarBaseLevesExcel}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-black px-4 py-2 rounded-xl flex items-center gap-1.5 shadow-lg transition-all cursor-pointer"
                  >
                    <Download size={14} /> Exportar Base Limpa (.XLSX)
                  </button>
                )}

                <button
                  onClick={() => {
                    setFiltroSegmentoLeve('TODOS');
                    setFiltroAreaLeve('TODAS');
                    setFiltroGerenteLeve('TODOS');
                    setFiltroTipoFrotaLeve('OPERACIONAL');
                    setTermoBuscaLeves('');
                  }}
                  className="bg-[#1C2541] hover:bg-gray-800 border border-gray-700 text-gray-300 hover:text-white text-xs font-bold px-3 py-2 rounded-xl flex items-center gap-1.5 transition-all cursor-pointer"
                >
                  <RotateCcw size={14} className="text-emerald-400" /> Limpar
                  Filtros
                </button>

                <label className="bg-cyan-600 hover:bg-cyan-700 text-white text-xs font-black px-4 py-2 rounded-xl cursor-pointer flex items-center gap-1.5 shadow-lg transition-all">
                  <Upload size={14} />{' '}
                  {carregadoFrotaAtiva
                    ? 'Trocar Posição Nacional'
                    : 'Carregar Posição Nacional (.XLSX)'}
                  <input
                    type="file"
                    accept=".xlsx, .xls"
                    onChange={lidarComUploadPosicaoNacional}
                    className="hidden"
                  />
                </label>
              </div>
            </div>

            {carregadoFrotaAtiva && (
              <div className="flex flex-wrap items-center gap-2.5 pt-2 border-t border-gray-800/80">
                {/* Escopo da Frota Leve */}
                <select
                  value={filtroTipoFrotaLeve}
                  onChange={(e) => setFiltroTipoFrotaLeve(e.target.value)}
                  className="bg-[#1C2541] border border-emerald-500/40 text-xs font-black px-3 py-2 rounded-xl text-emerald-400 outline-none cursor-pointer"
                >
                  <option value="OPERACIONAL">
                    🚗 Apenas Leves Operacionais (Locados + ALD Flex + Temp)
                  </option>
                  <option value="COM_EXECUTIVO">
                    👔 Incluir Frota Executiva (Diretores/Gerentes)
                  </option>
                </select>

                {/* Filtro Segmento Oficial */}
                <select
                  value={filtroSegmentoLeve}
                  onChange={(e) => setFiltroSegmentoLeve(e.target.value)}
                  className="bg-[#1C2541] border border-cyan-500/40 text-xs font-black px-3 py-2 rounded-xl text-cyan-400 outline-none cursor-pointer"
                >
                  <option value="TODOS">
                    🎯 Todos os Segmentos (DTS Leve + OT + Suporte)
                  </option>
                  <option value="OT">🏢 Apenas OT (Organized Trade)</option>
                  <option value="DTS LEVE">
                    🚙 Apenas DTS LEVE (Vendas/Entregas Leves)
                  </option>
                  <option value="SUPORTE / OUTROS">
                    🛠️ Apenas Suporte (EHS / Agro / Transp / Whs)
                  </option>
                  {filtroTipoFrotaLeve === 'COM_EXECUTIVO' && (
                    <option value="EXECUTIVO">👔 Apenas Frota Executiva</option>
                  )}
                </select>

                {/* Filtro Área / Regional */}
                <select
                  value={filtroAreaLeve}
                  onChange={(e) => setFiltroAreaLeve(e.target.value)}
                  className="bg-[#1C2541] border border-gray-800 text-xs font-bold px-3 py-2 rounded-xl text-amber-400 outline-none cursor-pointer"
                >
                  <option value="TODAS">📍 Todas as Áreas / Regionais</option>
                  {opcoesAreaLeve.map((a) => (
                    <option key={a} value={a}>
                      Área: {a}
                    </option>
                  ))}
                </select>

                {/* Filtro Gerente Regional */}
                <select
                  value={filtroGerenteLeve}
                  onChange={(e) => setFiltroGerenteLeve(e.target.value)}
                  className="bg-[#1C2541] border border-gray-800 text-xs font-bold px-3 py-2 rounded-xl text-purple-400 outline-none cursor-pointer"
                >
                  <option value="TODOS">👤 Todos os Gerentes Regionais</option>
                  {opcoesGerenteLeve.map((g) => (
                    <option key={g} value={g}>
                      Gerente: {g}
                    </option>
                  ))}
                </select>

                {/* Botão de Regra de Negócio: Somar Transportes/Warehouse ao DTS Leve */}
                <button
                  onClick={() =>
                    setIncluirTranspWhsNoDts(!incluirTranspWhsNoDts)
                  }
                  className={`text-xs font-black px-3 py-2 rounded-xl border transition-all cursor-pointer ${
                    incluirTranspWhsNoDts
                      ? 'bg-amber-500/20 border-amber-500 text-amber-300'
                      : 'bg-[#1C2541] border-gray-700 text-gray-400 hover:text-white'
                  }`}
                  title="Define se os 18 carros leves de Transportes e Warehouse nos CDVs contam dentro de DTS Leve ou em Suporte"
                >
                  {incluirTranspWhsNoDts
                    ? '✓ Transp/Warehouse somados no DTS Leve'
                    : '+ Somar Transp/Warehouse no DTS Leve'}
                </button>
              </div>
            )}
          </header>

          {!carregadoFrotaAtiva && !carregandoNuvemPosicao ? (
            <div className="bg-[#1C2541] border border-dashed border-gray-700 rounded-3xl p-12 text-center max-w-2xl mx-auto my-12">
              <Car
                className="mx-auto text-emerald-400 mb-4 animate-bounce"
                size={48}
              />
              <h2 className="text-xl font-black text-white">
                Carregue a Planilha "Posição Nacional Frota"
              </h2>
              <p className="text-xs text-gray-400 mt-2 max-w-md mx-auto leading-relaxed">
                Selecione o arquivo{' '}
                <strong>
                  10.Posição Nacional Frota Outubro Semana 4 2026.xlsx
                </strong>{' '}
                para o sistema ler a aba <strong>Ativos</strong>, corrigir as
                diretorias automaticamente e separar quem é{' '}
                <strong>DTS Leve</strong> e quem é <strong>OT</strong>.
              </p>
              <div className="mt-6 flex justify-center gap-3">
                <label className="bg-emerald-600 hover:bg-emerald-700 text-white font-black text-xs px-6 py-3 rounded-xl cursor-pointer shadow-lg transition-all flex items-center gap-2">
                  <Upload size={15} /> Carregar Posição Nacional (.XLSX)
                  <input
                    type="file"
                    accept=".xlsx, .xls"
                    onChange={lidarComUploadPosicaoNacional}
                    className="hidden"
                  />
                </label>
              </div>
            </div>
          ) : (
            <>
              {/* 🔍 BARRA DE BUSCA RÁPIDA */}
              <div className="bg-[#1C2541] border border-gray-800 p-3 rounded-xl mb-6 flex items-center gap-3 shadow-md">
                <Search size={16} className="text-gray-400" />
                <input
                  type="text"
                  value={termoBuscaLeves}
                  onChange={(e) => setTermoBuscaLeves(e.target.value)}
                  placeholder="Pesquisar por Placa, Código (LV/LK/LF/LT), Modelo, Diretoria, Centro de Custo, CDV ou Condutor..."
                  className="bg-transparent w-full text-xs text-white outline-none font-bold"
                />
              </div>

              {/* 🎯 4 CARDS EXECUTIVOS INTERATIVOS (SEPARAÇÃO OFICIAL DTS LEVE vs OT) */}
              <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-8">
                <div
                  onClick={() => setFiltroSegmentoLeve('TODOS')}
                  className={`bg-[#1C2541] border p-5 rounded-2xl shadow-md cursor-pointer transition-all ${
                    filtroSegmentoLeve === 'TODOS'
                      ? 'border-white/60 bg-white/5'
                      : 'border-gray-800 hover:border-gray-600'
                  }`}
                >
                  <span className="text-xs font-bold text-gray-400 uppercase flex items-center gap-1.5">
                    <Car size={14} className="text-white" /> TOTAL LEVES ATIVOS
                  </span>
                  <h2 className="text-3xl font-black text-white mt-2">
                    {baseLevesSemFiltroSegmento.length}
                  </h2>
                  <p className="text-[10px] text-gray-400 mt-1">
                    {filtroTipoFrotaLeve === 'OPERACIONAL'
                      ? 'Locados (695) + ALD Flex (47) + Temp (41)'
                      : 'Incluindo Frota Executiva'}
                  </p>
                </div>

                <div
                  onClick={() =>
                    setFiltroSegmentoLeve(
                      filtroSegmentoLeve === 'OT' ? 'TODOS' : 'OT'
                    )
                  }
                  className={`bg-[#1C2541] border p-5 rounded-2xl shadow-md cursor-pointer transition-all ${
                    filtroSegmentoLeve === 'OT'
                      ? 'border-cyan-500 bg-cyan-500/15 scale-[1.02]'
                      : 'border-gray-800 hover:border-cyan-500/50'
                  }`}
                >
                  <div className="flex justify-between items-center">
                    <span className="text-xs font-black text-cyan-400 uppercase">
                      🏢 TOTAL OT (ORGANIZED TRADE)
                    </span>
                    <span className="text-xs font-mono font-black text-white">
                      {baseLevesSemFiltroSegmento.length > 0
                        ? (
                            (resumoSegmentosLeves['OT'] /
                              baseLevesSemFiltroSegmento.length) *
                            100
                          ).toFixed(1)
                        : 0}
                      %
                    </span>
                  </div>
                  <h2 className="text-3xl font-black text-cyan-400 mt-2">
                    {resumoSegmentosLeves['OT']}
                  </h2>
                  <p className="text-[10px] text-gray-400 mt-1">
                    Suleste, SP+CO, NENO, National One, Sul+CO, SP Est, Leste+KA
                  </p>
                </div>

                <div
                  onClick={() =>
                    setFiltroSegmentoLeve(
                      filtroSegmentoLeve === 'DTS LEVE' ? 'TODOS' : 'DTS LEVE'
                    )
                  }
                  className={`bg-[#1C2541] border p-5 rounded-2xl shadow-md cursor-pointer transition-all ${
                    filtroSegmentoLeve === 'DTS LEVE'
                      ? 'border-emerald-500 bg-emerald-500/15 scale-[1.02]'
                      : 'border-gray-800 hover:border-emerald-500/50'
                  }`}
                >
                  <div className="flex justify-between items-center">
                    <span className="text-xs font-black text-emerald-400 uppercase">
                      🚙 TOTAL DTS LEVE
                    </span>
                    <span className="text-xs font-mono font-black text-white">
                      {baseLevesSemFiltroSegmento.length > 0
                        ? (
                            (resumoSegmentosLeves['DTS LEVE'] /
                              baseLevesSemFiltroSegmento.length) *
                            100
                          ).toFixed(1)
                        : 0}
                      %
                    </span>
                  </div>
                  <h2 className="text-3xl font-black text-emerald-400 mt-2">
                    {resumoSegmentosLeves['DTS LEVE']}
                  </h2>
                  <p className="text-[10px] text-gray-400 mt-1">
                    Hyper Transformation & Execution + Zonas DTS + Fiorinos
                  </p>
                </div>

                <div
                  onClick={() =>
                    setFiltroSegmentoLeve(
                      filtroSegmentoLeve === 'SUPORTE / OUTROS'
                        ? 'TODOS'
                        : 'SUPORTE / OUTROS'
                    )
                  }
                  className={`bg-[#1C2541] border p-5 rounded-2xl shadow-md cursor-pointer transition-all ${
                    filtroSegmentoLeve === 'SUPORTE / OUTROS'
                      ? 'border-amber-500 bg-amber-500/15 scale-[1.02]'
                      : 'border-gray-800 hover:border-amber-500/50'
                  }`}
                >
                  <div className="flex justify-between items-center">
                    <span className="text-xs font-black text-amber-400 uppercase">
                      🛠️ SUPORTE / OUTRAS ÁREAS
                    </span>
                    <span className="text-xs font-mono font-black text-white">
                      {baseLevesSemFiltroSegmento.length > 0
                        ? (
                            (resumoSegmentosLeves['SUPORTE / OUTROS'] /
                              baseLevesSemFiltroSegmento.length) *
                            100
                          ).toFixed(1)
                        : 0}
                      %
                    </span>
                  </div>
                  <h2 className="text-3xl font-black text-amber-400 mt-2">
                    {resumoSegmentosLeves['SUPORTE / OUTROS']}
                  </h2>
                  <p className="text-[10px] text-gray-400 mt-1">
                    EHS (15), Agro (13),{' '}
                    {incluirTranspWhsNoDts
                      ? 'Security (3), Outros (4)'
                      : 'Transp (12), Warehouse (6), Security (3)'}
                  </p>
                </div>
              </div>

              {/* 📊 GRÁFICOS COMPARATIVOS: REGIONAL (ÁREA) E GERENTE */}
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-8">
                {/* Distribuição por Área / Regional */}
                <div className="bg-[#1C2541] border border-gray-800 p-6 rounded-2xl shadow-xl">
                  <div className="flex justify-between items-center mb-5">
                    <div>
                      <h3 className="text-sm font-black text-white uppercase flex items-center gap-2">
                        <MapPin size={16} className="text-emerald-400" />
                        Divisão por Regional / Área (DTS Leve vs. OT)
                      </h3>
                      <p className="text-xs text-gray-400 mt-0.5">
                        Clique em uma área para filtrar os veículos
                      </p>
                    </div>
                    <div className="flex items-center gap-3 text-[10px] font-bold">
                      <span className="flex items-center gap-1 text-cyan-400">
                        <span className="w-2.5 h-2.5 bg-cyan-400 rounded-full"></span>{' '}
                        OT
                      </span>
                      <span className="flex items-center gap-1 text-emerald-400">
                        <span className="w-2.5 h-2.5 bg-emerald-400 rounded-full"></span>{' '}
                        DTS Leve
                      </span>
                      <span className="flex items-center gap-1 text-amber-400">
                        <span className="w-2.5 h-2.5 bg-amber-400 rounded-full"></span>{' '}
                        Suporte
                      </span>
                    </div>
                  </div>

                  <div className="space-y-3.5">
                    {comparativoRegionalLeves.map((item) => {
                      const maxTotal = comparativoRegionalLeves[0]?.total || 1;
                      const pOt = (item.ot / maxTotal) * 100;
                      const pDts = (item.dtsLeve / maxTotal) * 100;
                      const pOut = (item.outros / maxTotal) * 100;

                      return (
                        <div
                          key={item.area}
                          onClick={() =>
                            setFiltroAreaLeve(
                              filtroAreaLeve === item.area ? 'TODAS' : item.area
                            )
                          }
                          className={`p-2.5 rounded-xl cursor-pointer transition-all border ${
                            filtroAreaLeve === item.area
                              ? 'bg-[#0B1329] border-emerald-500'
                              : 'border-transparent hover:bg-[#0B1329]/60'
                          }`}
                        >
                          <div className="flex justify-between items-center text-xs font-bold mb-1.5">
                            <span className="text-white font-black">
                              {item.area}
                            </span>
                            <div className="flex items-center gap-3 font-mono text-[11px]">
                              <span className="text-cyan-400">
                                OT: <strong>{item.ot}</strong>
                              </span>
                              <span className="text-emerald-400">
                                DTS Leve: <strong>{item.dtsLeve}</strong>
                              </span>
                              {item.outros > 0 && (
                                <span className="text-amber-400">
                                  Sup: <strong>{item.outros}</strong>
                                </span>
                              )}
                              <span className="text-gray-300 bg-gray-800 px-2 py-0.5 rounded">
                                Total: {item.total}
                              </span>
                            </div>
                          </div>
                          <div className="w-full h-3 bg-gray-800 rounded-full overflow-hidden flex">
                            <div
                              style={{ width: `${pOt}%` }}
                              className="h-full bg-cyan-500 transition-all duration-500"
                              title={`OT: ${item.ot}`}
                            ></div>
                            <div
                              style={{ width: `${pDts}%` }}
                              className="h-full bg-emerald-500 transition-all duration-500"
                              title={`DTS Leve: ${item.dtsLeve}`}
                            ></div>
                            <div
                              style={{ width: `${pOut}%` }}
                              className="h-full bg-amber-500 transition-all duration-500"
                              title={`Suporte/Outros: ${item.outros}`}
                            ></div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>

                {/* Distribuição por Gerente Regional */}
                <div className="bg-[#1C2541] border border-gray-800 p-6 rounded-2xl shadow-xl">
                  <div className="flex justify-between items-center mb-5">
                    <div>
                      <h3 className="text-sm font-black text-white uppercase flex items-center gap-2">
                        <Users size={16} className="text-cyan-400" />
                        Frota Leve por Gerente Regional
                      </h3>
                      <p className="text-xs text-gray-400 mt-0.5">
                        Clique no gerente para filtrar sua carteira de veículos
                      </p>
                    </div>
                  </div>

                  <div className="space-y-3.5">
                    {comparativoGerentesLeves.map((item) => {
                      const maxTotal = comparativoGerentesLeves[0]?.total || 1;
                      const pOt = (item.ot / maxTotal) * 100;
                      const pDts = (item.dtsLeve / maxTotal) * 100;
                      const pOut = (item.outros / maxTotal) * 100;

                      return (
                        <div
                          key={item.gerente}
                          onClick={() =>
                            setFiltroGerenteLeve(
                              filtroGerenteLeve === item.gerente
                                ? 'TODOS'
                                : item.gerente
                            )
                          }
                          className={`p-2.5 rounded-xl cursor-pointer transition-all border ${
                            filtroGerenteLeve === item.gerente
                              ? 'bg-[#0B1329] border-cyan-500'
                              : 'border-transparent hover:bg-[#0B1329]/60'
                          }`}
                        >
                          <div className="flex justify-between items-center text-xs font-bold mb-1.5">
                            <span className="text-white font-black truncate pr-2">
                              {item.gerente}
                            </span>
                            <div className="flex items-center gap-3 font-mono text-[11px] flex-shrink-0">
                              <span className="text-cyan-400">
                                OT: <strong>{item.ot}</strong>
                              </span>
                              <span className="text-emerald-400">
                                DTS: <strong>{item.dtsLeve}</strong>
                              </span>
                              <span className="text-gray-300 bg-gray-800 px-2 py-0.5 rounded">
                                {item.total} carros
                              </span>
                            </div>
                          </div>
                          <div className="w-full h-3 bg-gray-800 rounded-full overflow-hidden flex">
                            <div
                              style={{ width: `${pOt}%` }}
                              className="h-full bg-cyan-500 transition-all duration-500"
                            ></div>
                            <div
                              style={{ width: `${pDts}%` }}
                              className="h-full bg-emerald-500 transition-all duration-500"
                            ></div>
                            <div
                              style={{ width: `${pOut}%` }}
                              className="h-full bg-amber-500 transition-all duration-500"
                            ></div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              </div>

              {/* 📋 TABELA COMPLETA HIGIENIZADA DE VEÍCULOS LEVES ATIVOS */}
              <div className="bg-[#1C2541] border border-gray-800 rounded-2xl overflow-hidden shadow-xl">
                <div className="p-4 border-b border-gray-800 flex flex-col md:flex-row justify-between items-start md:items-center gap-2">
                  <div>
                    <h3 className="text-sm font-black text-white uppercase">
                      Base Higienizada de Veículos Leves Ativos (
                      {levesFiltradosFinal.length} veículos exibidos)
                    </h3>
                    <p className="text-xs text-gray-400 mt-0.5">
                      Classificação automática aplicada sobre Diretoria + Centro
                      de Custo SAP • Pronta para exportar
                    </p>
                  </div>
                </div>

                <div className="overflow-x-auto max-h-[550px]">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-[#0B1329] text-gray-400 font-bold uppercase border-b border-gray-800 sticky top-0 z-10">
                      <tr>
                        <th className="p-3">Segmento Oficial</th>
                        <th className="p-3">Código</th>
                        <th className="p-3">Placa</th>
                        <th className="p-3">Modelo</th>
                        <th className="p-3">Tipo Frota</th>
                        <th className="p-3">Diretoria (Pos. Nacional)</th>
                        <th className="p-3">Descrição Centro Custo</th>
                        <th className="p-3">CDV / Unidade</th>
                        <th className="p-3">Área</th>
                        <th className="p-3">Gerente</th>
                        <th className="p-3">Condutor / Usuário</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-800 text-gray-200 font-medium">
                      {levesFiltradosFinal.slice(0, 200).map((v) => (
                        <tr
                          key={v.id}
                          className="hover:bg-[#0B1329]/50 transition-all"
                        >
                          <td className="p-3 whitespace-nowrap">
                            <span
                              className={`px-2.5 py-1 rounded-full text-[10px] font-black border ${
                                v.segmentoOficial === 'OT'
                                  ? 'bg-cyan-500/20 text-cyan-400 border-cyan-500/30'
                                  : v.segmentoOficial === 'DTS LEVE'
                                  ? 'bg-emerald-500/20 text-emerald-400 border-emerald-500/30'
                                  : 'bg-amber-500/20 text-amber-400 border-amber-500/30'
                              }`}
                            >
                              {v.segmentoOficial}
                            </span>
                          </td>
                          <td className="p-3 font-mono text-gray-400 font-bold">
                            {v.codigo}
                          </td>
                          <td className="p-3 font-mono font-black text-white">
                            {v.placa}
                          </td>
                          <td className="p-3 text-gray-200 whitespace-nowrap">
                            {v.modelo}
                            {v.alertaErroCadastro && (
                              <span
                                className="ml-1.5 text-[10px] bg-red-500/20 text-red-400 px-1.5 py-0.5 rounded border border-red-500/30"
                                title={v.alertaErroCadastro}
                              >
                                ⚠️ Corrigido
                              </span>
                            )}
                          </td>
                          <td className="p-3 text-gray-300 whitespace-nowrap">
                            {v.frota}
                          </td>
                          <td className="p-3 text-gray-300 whitespace-nowrap">
                            {v.diretoriaOriginal}
                          </td>
                          <td className="p-3 font-bold text-amber-400 whitespace-nowrap">
                            {v.descCcusto}
                          </td>
                          <td className="p-3 text-gray-300 whitespace-nowrap">
                            {v.cdv}
                          </td>
                          <td className="p-3 font-black text-white">
                            {v.area}
                          </td>
                          <td className="p-3 text-gray-300 whitespace-nowrap">
                            {v.gerente}
                          </td>
                          <td className="p-3 text-gray-400 truncate max-w-xs">
                            {v.usuario}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </>
          )}
        </div>
      )}
    </div>
  );
}
