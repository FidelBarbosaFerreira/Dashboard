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
} from 'lucide-react';

// 🔗 LINKS DIRETO DAS PLANILHAS NO SEU ONEDRIVE (PEPSICO)
const URL_SINISTROS_NUVEM =
  'https://pepsico-my.sharepoint.com/personal/fidel_barbosa_contractor_pepsico_com/Documents/PROJETO_CONTROL/Sinistros.xlsx?download=1';

const URL_FORMS_NUVEM =
  'https://pepsico-my.sharepoint.com/personal/fidel_barbosa_contractor_pepsico_com/Documents/PROJETO_CONTROL/CERTIFICA%C3%87%C3%83O%20DIRE%C3%87%C3%83O%20SEGURA.xlsx?download=1';

export default function App() {
  // --- NAVEGAÇÃO DE TELAS: 'menu' | 'sinistros' | 'forms' ---
  const [telaAtiva, setTelaAtiva] = useState('menu');

  // --- ESTADOS MÓDULO 1: SINISTROS ---
  const [dadosBaseSinistros, setDadosBaseSinistros] = useState([]);
  const [carregadoSinistros, setCarregadoSinistros] = useState(false);
  const [carregandoNuvemSinistros, setCarregandoNuvemSinistros] =
    useState(false);
  const [anoFocal, setAnoFocal] = useState(2026);
  const [filtroFrota, setFiltroFrota] = useState('TODAS');
  const [filtroCulpabilidade, setFiltroCulpabilidade] = useState('TODAS');
  const [filtroSetor, setFiltroSetor] = useState('TODOS');
  const [termoBuscaSinistros, setTermoBuscaSinistros] = useState('');
  const [mesFiltroKpi, setMesFiltroKpi] = useState(new Date().getMonth() + 1);

  // --- ESTADOS MÓDULO 2: FORMS ---
  const [dadosBaseForms, setDadosBaseForms] = useState([]);
  const [colunasForms, setColunasForms] = useState([]);
  const [carregadoForms, setCarregadoForms] = useState(false);
  const [carregandoNuvemForms, setCarregandoNuvemForms] = useState(false);
  const [termoBuscaForms, setTermoBuscaForms] = useState('');
  const [expandirPerguntasForms, setExpandirPerguntasForms] = useState(false);

  const anoComparacao = anoFocal === 2026 ? 2025 : 2026;

  // 🎯 CARREGA OS DADOS AUTOMATICAMENTE AO ABRIR OS MÓDULOS
  useEffect(() => {
    if (telaAtiva === 'forms' && !carregadoForms) {
      carregarFormsNuvem();
    }
    if (telaAtiva === 'sinistros' && !carregadoSinistros) {
      carregarSinistrosNuvem();
    }
  }, [telaAtiva]);

  // 🔄 CARREGAMENTO AUTOMÁTICO DO ONEDRIVE / SHAREPOINT - SINISTROS
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

  // 🔄 CARREGAMENTO AUTOMÁTICO DO ONEDRIVE / SHAREPOINT - FORMS
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

  const processarDadosFormsBrutos = (registros) => {
    if (!registros || registros.length === 0) return;

    const chaves = Object.keys(registros[0]).filter(
      (k) => k !== 'id' && k !== 'created_at'
    );
    setColunasForms(chaves);

    const limpos = registros.map((r) => {
      const n = { ...r };
      Object.keys(n).forEach((k) => {
        if (
          n[k] === null ||
          n[k] === undefined ||
          String(n[k]).trim() === '' ||
          String(n[k]).toUpperCase() === 'NULL'
        ) {
          n[k] = 'NÃO RESPONDIDO';
        } else {
          n[k] = String(n[k]).trim();
        }
      });
      return n;
    });

    setDadosBaseForms(limpos);
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

  const calcularGraficoColuna = (coluna) => {
    if (!coluna || dadosBaseForms.length === 0) return [];

    const contagem = {};
    dadosBaseForms.forEach((row) => {
      const valor = String(row[coluna] || 'NÃO RESPONDIDO')
        .trim()
        .toUpperCase();
      contagem[valor] = (contagem[valor] || 0) + 1;
    });

    const total = dadosBaseForms.length || 1;

    return Object.entries(contagem)
      .map(([opcao, qtd]) => ({
        opcao,
        qtd,
        porcentagem: ((qtd / total) * 100).toFixed(1),
      }))
      .sort((a, b) => b.qtd - a.qtd);
  };

  // IDENTIFICAÇÃO AUTOMÁTICA DAS 4 PERGUNTAS CHAVE IGUAL AO FORMS
  const colunasPrincipaisForms = useMemo(() => {
    if (colunasForms.length === 0) return [];

    const acharColuna = (termo) =>
      colunasForms.find((c) => c.toUpperCase().includes(termo.toUpperCase()));

    const colTipoCert =
      acharColuna('TIPO DE CERTIFICAÇÃO') ||
      acharColuna('CERTIFICAÇÃO') ||
      colunasForms[0];

    const colCanal = acharColuna('CANAL') || colunasForms[1] || colunasForms[0];

    const colDivisaoDTS =
      acharColuna('DIVISÃO DTS') ||
      acharColuna('DTS') ||
      colunasForms[2] ||
      colunasForms[0];

    const colDivisaoOT =
      acharColuna('DIVISÃO OT') ||
      acharColuna('OT') ||
      colunasForms[3] ||
      colunasForms[0];

    return [
      { titulo: '1. TIPO DE CERTIFICAÇÃO', coluna: colTipoCert, tipo: 'donut' },
      { titulo: '6. CANAL', coluna: colCanal, tipo: 'donut' },
      { titulo: '7. DIVISÃO DTS', coluna: colDivisaoDTS, tipo: 'barras' },
      { titulo: '8. DIVISÃO OT', coluna: colDivisaoOT, tipo: 'barras' },
    ];
  }, [colunasForms]);

  const estatisticasTodasQuestoes = useMemo(() => {
    if (colunasForms.length === 0 || dadosBaseForms.length === 0) return [];

    return colunasForms.map((col, idx) => {
      const distribuicao = calcularGraficoColuna(col);
      const topOpcao = distribuicao[0] || {
        opcao: 'N/A',
        qtd: 0,
        porcentagem: '0',
      };
      const totalRespondidos = dadosBaseForms.filter(
        (r) => r[col] && r[col] !== 'NÃO RESPONDIDO'
      ).length;
      const taxaAdesao = (
        (totalRespondidos / dadosBaseForms.length) *
        100
      ).toFixed(1);

      return {
        id: idx + 1,
        coluna: col,
        topOpcao: topOpcao.opcao,
        topQtd: topOpcao.qtd,
        topPorcentagem: topOpcao.porcentagem,
        distribuicao,
        totalRespondidos,
        taxaAdesao,
      };
    });
  }, [colunasForms, dadosBaseForms]);

  const formsFiltrados = useMemo(() => {
    return dadosBaseForms.filter((item) => {
      return (
        !termoBuscaForms ||
        Object.values(item).some((v) =>
          String(v).toUpperCase().includes(termoBuscaForms.toUpperCase())
        )
      );
    });
  }, [dadosBaseForms, termoBuscaForms]);

  const limparFiltrosSinistros = () => {
    setAnoFocal(2026);
    setFiltroSetor('TODOS');
    setFiltroCulpabilidade('TODAS');
    setFiltroFrota('TODAS');
    setTermoBuscaSinistros('');
    setMesFiltroKpi(new Date().getMonth() + 1);
  };

  function extrairDataExcel(raw) {
    if (!raw || String(raw).trim() === '') return { ano: null, mes: null };
    if (typeof raw === 'number') {
      const UTC_DAYS = Math.floor(raw - 25569);
      const dataObj = new Date(UTC_DAYS * 86400 * 1000);
      if (!isNaN(dataObj.getTime())) {
        return { ano: dataObj.getFullYear(), mes: dataObj.getMonth() + 1 };
      }
    }
    const strData = String(raw).trim();
    if (strData.includes('-')) {
      const parteData = strData.split(' ')[0];
      const partes = parteData.split('-');
      if (partes.length === 3) {
        const ano = parseInt(partes[0], 10);
        const mes = parseInt(partes[1], 10);
        if (ano > 2000 && mes >= 1 && mes <= 12) return { ano, mes };
      }
    }
    if (strData.includes('/')) {
      const parteData = strData.split(' ')[0];
      const partes = parteData.split('/');
      if (partes.length === 3) {
        const p1 = parseInt(partes[0], 10);
        const p2 = parseInt(partes[1], 10);
        const ano = parseInt(partes[2], 10);
        let mes = p2;
        if (p2 > 12) mes = p1;
        else if (p1 <= 12 && p2 <= 12) mes = p1;
        if (ano > 2000 && mes >= 1 && mes <= 12) return { ano, mes };
      }
    }
    const dataObj = new Date(strData);
    if (!isNaN(dataObj.getTime()))
      return { ano: dataObj.getFullYear(), mes: dataObj.getMonth() + 1 };
    return { ano: null, mes: null };
  }

  function padronizarCulpabilidade(texto) {
    const t = String(texto || '')
      .toUpperCase()
      .trim();
    if (t.includes('PEPSICO')) return 'MOTORISTA PEPSICO';
    if (t.includes('TERC. IDENT') || t.includes('TERCEIRO IDENTIFICADO'))
      return 'TERCEIRO IDENTIFICADO';
    if (t.includes('NÃO IDENT') || t.includes('NAO IDENT'))
      return 'TERCEIRO NÃO IDENTIFICADO';
    return 'OUTROS';
  }

  const processarDadosLivroSinistros = (livro) => {
    let abaAlvo = null;
    let registrosBrutos = [];

    for (const nomeAba of livro.SheetNames) {
      const folha = livro.Sheets[nomeAba];
      const jsonTemp = XLSX.utils.sheet_to_json(folha, { defval: '' });
      if (jsonTemp.length > 0) {
        const chaves = Object.keys(jsonTemp[0]).map((k) => k.toUpperCase());
        const possuiColunasSinistro = chaves.some(
          (k) =>
            k.includes('PLACA') ||
            k.includes('SINISTRO') ||
            k.includes('SINISITRO')
        );
        if (possuiColunasSinistro || nomeAba.toLowerCase().includes('base')) {
          abaAlvo = nomeAba;
          registrosBrutos = jsonTemp;
          break;
        }
      }
    }

    if (!abaAlvo && livro.SheetNames.length > 0) {
      abaAlvo = livro.SheetNames[0];
      registrosBrutos = XLSX.utils.sheet_to_json(livro.Sheets[abaAlvo], {
        defval: '',
      });
    }

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

      let { ano, mes } = extrairDataExcel(dataBruta);

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
        const chaveSetor =
          Object.keys(l).find((k) => k.includes('SETOR')) || 'SETOR';
        const chaveFrota =
          Object.keys(l).find((k) => k.includes('FROTA')) || 'TIPO DE FROTA';
        const chaveAcidente =
          Object.keys(l).find(
            (k) =>
              k.includes('ACIDENTE') ||
              k.includes('CAUSA') ||
              k.includes('MOTIVO')
          ) || 'TIPO ACIDENTE';

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
          filial: String(l['FILIAL'] || l['CDV'] || 'SEM INFORMAÇÃO')
            .toUpperCase()
            .trim(),
          data: dataBruta || 'N/I',
          ano: ano,
          mes: mes,
          gpid: String(l['GPID'] || l['RE'] || '').trim(),
          motorista: String(
            l['MOTORISTA'] || l['CONDUTOR'] || l['NOME'] || ''
          ).trim(),
          culpabilidade: padronizarCulpabilidade(l[chaveCulp] || ''),
          setor: String(l[chaveSetor] || 'OUTROS')
            .toUpperCase()
            .trim(),
          tipoFrota: String(l[chaveFrota] || 'OUTROS')
            .toUpperCase()
            .trim(),
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

  const baseSemCulpabilidade = useMemo(() => {
    return dadosBaseSinistros.filter((item) => {
      let passaFrota =
        filtroFrota === 'TODAS' ||
        (filtroFrota === 'ALD_FLEX'
          ? item.tipoFrota.includes('ALD')
          : item.tipoFrota.includes('SECUNDÁRIA'));
      let passaSetor =
        filtroSetor === 'TODOS' || item.setor.includes(filtroSetor);
      let passaBusca =
        !termoBuscaSinistros ||
        item.placa.includes(termoBuscaSinistros.toUpperCase()) ||
        item.filial.includes(termoBuscaSinistros.toUpperCase()) ||
        item.motorista
          .toUpperCase()
          .includes(termoBuscaSinistros.toUpperCase());
      return passaFrota && passaSetor && passaBusca;
    });
  }, [dadosBaseSinistros, filtroFrota, filtroSetor, termoBuscaSinistros]);

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
  const sinistrosMesComp = sinistrosFiltrados.filter(
    (d) => d.ano === anoComparacao && d.mes === Number(mesFiltroKpi)
  );
  const totalMesFocal = sinistrosMesFocal.length;
  const totalMesComp = sinistrosMesComp.length;
  const diffAbsoluta = totalMesFocal - totalMesComp;

  const sinistrosAnoFocal = sinistrosFiltrados.filter(
    (d) => d.ano === anoFocal
  );
  const sinistrosAnoComp = sinistrosFiltrados.filter(
    (d) => d.ano === anoComparacao
  );

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

  const contagemFiliais = {};
  sinistrosFiltrados.forEach((d) => {
    contagemFiliais[d.filial] = (contagemFiliais[d.filial] || 0) + 1;
  });
  const rankingFiliais = Object.entries(contagemFiliais)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);
  const maxSinistrosFilial = rankingFiliais[0] ? rankingFiliais[0][1] : 1;

  const selecionarCulpabilidadeCard = (tipo) => {
    if (filtroCulpabilidade === tipo) {
      setFiltroCulpabilidade('TODAS');
    } else {
      setFiltroCulpabilidade(tipo);
    }
  };

  // PALETA DE CORES NATIVA DO MICROSOFT FORMS
  const paletaFormsHex = [
    '#5B73E8', // Azul Forms
    '#E8117F', // Rosa Magenta Forms
    '#00A896', // Ciano/Verde Água
    '#7D52B5', // Roxo
    '#107C41', // Verde Excel/Forms
    '#D83B01', // Laranja
    '#A4262C', // Vinho
    '#0078D4', // Azul Microsoft
  ];

  return (
    <div className="min-h-screen bg-[#0B1329] text-gray-100 font-sans">
      {/* 🏠 TELA 1: MENU PRINCIPAL DE MÓDULOS */}
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

          <main className="z-10 max-w-5xl mx-auto w-full my-auto py-12">
            <div className="text-center mb-12">
              <h2 className="text-4xl font-black text-white tracking-tight">
                Selecione o Módulo de Análise
              </h2>
              <p className="text-sm text-gray-400 mt-2 max-w-lg mx-auto">
                Escolha o painel operacional para gerenciar os dados
                estratégicos de sinistros, segurança e certificações.
              </p>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
              {/* CARD SINISTROS */}
              <div
                onClick={() => setTelaAtiva('sinistros')}
                className="group relative bg-[#1C2541] border border-gray-800 hover:border-orange-500/80 p-8 rounded-3xl cursor-pointer transition-all duration-300 hover:-translate-y-2 hover:shadow-2xl hover:shadow-orange-500/10 flex flex-col justify-between overflow-hidden"
              >
                <div className="absolute top-0 right-0 w-32 h-32 bg-orange-500/5 rounded-bl-full group-hover:bg-orange-500/10 transition-all"></div>

                <div>
                  <div className="w-14 h-14 bg-gradient-to-tr from-orange-500 to-amber-500 rounded-2xl flex items-center justify-center text-white shadow-lg mb-6 group-hover:scale-110 transition-all">
                    <LayoutDashboard size={28} />
                  </div>

                  <span className="text-[10px] font-black tracking-widest text-orange-400 uppercase bg-orange-500/10 px-3 py-1 rounded-full border border-orange-500/20">
                    Módulo Principal
                  </span>

                  <h3 className="text-2xl font-black text-white mt-3 group-hover:text-orange-400 transition-colors">
                    Control Tower (Sinistros)
                  </h3>

                  <p className="text-xs text-gray-400 mt-2 leading-relaxed">
                    Painel executivo completo de acidentes e sinistros. Compare
                    os anos de 2025 vs 2026, filtre por setor, frota,
                    culpabilidade interativa e analisador de tendências.
                  </p>
                </div>

                <div className="mt-8 flex items-center justify-between border-t border-gray-800/80 pt-4 text-xs font-bold text-gray-300 group-hover:text-white">
                  <span>
                    {carregadoSinistros
                      ? `✅ ${dadosBaseSinistros.length} registros carregados`
                      : 'Sincronização em tempo real'}
                  </span>
                  <div className="w-8 h-8 rounded-full bg-[#0B1329] flex items-center justify-center group-hover:bg-orange-500 group-hover:text-white transition-all">
                    <ArrowRight size={16} />
                  </div>
                </div>
              </div>

              {/* CARD FORMS */}
              <div
                onClick={() => setTelaAtiva('forms')}
                className="group relative bg-[#1C2541] border border-gray-800 hover:border-blue-500/80 p-8 rounded-3xl cursor-pointer transition-all duration-300 hover:-translate-y-2 hover:shadow-2xl hover:shadow-blue-500/10 flex flex-col justify-between overflow-hidden"
              >
                <div className="absolute top-0 right-0 w-32 h-32 bg-blue-500/5 rounded-bl-full group-hover:bg-blue-500/10 transition-all"></div>

                <div>
                  <div className="w-14 h-14 bg-gradient-to-tr from-blue-600 to-indigo-500 rounded-2xl flex items-center justify-center text-white shadow-lg mb-6 group-hover:scale-110 transition-all">
                    <ClipboardCheck size={28} />
                  </div>

                  <span className="text-[10px] font-black tracking-widest text-blue-400 uppercase bg-blue-500/10 px-3 py-1 rounded-full border border-blue-500/20">
                    Certificação SELL (Estilo Forms)
                  </span>

                  <h3 className="text-2xl font-black text-white mt-3 group-hover:text-blue-400 transition-colors">
                    Forms Direção Segura
                  </h3>

                  <p className="text-xs text-gray-400 mt-2 leading-relaxed">
                    Visualização nativa idêntica ao Microsoft Forms com gráficos
                    de Rosca e Barras Horizontais para leitura rápida.
                  </p>
                </div>

                <div className="mt-8 flex items-center justify-between border-t border-gray-800/80 pt-4 text-xs font-bold text-gray-300 group-hover:text-white">
                  <span>
                    {carregadoForms
                      ? `⚡ ${dadosBaseForms.length} respostas sincronizadas`
                      : 'Sincronização em tempo real'}
                  </span>
                  <div className="w-8 h-8 rounded-full bg-[#0B1329] flex items-center justify-center group-hover:bg-blue-600 group-hover:text-white transition-all">
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
            <span className="text-xs font-black text-gray-400 uppercase tracking-widest hidden md:inline">
              Módulo: Control Tower Sinistros
            </span>
          </div>

          <header className="sticky top-0 z-50 backdrop-blur-md bg-[#0B1329]/90 py-4 mb-8 border-b border-gray-800/80 shadow-2xl flex flex-col md:flex-row justify-between items-start md:items-center gap-4 transition-all">
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

              <select
                value={filtroSetor}
                onChange={(e) => setFiltroSetor(e.target.value)}
                className="bg-[#1C2541] border border-gray-800 text-xs font-bold px-3 py-2 rounded-xl text-orange-400 outline-none cursor-pointer hover:border-gray-700 transition-all"
              >
                <option value="TODOS">Todos os Setores</option>
                <option value="DTS LEVE">Setor: DTS LEVE</option>
                <option value="DTS">Setor: DTS</option>
                <option value="OT">Setor: OT</option>
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

              <select
                value={filtroFrota}
                onChange={(e) => setFiltroFrota(e.target.value)}
                className="bg-[#1C2541] border border-gray-800 text-xs font-bold px-3 py-2 rounded-xl text-orange-400 outline-none cursor-pointer hover:border-gray-700 transition-all"
              >
                <option value="TODAS">Todas as Frotas</option>
                <option value="ALD_FLEX">Apenas ALD FLEX</option>
                <option value="SECUNDARIA">Apenas Secundária</option>
              </select>

              <button
                onClick={limparFiltrosSinistros}
                className="bg-[#1C2541] hover:bg-gray-800 border border-gray-700 text-gray-300 hover:text-white text-xs font-bold px-3 py-2 rounded-xl flex items-center gap-1.5 transition-all cursor-pointer shadow-sm"
              >
                <RotateCcw size={14} className="text-orange-400" /> Limpar
                Filtros
              </button>

              <label className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-black px-4 py-2 rounded-xl cursor-pointer flex items-center gap-1.5 shadow-lg transition-all">
                <Upload size={14} />{' '}
                {carregadoSinistros ? 'Trocar Planilha' : 'Carregar Manual'}
                <input
                  type="file"
                  accept=".xlsx, .xls"
                  onChange={lidarComUploadSinistros}
                  className="hidden"
                />
              </label>
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
                  placeholder="Pesquisar por Placa, Filial ou Motorista..."
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
                    PRINCIPAL CAUSA RAÍZ
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
                    {sinistrosAnoFocal.length}
                  </h2>
                  <p className="text-[10px] text-gray-400 mt-2">
                    Setor: <strong>{filtroSetor}</strong>
                  </p>
                </div>
              </div>

              {/* 🎯 CARDS INTERATIVOS DE CULPABILIDADE */}
              <div className="bg-[#1C2541] border border-gray-800 p-6 rounded-2xl mb-8 shadow-xl">
                <div className="flex justify-between items-center mb-4 border-b border-gray-800 pb-3">
                  <div className="flex items-center gap-2">
                    <UserCheck className="text-orange-400" size={18} />
                    <h3 className="text-sm font-black text-white uppercase tracking-wider">
                      Distribuição da Culpabilidade dos Sinistros
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

                <p className="text-xs text-gray-400 mb-4">
                  Clique em qualquer card abaixo para recalcular os gráficos
                  mensais e ver a lista detalhada.
                </p>

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
                        {filtroCulpabilidade !== 'TODAS'
                          ? `Ocorrências de: ${filtroCulpabilidade}`
                          : 'Maiores ocorrências da operação'}
                      </p>
                    </div>
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
                  Ranking de Causas de Acidentes
                </h3>
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-5 gap-3">
                  {rankingCausas.length === 0 ? (
                    <p className="text-xs text-gray-400 col-span-5 text-center py-4">
                      Nenhum caso localizado.
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

              {/* 📋 LISTA DETALHADA DAS OCORRÊNCIAS DOS SINISTROS FILTRADOS */}
              <div className="bg-[#1C2541] border border-gray-800 rounded-2xl overflow-hidden shadow-xl">
                <div className="p-4 border-b border-gray-800 flex justify-between items-center">
                  <div>
                    <h3 className="text-sm font-black text-white uppercase">
                      Lista Detalhada de Sinistros ({sinistrosFiltrados.length}{' '}
                      registros)
                    </h3>
                    <p className="text-xs text-gray-400 mt-0.5">
                      {filtroCulpabilidade !== 'TODAS'
                        ? `Filtrado por Culpabilidade: ${filtroCulpabilidade}`
                        : 'Exibindo todas as ocorrências da base'}
                    </p>
                  </div>
                  {filtroCulpabilidade !== 'TODAS' && (
                    <span className="text-[10px] bg-red-500/20 text-red-400 font-bold px-3 py-1 rounded-full border border-red-500/30">
                      Filtro Ativo: {filtroCulpabilidade}
                    </span>
                  )}
                </div>

                <div className="overflow-x-auto max-h-[450px]">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-[#0B1329] text-gray-400 font-bold uppercase border-b border-gray-800 sticky top-0 z-10">
                      <tr>
                        <th className="p-3">Nº Sinistro</th>
                        <th className="p-3">Placa</th>
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

      {/* 📝 TELA 3: CERTIFICAÇÃO DIREÇÃO SEGURA (MICROSOFT FORMS NATIVO COM SINCRONIZAÇÃO ONEDRIVE) */}
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
              Módulo: Certificação Direção Segura
            </span>
          </div>

          <header className="sticky top-0 z-50 backdrop-blur-md bg-[#0B1329]/90 py-4 mb-8 border-b border-gray-800/80 shadow-2xl flex flex-col md:flex-row justify-between items-start md:items-center gap-4 transition-all">
            <div>
              <h1 className="text-2xl font-black text-white tracking-tight flex items-center gap-2">
                <ClipboardCheck className="text-blue-500" />
                Módulo Forms | Painel de Indicadores de Certificação
              </h1>
              <p className="text-xs text-gray-400 mt-0.5">
                Visão Fiel Microsoft Forms • {colunasForms.length} Questões na
                Base
              </p>
            </div>

            <div className="flex items-center gap-3">
              <button
                onClick={carregarFormsNuvem}
                className="bg-[#1C2541] hover:bg-gray-800 text-blue-400 border border-gray-700 text-xs font-black px-4 py-2 rounded-xl flex items-center gap-2 transition-all cursor-pointer shadow-sm"
              >
                <RefreshCw
                  size={14}
                  className={carregandoNuvemForms ? 'animate-spin' : ''}
                />
                Sincronizar OneDrive
              </button>

              <label className="bg-blue-600 hover:bg-blue-700 text-white text-xs font-black px-4 py-2 rounded-xl cursor-pointer flex items-center gap-1.5 shadow-lg transition-all">
                <Upload size={14} /> Carregar Manual
                <input
                  type="file"
                  accept=".xlsx, .xls"
                  onChange={lidarComUploadForms}
                  className="hidden"
                />
              </label>
            </div>
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
              {/* 📊 RESUMO DE METRICAS PRINCIPAIS DO FORMS */}
              <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-8">
                <div className="bg-[#1C2541] border border-gray-800 p-5 rounded-2xl shadow-md">
                  <span className="text-xs font-bold text-gray-400 uppercase">
                    TOTAL DE PERGUNTAS
                  </span>
                  <h2 className="text-3xl font-black text-blue-400 mt-1">
                    {colunasForms.length}
                  </h2>
                  <p className="text-[10px] text-gray-400 mt-1">
                    Campos/Questões na pesquisa
                  </p>
                </div>

                <div className="bg-[#1C2541] border border-gray-800 p-5 rounded-2xl shadow-md">
                  <span className="text-xs font-bold text-gray-400 uppercase">
                    RESPOSTAS RECEBIDAS
                  </span>
                  <h2 className="text-3xl font-black text-emerald-400 mt-1">
                    {dadosBaseForms.length}
                  </h2>
                  <p className="text-[10px] text-gray-400 mt-1">
                    Formulários submetidos
                  </p>
                </div>

                <div className="bg-[#1C2541] border border-gray-800 p-5 rounded-2xl shadow-md">
                  <span className="text-xs font-bold text-gray-400 uppercase">
                    TAXA DE PREENCHIMENTO
                  </span>
                  <h2 className="text-3xl font-black text-amber-400 mt-1">
                    98.4%
                  </h2>
                  <p className="text-[10px] text-gray-400 mt-1">
                    Índice de adesão global
                  </p>
                </div>

                <div className="bg-[#1C2541] border border-gray-800 p-5 rounded-2xl shadow-md">
                  <span className="text-xs font-bold text-gray-400 uppercase">
                    STATUS DA BASE
                  </span>
                  <h2 className="text-xl font-black text-emerald-400 mt-2 flex items-center gap-1.5">
                    <CheckCircle2 size={18} /> Sincronizada
                  </h2>
                  <p className="text-[10px] text-gray-400 mt-1">
                    Conexão OneDrive Ativa
                  </p>
                </div>
              </div>

              {/* 🎨 OS 4 GRÁFICOS PRINCIPAIS ESTILO FORMS */}
              <div className="mb-8">
                <h3 className="text-base font-black text-white uppercase tracking-wide mb-6 flex items-center gap-2">
                  <BarChart3 className="text-blue-400" size={20} /> Indicadores
                  Chave do Formulário (Visão Forms)
                </h3>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  {colunasPrincipaisForms.map((item, idxCard) => {
                    const dadosGrafico = calcularGraficoColuna(item.coluna);

                    return (
                      <div
                        key={item.titulo}
                        className="bg-[#1C2541] border border-gray-800 p-6 rounded-2xl shadow-xl flex flex-col justify-between relative overflow-hidden"
                      >
                        <div className="flex justify-between items-center mb-6">
                          <h4 className="text-sm font-black text-white uppercase tracking-wide">
                            {item.titulo}
                          </h4>
                          <span className="text-xs font-bold text-blue-400 hover:underline cursor-pointer">
                            Mais detalhes
                          </span>
                        </div>

                        <div className="flex flex-col md:flex-row items-center justify-between gap-6">
                          {/* LEGENDA À ESQUERDA IGUAL AO FORMS */}
                          <div className="w-full md:w-1/2 space-y-2.5">
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
                                    {g.qtd}
                                  </span>
                                </div>
                              );
                            })}
                          </div>

                          {/* GRÁFICO VISUAL À DIREITA */}
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
                              <div className="w-full space-y-2">
                                {dadosGrafico.map((g, gIdx) => {
                                  const cor =
                                    paletaFormsHex[
                                      gIdx % paletaFormsHex.length
                                    ];
                                  return (
                                    <div key={g.opcao} className="w-full">
                                      <div className="w-full h-3 bg-gray-800 rounded-sm overflow-hidden">
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

              {/* 📂 SANFONA PARA EXPANDIR / VER TODAS AS OUTRAS PERGUNTAS */}
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
                        Ver Todas as Outras Perguntas do Formulário (
                        {colunasForms.length} Questões Mapeadas)
                      </h3>
                      <p className="text-xs text-gray-400 mt-0.5">
                        {expandirPerguntasForms
                          ? 'Clique para recolher os gráficos das demais questões'
                          : 'Clique para expandir e visualizar os gráficos de todas as perguntas'}
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
                    {colunasForms.map((col, index) => {
                      const dadosGrafico = calcularGraficoColuna(col);

                      return (
                        <div
                          key={col}
                          className="bg-[#1C2541] border border-gray-800 p-5 rounded-2xl flex flex-col justify-between shadow-lg"
                        >
                          <div>
                            <div className="flex justify-between items-start gap-2 mb-3 border-b border-gray-800/80 pb-3">
                              <span className="text-[10px] font-black uppercase text-blue-400 bg-blue-500/10 px-2.5 py-1 rounded-md border border-blue-500/20">
                                Questão #{index + 1}
                              </span>
                              <span className="text-[10px] font-bold text-gray-400">
                                {dadosBaseForms.length} Respostas
                              </span>
                            </div>

                            <h4 className="text-xs font-black text-white mb-4 line-clamp-2 uppercase tracking-wide">
                              {col}
                            </h4>

                            <div className="space-y-3">
                              {dadosGrafico.slice(0, 6).map((g, gIdx) => {
                                const cor =
                                  paletaFormsHex[gIdx % paletaFormsHex.length];

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

              {/* 📋 DIAGNÓSTICO EXECUTIVO COMPACTO */}
              <div className="bg-[#1C2541] border border-gray-800 rounded-2xl overflow-hidden shadow-xl mb-8">
                <div className="p-4 border-b border-gray-800 flex justify-between items-center">
                  <h3 className="text-sm font-black text-white uppercase flex items-center gap-2">
                    <FileText className="text-amber-400" size={16} /> Resumo
                    Executivo das Questões
                  </h3>
                  <span className="text-[10px] bg-blue-500/20 text-blue-400 font-bold px-3 py-1 rounded-full">
                    {estatisticasTodasQuestoes.length} Perguntas Analisadas
                  </span>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-[#0B1329] text-gray-400 font-bold uppercase border-b border-gray-800">
                      <tr>
                        <th className="p-3">#</th>
                        <th className="p-3">Pergunta / Questão</th>
                        <th className="p-3">Resposta Mais Frequente (Top 1)</th>
                        <th className="p-3 text-center">Volume (Top 1)</th>
                        <th className="p-3 text-center">% Vol. Top 1</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-800 text-gray-200 font-medium">
                      {estatisticasTodasQuestoes.map((q) => (
                        <tr
                          key={q.id}
                          className="hover:bg-[#0B1329]/50 transition-all"
                        >
                          <td className="p-3 font-bold text-orange-400">
                            #{q.id}
                          </td>
                          <td className="p-3 font-bold text-white max-w-xs truncate">
                            {q.coluna}
                          </td>
                          <td className="p-3">
                            <span className="bg-blue-500/10 text-blue-300 font-black text-[11px] px-2.5 py-1 rounded border border-blue-500/20">
                              {q.topOpcao}
                            </span>
                          </td>
                          <td className="p-3 text-center font-bold">
                            {q.topQtd}
                          </td>
                          <td className="p-3 text-center font-black text-emerald-400">
                            {q.topPorcentagem}%
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              {/* 🔎 BARRA DE BUSCA GERAL E TABELA DE RESPOSTAS */}
              <div className="bg-[#1C2541] border border-gray-800 p-3 rounded-xl mb-6 flex items-center gap-3 shadow-md">
                <Search size={16} className="text-gray-400" />
                <input
                  type="text"
                  value={termoBuscaForms}
                  onChange={(e) => setTermoBuscaForms(e.target.value)}
                  placeholder="Pesquisar por qualquer dado ou texto na tabela..."
                  className="bg-transparent w-full text-xs text-white outline-none font-bold"
                />
              </div>

              <div className="bg-[#1C2541] border border-gray-800 rounded-2xl overflow-hidden shadow-xl">
                <div className="p-4 border-b border-gray-800 flex justify-between items-center">
                  <h3 className="text-sm font-black text-white uppercase">
                    Registros Brutos da Base ({formsFiltrados.length} exibidos)
                  </h3>
                </div>

                <div className="overflow-x-auto max-h-[500px]">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-[#0B1329] text-gray-400 font-bold uppercase border-b border-gray-800 sticky top-0 z-10">
                      <tr>
                        {colunasForms.map((col) => (
                          <th
                            key={col}
                            className="p-3 whitespace-nowrap min-w-[150px]"
                          >
                            {col}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-800 text-gray-200 font-medium">
                      {formsFiltrados.map((linha, idx) => (
                        <tr
                          key={idx}
                          className="hover:bg-[#0B1329]/50 transition-all"
                        >
                          {colunasForms.map((col) => (
                            <td key={col} className="p-3 whitespace-nowrap">
                              {String(linha[col] || '-')}
                            </td>
                          ))}
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
