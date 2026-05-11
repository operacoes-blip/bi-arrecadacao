import { 
 useEffect,
 useMemo,
 useState } from "react";
import { supabase } from "./lib/supabase";
import * as XLSX from "xlsx";

import {
  Building2,
 CreditCard,
 Landmark,
 Smartphone,
 Upload,
 Search,
 X,
 ChevronLeft,
 ChevronRight,
 ChevronUp,
 ChevronDown,
} from "lucide-react";

import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  ResponsiveContainer,
  LineChart,
  Line,
  CartesianGrid,
  PieChart,
  Pie,
  Cell,
  LabelList,
  Legend,
} from "recharts";

const coresPizza = [
  "#2563eb",
  "#22c55e",
  "#facc15",
  "#7c3aed",
  "#06b6d4",
  "#f97316",
  "#1e3a8a",
  "#db2777",
  "#0f766e",
  "#9ca3af",
];

function limparTexto(valor) {
  return String(valor || "").trim();
}

function converterValor(valor) {
  if (typeof valor === "number") return valor;

  return Number(
    String(valor || "0")
      .replace("R$", "")
      .replace(/\./g, "")
      .replace(",", ".")
      .trim()
  );
}

function converterData(valor) {
  if (!valor) return null;

  if (typeof valor === "number") {
    const data = XLSX.SSF.parse_date_code(valor);
    if (!data) return null;

    return `${data.y}-${String(data.m).padStart(2, "0")}-${String(
      data.d
    ).padStart(2, "0")}`;
  }

  let texto = String(valor)
    .trim()
    .replace(/["']/g, "")
    .replace(/\s.*/, "");

  if (!texto) return null;

  if (texto.includes("T")) texto = texto.substring(0, 10);

  const partes = texto.split(/[-/]/).map((p) => p.trim());

  if (partes.length !== 3) return null;

  let ano;
  let mes;
  let dia;

  if (partes[0].length === 4) {
    ano = partes[0];

    const parte2 = Number(partes[1]);
    const parte3 = Number(partes[2]);

    if (parte2 > 12 && parte3 <= 12) {
      dia = partes[1];
      mes = partes[2];
    } else {
      mes = partes[1];
      dia = partes[2];
    }
  } else {
    dia = partes[0];
    mes = partes[1];
    ano = partes[2];

    if (ano.length === 2) ano = `20${ano}`;
  }

  if (Number(mes) < 1 || Number(mes) > 12) return null;
  if (Number(dia) < 1 || Number(dia) > 31) return null;

  return `${ano}-${String(mes).padStart(2, "0")}-${String(dia).padStart(
    2,
    "0"
  )}`;
}

function formatarMoeda(valor) {
  return Number(valor || 0).toLocaleString("pt-BR", {
    style: "currency",
    currency: "BRL",
  });
}

function formatarNumero(valor) {
  return Number(valor || 0).toLocaleString("pt-BR");
}

function formatarDataBr(data) {
  if (!data) return "";

  const partes = String(data).split("-");

  if (partes.length !== 3) return data;

  return `${partes[2]}/${partes[1]}/${partes[0]}`;
}

function formatarDataCurta(data) {
  if (!data) return "";

  const partes = String(data).split("-");

  if (partes.length !== 3) return data;

  return `${partes[2]}/${partes[1]}`;
}

function normalizarTipoCarne(valor) {
  const texto = limparTexto(valor).toLowerCase();

  if (texto.includes("retorno")) return "CARNÊ FÍSICO";
  if (texto.includes("matheus")) return "CARNÊ DIGITAL";
  if (texto.includes("fisico")) return "CARNÊ FÍSICO";
  if (texto.includes("físico")) return "CARNÊ FÍSICO";
  if (texto.includes("digital")) return "CARNÊ DIGITAL";

  return limparTexto(valor) || "Não informado";
}

function normalizarParcela(valor) {
  const texto = limparTexto(valor).toUpperCase();

  if (!texto) return "Não informado";

  return texto;
}

function bairroValido(nome) {
  const texto = limparTexto(nome).toLowerCase();

  if (!texto) return false;
  if (texto === "não informado") return false;
  if (texto === "nao informado") return false;
  if (texto === "null") return false;
  if (texto === "undefined") return false;
  if (texto === "-") return false;
  if (texto === ".") return false;

  return true;
}

function abreviarBanco(nome) {
  return limparTexto(nome)
    .replace("BANCO DO BRASIL", "BCO DO BRASIL")
    .replace("BANCO BRADESCO", "BCO BRADESCO")
    .replace("BANCO ", "BCO ")
    .replace("CAIXA ECONOMICA FEDERAL", "CAIXA")
    .replace("CAIXA ECONÔMICA FEDERAL", "CAIXA")
    .replace("S.A.", "SA")
    .replace("(BRASIL) S.A.", "")
    .replace("UNIBANCO", "")
    .trim();
}

function abreviarTributo(nome) {
  const texto = limparTexto(nome);
  if (!texto) return "Não informado";
  // Abreviações para melhorar leitura no gráfico de Tributo (telão)
  return texto
    .replace(/ALVAR[ÁA] DE FUNCIONAMENTO/gi, "ALV. FUNC.")
    .replace(/TAXA DE LICENCIAMENTO/gi, "TAXA LIC.")
    .replace(/IMPOSTO SOBRE SERVI[ÇC]OS/gi, "ISS")
    .trim();
}


function agruparQuantidade(dados, campo, normalizador = null) {
  const mapa = {};

  dados.forEach((item) => {
    const chave = normalizador
      ? normalizador(item[campo])
      : item[campo] || "Não informado";

    if (!mapa[chave]) {
      mapa[chave] = {
        nome: chave,
        quantidade: 0,
        valor: 0,
      };
    }

    mapa[chave].quantidade += 1;
    mapa[chave].valor += Number(item.valor_pago || 0);
  });

  return Object.values(mapa);
}

function ordenarMaiorQuantidade(lista) {
  return [...lista].sort((a, b) => b.quantidade - a.quantidade);
}

async function carregarTodosRegistros(tabela, ordemCampo = null) {
  const tamanhoPagina = 1000;
  let inicio = 0;
  let todos = [];

  while (true) {
    let query = supabase
      .from(tabela)
      .select("*")
      .range(inicio, inicio + tamanhoPagina - 1);

    if (ordemCampo) {
      query = query.order(ordemCampo, { ascending: false });
    }

    const { data, error } = await query;

    if (error) {
      console.log(`Erro ao carregar ${tabela}:`, error);
      break;
    }

    if (!data || data.length === 0) break;

    todos = [...todos, ...data];

    if (data.length < tamanhoPagina) break;

    inicio += tamanhoPagina;
  }

  return todos;
}

async function limparTabela(tabela) {
  const { error } = await supabase.from(tabela).delete().neq("id", 0);

  if (error) {
    console.log(`Erro ao limpar ${tabela}:`, error);
    alert(`Erro ao limpar a tabela ${tabela}. Veja o console.`);
    return false;
  }

  return true;
}

function CardKpi({ titulo, valor, subtitulo, icon: Icone, cor }) {
  return (
    <div className="bg-white/90 backdrop-blur rounded-3xl shadow-sm hover:shadow-md transition p-6 border border-gray-200/70">
      <div className="flex items-center gap-4">
        <div
          className={`w-12 h-12 rounded-2xl flex items-center justify-center ${cor} shadow-sm ring-1 ring-white/50`}
        >
          <Icone size={26} className="text-white" />
        </div>

        <div>
          <p className="text-gray-600 text-sm xl:text-base font-semibold">{titulo}</p>
          <h2 className="text-2xl xl:text-3xl font-bold leading-tight">
            {valor}
          </h2>
          {subtitulo && <p className="text-gray-500 text-xs">{subtitulo}</p>}
        </div>
      </div>
    </div>
  );
}

function GraficoBarraVertical({
  titulo,
  data,
  cor = "#2563eb",
  eixoVertical = false,
}) {
  const dados = useMemo(() => {
    if (!Array.isArray(data)) return [];
    return data.map((item) => ({
      ...item,
      nomeOriginal: item.nomeOriginal ?? item.nome,
    }));
  }, [data]);

  if (eixoVertical) {
    return (
      <div className="bg-white/90 backdrop-blur rounded-3xl shadow-sm hover:shadow-md transition p-6 border border-gray-200/70 h-[420px]">
        <h2 className="font-semibold mb-3 text-base xl:text-lg">{titulo}</h2>
        <ResponsiveContainer width="100%" height="88%">
          <BarChart
            data={dados}
            layout="vertical"
            margin={{ top: 10, right: 42, left: 140, bottom: 10 }}
            barSize={16}
          >
            <CartesianGrid strokeDasharray="3 3" />
            <XAxis type="number" />
            <YAxis
              dataKey="nome"
              type="category"
              width={170}
              tick={{ fontSize: 11 }}
              interval={0}
            />
            <Tooltip
              formatter={(value, name) => [
                formatarNumero(value),
                name === "quantidade" ? "Boletos Liquidados" : name,
              ]}
              labelFormatter={(label, payload) =>
                payload && payload[0]?.payload?.nomeOriginal
                  ? payload[0].payload.nomeOriginal
                  : label
              }
            />
            <Bar dataKey="quantidade" fill={cor} radius={[0, 6, 6, 0]}>
              <LabelList
                dataKey="quantidade"
                position="right"
                formatter={(value) => formatarNumero(value)}
                style={{ fontWeight: 800, fill: "#111827", fontSize: 12 }}
              />
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>
    );
  }

  return (
    <div className="bg-white/90 backdrop-blur rounded-3xl shadow-sm hover:shadow-md transition p-6 border border-gray-200/70 h-[380px]">
      <h2 className="font-semibold mb-3 text-base xl:text-lg">{titulo}</h2>
      <ResponsiveContainer width="100%" height="88%">
        <BarChart
          data={dados}
          margin={{ top: 25, right: 20, left: 0, bottom: 10 }}
        >
          <CartesianGrid strokeDasharray="3 3" />
          <XAxis dataKey="nome" tick={{ fontSize: 11 }} interval={0} />
          <YAxis />
          <Tooltip
            formatter={(value, name) => [
              formatarNumero(value),
              name === "quantidade" ? "Boletos Liquidados" : name,
            ]}
            labelFormatter={(label, payload) =>
              payload && payload[0]?.payload?.nomeOriginal
                ? payload[0].payload.nomeOriginal
                : label
            }
          />
          <Bar dataKey="quantidade" fill={cor} radius={[6, 6, 0, 0]}>
            <LabelList
              dataKey="quantidade"
              position="top"
              formatter={(value) => formatarNumero(value)}
              style={{ fontWeight: 800, fill: cor, fontSize: 12 }}
            />
          </Bar>
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}

function GraficoBanco({ data }) {
  const dados = data.map((item) => ({
    ...item,
    nomeOriginal: item.nome,
    nome: abreviarBanco(item.nome),
  }));

  return (
    <div className="bg-white rounded-2xl shadow p-5 border border-gray-100 h-[380px]">
      <h2 className="font-semibold mb-3">Banco Arrecadador</h2>

      <ResponsiveContainer width="100%" height="88%">
        <BarChart
          data={dados}
          layout="vertical"
          margin={{ top: 5, right: 55, left: 120, bottom: 5 }}
          barSize={12}
        >
          <CartesianGrid strokeDasharray="3 3" />
          <XAxis type="number" />
          <YAxis
            dataKey="nome"
            type="category"
            width={170}
            tick={{ fontSize: 10 }}
            interval={0}
          />
          <Tooltip
            formatter={(value, name, props) => {
              if (name === "quantidade") {
                return [formatarNumero(value), "Boletos Liquidados"];
              }

              return [formatarMoeda(props.payload.valor), "Valor Liquidado"];
            }}
            labelFormatter={(label, payload) =>
              payload && payload[0]?.payload?.nomeOriginal
                ? payload[0].payload.nomeOriginal
                : label
            }
          />
          <Bar dataKey="quantidade" fill="#22c55e" radius={[0, 5, 5, 0]}>
            <LabelList
              dataKey="quantidade"
              position="right"
              formatter={(value) => formatarNumero(value)}
              style={{ fontWeight: 700, fill: "#111827", fontSize: 11 }}
            />
          </Bar>
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}

function TooltipBairro({ active, payload, label }) {
  if (!active || !payload || !payload.length) return null;

  const item = payload[0].payload;

  return (
    <div className="bg-white border border-gray-200 rounded-xl shadow p-3 text-sm">
      <p className="font-bold mb-1">{label}</p>
      <p>
        Boletos Liquidados:{" "}
        <span className="font-bold text-blue-600">
          {formatarNumero(item.quantidade)}
        </span>
      </p>
      <p>
        Valor Liquidado:{" "}
        <span className="font-bold text-blue-600">
          {formatarMoeda(item.valor)}
        </span>
      </p>
    </div>
  );
}

function GraficoBairros({ data }) {
  return (
    <div className="bg-white rounded-2xl shadow p-6 h-[430px] mb-4 border border-gray-100">
      <h2 className="text-xl font-semibold mb-4">Qtd por Bairro</h2>

      <ResponsiveContainer width="100%" height="85%">
        <BarChart
          data={data}
          margin={{ top: 15, right: 20, left: 0, bottom: 95 }}
          barSize={18}
        >
          <CartesianGrid strokeDasharray="3 3" />
          <XAxis
            dataKey="nome"
            interval={0}
            angle={-90}
            textAnchor="end"
            height={100}
            tick={{ fontSize: 10 }}
          />
          <YAxis />
          <Tooltip content={<TooltipBairro />} />
          <Bar dataKey="quantidade" fill="#2563eb" radius={[5, 5, 0, 0]} />
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}

function GraficoParcela({ data, parcelaSelecionada, setParcelaSelecionada }) {
  const total = data.reduce((acc, item) => acc + item.quantidade, 0);

  const dataFiltrada = parcelaSelecionada
    ? data.filter((item) => item.nome === parcelaSelecionada)
    : data;

  return (
    <div className="bg-white rounded-2xl shadow p-5 border border-gray-100 h-[380px] overflow-hidden">
      <div className="flex flex-col gap-2 mb-3">
        <h2 className="font-semibold">Tipo por Parcela</h2>

        <select
          className="border rounded-lg px-3 py-1.5 text-xs w-full"
          value={parcelaSelecionada}
          onChange={(e) => setParcelaSelecionada(e.target.value)}
        >
          <option value="">Todas as parcelas</option>
          {data.map((item) => (
            <option key={item.nome} value={item.nome}>
              {item.nome}
            </option>
          ))}
        </select>
      </div>

      <div className="grid grid-cols-1 xl:grid-cols-2 gap-3 h-[290px] overflow-hidden">
        <ResponsiveContainer width="100%" height="100%">
          <PieChart>
            <Pie
              data={dataFiltrada}
              dataKey="quantidade"
              nameKey="nome"
              outerRadius={78}
              labelLine={false}
              label={({ percent }) => {
                const percentual = percent * 100;
                return percentual >= 3 ? `${percentual.toFixed(1)}%` : "";
              }}
            >
              {dataFiltrada.map((item, index) => (
                <Cell
                  key={item.nome}
                  fill={coresPizza[index % coresPizza.length]}
                />
              ))}
            </Pie>
            <Tooltip
              formatter={(value, name, props) => [
                `${formatarNumero(value)} boletos`,
                props.payload.nome,
              ]}
            />
          </PieChart>
        </ResponsiveContainer>

        <div className="border rounded-xl p-3 overflow-auto text-xs max-h-[260px]">
          <div className="grid grid-cols-2 font-bold text-gray-500 mb-2">
            <span>Parcela</span>
            <span className="text-right">Boletos</span>
          </div>

          {dataFiltrada.map((item, index) => {
            const perc = total ? (item.quantidade / total) * 100 : 0;

            return (
              <div
                key={item.nome}
                className="grid grid-cols-2 items-center py-1"
              >
                <span className="flex items-center gap-2 truncate">
                  <span
                    className="w-3 h-3 rounded shrink-0"
                    style={{
                      backgroundColor: coresPizza[index % coresPizza.length],
                    }}
                  />
                  {item.nome}
                </span>

                <span className="text-right whitespace-nowrap">
                  <strong>{formatarNumero(item.quantidade)}</strong>{" "}
                  <span className="text-gray-500">({perc.toFixed(1)}%)</span>
                </span>
              </div>
            );
          })}

          <div className="border-t mt-2 pt-2 grid grid-cols-2 font-bold">
            <span>Total</span>
            <span className="text-right">{formatarNumero(total)}</span>
          </div>
        </div>
      </div>
    </div>
  );
}

function TooltipPerformance({ active, payload, label }) {
  if (!active || !payload || !payload.length) return null;

  const item = payload[0].payload;

  return (
    <div className="bg-white border border-gray-200 rounded-xl shadow p-3 text-sm">
      <p className="font-bold mb-2">{formatarDataBr(label)}</p>

      <p>
        <span className="text-orange-500">●</span> Valor Liquidado:{" "}
        <span className="font-bold text-orange-600">
          {formatarMoeda(item.valor)}
        </span>
      </p>

      <p>
        <span className="text-blue-600">●</span> Boletos Liquidados:{" "}
        <span className="font-bold text-blue-600">
          {formatarNumero(item.quantidade)}
        </span>
      </p>
    </div>
  );
}

function GraficoPerformance({ data }) {
  const totalValor = data.reduce((acc, item) => acc + Number(item.valor || 0), 0);
  const totalBoletos = data.reduce(
    (acc, item) => acc + Number(item.quantidade || 0),
    0
  );

  return (
    <div className="bg-white rounded-2xl shadow p-6 mb-4 border border-gray-100">
      <div className="flex flex-col xl:flex-row xl:justify-between xl:items-center gap-2 mb-4">
        <h2 className="text-xl font-semibold">
          Diário de desempenho — Valor Liquidado e Boletos por Dia
        </h2>

        <div className="border rounded-lg px-4 py-2 text-sm bg-white">
          Período filtrado
        </div>
      </div>

      <ResponsiveContainer width="100%" height={330}>
        <LineChart data={data} margin={{ top: 15, right: 40, left: 20, bottom: 10 }}>
          <CartesianGrid strokeDasharray="3 3" />

          <XAxis
            dataKey="nome"
            tickFormatter={formatarDataCurta}
            tick={{ fontSize: 11 }}
          />

          <YAxis
            yAxisId="valor"
            orientation="left"
            tickFormatter={(value) => formatarNumero(value)}
          />

          <YAxis
            yAxisId="quantidade"
            orientation="right"
            tickFormatter={(value) => formatarNumero(value)}
          />

          <Tooltip content={<TooltipPerformance />} />
          <Legend />

          <Line
            yAxisId="valor"
            type="monotone"
            dataKey="valor"
            name="Valor Liquidado (R$)"
            stroke="#f97316"
            strokeWidth={3}
            dot={{ r: 3 }}
          />

          <Line
            yAxisId="quantidade"
            type="monotone"
            dataKey="quantidade"
            name="Boletos Liquidados (Qtd)"
            stroke="#2563eb"
            strokeWidth={3}
            dot={{ r: 3 }}
          />
        </LineChart>
      </ResponsiveContainer>

      <div className="grid grid-cols-1 xl:grid-cols-2 gap-4 mt-5">
        <div className="border rounded-xl p-4 flex items-center gap-4">
          <div className="w-12 h-12 bg-orange-500 rounded-full flex items-center justify-center text-white font-bold">
            R$
          </div>
          <div>
            <p className="text-gray-600">Total do Período</p>
            <p className="text-2xl font-bold text-orange-600">
              {formatarMoeda(totalValor)}
            </p>
          </div>
        </div>

        <div className="border rounded-xl p-4 flex items-center gap-4">
          <div className="w-12 h-12 bg-blue-600 rounded-full flex items-center justify-center text-white font-bold">
            #
          </div>
          <div>
            <p className="text-gray-600">Boletos do Período</p>
            <p className="text-2xl font-bold text-blue-600">
              {formatarNumero(totalBoletos)}
            </p>
          </div>
        </div>
      </div>
    </div>
  );
}

function LinhaPeriodo({
  datasDisponiveis,
  indiceAtual,
  setIndiceAtual,
  dataFinalManual,
  setDataInicial,
}) {
  if (!datasDisponiveis.length) return null;

  const dataSelecionada = datasDisponiveis[indiceAtual] || datasDisponiveis[0];

  return (
    <div className="bg-white rounded-2xl shadow p-4 border border-gray-100">
      <div className="flex flex-col xl:flex-row xl:items-center xl:justify-between gap-3 mb-2">
        <h2 className="font-semibold">Linha de Período</h2>
        <div className="text-sm text-gray-600">
          Início selecionado: <strong>{formatarDataBr(dataSelecionada)}</strong>
          {dataFinalManual && (
            <span> | Final: <strong>{formatarDataBr(dataFinalManual)}</strong></span>
          )}
        </div>
      </div>

      <input
        type="range"
        min="0"
        max={datasDisponiveis.length - 1}
        value={indiceAtual}
        onChange={(e) => {
          const novoIndice = Number(e.target.value);
          setIndiceAtual(novoIndice);
          setDataInicial(datasDisponiveis[novoIndice] || "");
        }}
        className="w-full"
      />

      <div className="flex justify-between text-xs text-gray-500 mt-1">
        <span>{formatarDataBr(datasDisponiveis[0])}</span>
        <span>{formatarDataBr(datasDisponiveis[datasDisponiveis.length - 1])}</span>
      </div>
    </div>
  );
}

function App() {
  const [arrecadacao, setArrecadacao] = useState([]);
  const [municipes, setMunicipes] = useState([]);
  const [iptuDigital, setIptuDigital] = useState([]);
  const [importando, setImportando] = useState(false);

  const [filtroNome, setFiltroNome] = useState("");
  const [filtroCpf, setFiltroCpf] = useState("");
  const [filtroMatricula, setFiltroMatricula] = useState("");
  const [filtroCodigoBarras, setFiltroCodigoBarras] = useState("");
  const [filtroTributo, setFiltroTributo] = useState("");
  const [filtroBanco, setFiltroBanco] = useState("");
  const [mostrarFiltros, setMostrarFiltros] = useState(true);
  const [dataInicial, setDataInicial] = useState("");
  const [dataFinal, setDataFinal] = useState("");
  const [indicePeriodo, setIndicePeriodo] = useState(0);
  const [parcelaSelecionada, setParcelaSelecionada] = useState("");
  const [paginaAtual, setPaginaAtual] = useState(1);

  const itensPorPagina = 10;

  async function carregarDados() {
    const arrecadacaoData = await carregarTodosRegistros(
      "arrecadacao_bi",
      "liquidacao"
    );

    const municipesData = await carregarTodosRegistros("municipes");
    const iptuData = await carregarTodosRegistros("iptu_digital");

    setArrecadacao(arrecadacaoData || []);
    setMunicipes(municipesData || []);
    setIptuDigital(iptuData || []);
  }

  useEffect(() => {
    carregarDados();
  }, []);

  const datasDisponiveis = useMemo(() => {
    return [
      ...new Set(
        arrecadacao
          .map((item) => item.liquidacao)
          .filter((data) => data && String(data).includes("-"))
      ),
    ].sort((a, b) => String(a).localeCompare(String(b)));
  }, [arrecadacao]);

  useEffect(() => {
    if (datasDisponiveis.length > 0 && !dataInicial && !dataFinal) {
      setIndicePeriodo(0);
      setDataInicial(datasDisponiveis[0]);
      setDataFinal(datasDisponiveis[datasDisponiveis.length - 1]);
    }
  }, [datasDisponiveis, dataInicial, dataFinal]);

  async function importarMunicipes(event) {
    const arquivo = event.target.files[0];

    if (!arquivo) return;

    const confirmar = window.confirm(
      "Deseja substituir a base atual de Munícipes?"
    );

    if (!confirmar) {
      event.target.value = "";
      return;
    }

    setImportando(true);

    try {
      const baseLimpa = await limparTabela("municipes");

      if (!baseLimpa) return;

      const buffer = await arquivo.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];

      const linhas = XLSX.utils.sheet_to_json(sheet, {
        defval: "",
        raw: false,
      });

      const registros = linhas
        .map((linha) => ({
          id: Number(linha.id || linha.ID),
          cpf: limparTexto(linha.cpf || linha.CPF),
          nome: limparTexto(linha.nome || linha.NOME),
          telefone: limparTexto(linha.telefone || linha.TELEFONE),
          email: limparTexto(linha.email || linha.EMAIL),
          aceiteiptudigital: limparTexto(
            linha.aceiteiptudigital || linha.ACEITEIPTUDIGITAL
          ),
          status: limparTexto(linha.status || linha.STATUS),
          date_created: linha.date_created || linha.DATE_CREATED || null,
          last_updated: linha.last_updated || linha.LAST_UPDATED || null,
          cnpj: limparTexto(linha.cnpj || linha.CNPJ),
          razaosocial: limparTexto(linha.razaosocial || linha.RAZAOSOCIAL),
        }))
        .filter((item) => item.id);

      const tamanhoLote = 1000;

      for (let i = 0; i < registros.length; i += tamanhoLote) {
        const lote = registros.slice(i, i + tamanhoLote);

        const { error } = await supabase.from("municipes").insert(lote);

        if (error) {
          console.log(error);
          alert(`Erro ao importar Munícipes no lote ${i + 1}.`);
          return;
        }
      }

      alert(`Munícipes importados: ${formatarNumero(registros.length)}`);
      await carregarDados();
    } catch (erro) {
      console.log(erro);
      alert("Erro ao importar Munícipes.");
    } finally {
      setImportando(false);
      event.target.value = "";
    }
  }

  async function importarIptuDigital(event) {
    const arquivo = event.target.files[0];

    if (!arquivo) return;

    const confirmar = window.confirm(
      "Deseja substituir a base atual de IPTU Digital?"
    );

    if (!confirmar) {
      event.target.value = "";
      return;
    }

    setImportando(true);

    try {
      const baseLimpa = await limparTabela("iptu_digital");

      if (!baseLimpa) return;

      const buffer = await arquivo.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];

      const linhas = XLSX.utils.sheet_to_json(sheet, {
        defval: "",
        raw: false,
      });

      const registros = linhas
        .map((linha) => ({
          id: Number(linha.id || linha.ID),
          cpf: limparTexto(linha.cpf || linha.CPF),
          tokenconfirmacao: limparTexto(
            linha.tokenconfirmacao || linha.TOKENCONFIRMACAO
          ),
          data_aceite: linha.data_aceite || linha.DATA_ACEITE || null,
          data_vencimento: converterData(
            linha.data_vencimento || linha.DATA_VENCIMENTO
          ),
          protocolo: limparTexto(linha.protocolo || linha.PROTOCOLO),
          cnpj: limparTexto(linha.cnpj || linha.CNPJ),
        }))
        .filter((item) => item.id);

      const tamanhoLote = 1000;

      for (let i = 0; i < registros.length; i += tamanhoLote) {
        const lote = registros.slice(i, i + tamanhoLote);

        const { error } = await supabase.from("iptu_digital").insert(lote);

        if (error) {
          console.log(error);
          alert(`Erro ao importar IPTU Digital no lote ${i + 1}.`);
          return;
        }
      }

      alert(`IPTU Digital importado: ${formatarNumero(registros.length)}`);
      await carregarDados();
    } catch (erro) {
      console.log(erro);
      alert("Erro ao importar IPTU Digital.");
    } finally {
      setImportando(false);
      event.target.value = "";
    }
  }

  async function importarArrecadacao(event) {
    const arquivo = event.target.files[0];

    if (!arquivo) return;

    const confirmar = window.confirm(
      "Deseja substituir a base atual de Arrecadação BI?"
    );

    if (!confirmar) {
      event.target.value = "";
      return;
    }

    setImportando(true);

    try {
      const baseLimpa = await limparTabela("arrecadacao_bi");

      if (!baseLimpa) return;

      const buffer = await arquivo.arrayBuffer();

      const workbook = XLSX.read(buffer, {
        type: "array",
      });

      const sheet = workbook.Sheets[workbook.SheetNames[0]];

      const linhas = XLSX.utils.sheet_to_json(sheet, {
        defval: "",
        raw: false,
      });

      const registros = linhas
        .map((linha) => ({
          created_by: limparTexto(linha.created_by),
          status: limparTexto(linha.status),
          tributo: limparTexto(linha.tributo),
          liquidacao: converterData(linha.liquidacao),
          codigo_barra_prefeitura: limparTexto(
            linha.codigo_barra_prefeitura
          ),
          numero_lote: limparTexto(linha.numero_lote),
          data_movimento: converterData(linha.data_movimento),
          nome_arquivo: limparTexto(linha.nome_arquivo),
          sacado_bairro: limparTexto(linha.sacado_bairro),
          sacado_documento: limparTexto(linha.sacado_documento),
          sacado_nome: limparTexto(linha.sacado_nome),
          valor_pago: converterValor(linha.valor_pago),
          banco_pagamento: limparTexto(linha.banco_pagamento),
          parcelamento_info: limparTexto(linha.parcelamento_info),
          matricula: limparTexto(linha.matricula),
          tipo_pagamento: limparTexto(
            linha.tipo_pagamento || linha.TIPO_PAGAMENTO
          ),
          nome_banco_pagamento: limparTexto(
            linha.nome_banco_pagamento || linha.NOME_BANCO_PAGAMENTO
          ),
        }))
        .filter(
          (item) =>
            item.sacado_nome ||
            item.sacado_documento ||
            item.codigo_barra_prefeitura
        );

      const tamanhoLote = 1000;

      for (let i = 0; i < registros.length; i += tamanhoLote) {
        const lote = registros.slice(i, i + tamanhoLote);

        const { error } = await supabase.from("arrecadacao_bi").insert(lote);

        if (error) {
          console.log(error);
          alert(`Erro ao importar Arrecadação no lote ${i + 1}.`);
          return;
        }
      }

      alert(`Arrecadação importada: ${formatarNumero(registros.length)}`);
      setPaginaAtual(1);
      await carregarDados();
    } catch (erro) {
      console.log(erro);
      alert("Erro ao importar Arrecadação.");
    } finally {
      setImportando(false);
      event.target.value = "";
    }
  }

  const tributosDisponiveis = [
    ...new Set(arrecadacao.map((i) => i.tributo).filter(Boolean)),
  ];

  const bancosDisponiveis = [
    ...new Set(
      arrecadacao
        .map((i) => i.nome_banco_pagamento || i.banco_pagamento)
        .filter(Boolean)
    ),
  ];

  const dadosFiltrados = useMemo(() => {
    return arrecadacao.filter((item) => {
      const nome = String(item.sacado_nome || "").toLowerCase();
      const cpf = String(item.sacado_documento || "");
      const matricula = String(item.matricula || "");
      const codigo = String(item.codigo_barra_prefeitura || "");
      const banco = item.nome_banco_pagamento || item.banco_pagamento;
      const dataItem = item.liquidacao || "";

      return (
        (filtroNome ? nome.includes(filtroNome.toLowerCase()) : true) &&
        (filtroCpf ? cpf.includes(filtroCpf) : true) &&
        (filtroMatricula ? matricula.includes(filtroMatricula) : true) &&
        (filtroCodigoBarras ? codigo.includes(filtroCodigoBarras) : true) &&
        (filtroTributo ? item.tributo === filtroTributo : true) &&
        (filtroBanco ? banco === filtroBanco : true) &&
        (dataInicial ? dataItem >= dataInicial : true) &&
        (dataFinal ? dataItem <= dataFinal : true) &&
        (parcelaSelecionada
          ? normalizarParcela(item.tipo_pagamento) === parcelaSelecionada
          : true)
      );
    });
  }, [
    arrecadacao,
    filtroNome,
    filtroCpf,
    filtroMatricula,
    filtroCodigoBarras,
    filtroTributo,
    filtroBanco,
    dataInicial,
    dataFinal,
    parcelaSelecionada,
  ]);

  useEffect(() => {
    setPaginaAtual(1);
  }, [
    filtroNome,
    filtroCpf,
    filtroMatricula,
    filtroCodigoBarras,
    filtroTributo,
    filtroBanco,
    dataInicial,
    dataFinal,
    parcelaSelecionada,
  ]);

  const totalGuias = dadosFiltrados.length;

  const valorTotal = dadosFiltrados.reduce(
    (acc, item) => acc + Number(item.valor_pago || 0),
    0
  );

  const tipoCarne = ordenarMaiorQuantidade(
    agruparQuantidade(dadosFiltrados, "created_by", normalizarTipoCarne)
  );

  const tributos = ordenarMaiorQuantidade(
    agruparQuantidade(dadosFiltrados, "tributo")
  );

  // Versão apenas para exibição no gráfico: abrevia rótulos longos (mantém nomeOriginal no tooltip)
  const tributosTelao = useMemo(() => {
    return tributos.map((item) => ({
      ...item,
      nomeOriginal: item.nome,
      nome: abreviarTributo(item.nome),
    }));
  }, [tributos]);


  const bancos = ordenarMaiorQuantidade(
    agruparQuantidade(dadosFiltrados, "nome_banco_pagamento")
  ).slice(0, 10);

  const bairros = ordenarMaiorQuantidade(
    agruparQuantidade(dadosFiltrados, "sacado_bairro").filter((item) =>
      bairroValido(item.nome)
    )
  ).slice(0, 30);

  const parcelasOriginais = ordenarMaiorQuantidade(
    agruparQuantidade(dadosFiltrados, "tipo_pagamento", normalizarParcela)
  );

  const parcelas = parcelasOriginais.slice(0, 10);

  const performance = useMemo(() => {
    const mapa = {};

    dadosFiltrados.forEach((item) => {
      const dia = item.liquidacao || "Sem data";

      if (!mapa[dia]) {
        mapa[dia] = {
          nome: dia,
          valor: 0,
          quantidade: 0,
        };
      }

      mapa[dia].valor += Number(item.valor_pago || 0);
      mapa[dia].quantidade += 1;
    });

    return Object.values(mapa).sort((a, b) =>
      String(a.nome).localeCompare(String(b.nome))
    );
  }, [dadosFiltrados]);

  const totalPaginas = Math.max(1, Math.ceil(dadosFiltrados.length / 10));
  const inicio = (paginaAtual - 1) * 10;
  const fim = inicio + 10;
  const dadosTabela = dadosFiltrados.slice(inicio, fim);

  function limparFiltros() {
    setFiltroNome("");
    setFiltroCpf("");
    setFiltroMatricula("");
    setFiltroCodigoBarras("");
    setFiltroTributo("");
    setFiltroBanco("");
    setParcelaSelecionada("");
    setPaginaAtual(1);

    if (datasDisponiveis.length > 0) {
      setIndicePeriodo(0);
      setDataInicial(datasDisponiveis[0]);
      setDataFinal(datasDisponiveis[datasDisponiveis.length - 1]);
    }
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-gray-50 to-slate-100 text-gray-900">
      <header className="sticky top-0 z-50 bg-gradient-to-r from-yellow-500 via-amber-500 to-yellow-600 shadow-lg/20 border-b border-yellow-200/40">
      <div className="max-w-[1600px] mx-auto w-full min-h-20 flex items-center justify-between px-6 py-4">
        <div className="flex items-center gap-4">
          <div className="bg-white/95 backdrop-blur rounded-2xl p-3 shadow-sm border border-white/60">
            <Building2 size={36} className="text-yellow-600" />
          </div>

          <div>
            <h1 className="text-2xl xl:text-3xl font-extrabold tracking-tight text-white">SEMEF</h1>
            <p className="text-sm xl:text-base text-white/90">Painel de Arrecadação</p>
          </div>
        </div>

        <div className="flex gap-2 flex-wrap justify-end">
          <label className="bg-white/95 backdrop-blur px-4 py-2.5 rounded-2xl shadow-sm border border-white/60 text-sm font-semibold flex items-center gap-2 cursor-pointer transition-all hover:shadow-md hover:-translate-y-0.5 active:translate-y-0 focus-within:ring-2 focus-within:ring-yellow-200/80">
            <Upload size={16} />
            {importando ? "Importando..." : "Munícipes"}
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              className="hidden"
              onChange={importarMunicipes}
              disabled={importando}
            />
          </label>

          <label className="bg-white/95 backdrop-blur px-4 py-2.5 rounded-2xl shadow-sm border border-white/60 text-sm font-semibold flex items-center gap-2 cursor-pointer transition-all hover:shadow-md hover:-translate-y-0.5 active:translate-y-0 focus-within:ring-2 focus-within:ring-yellow-200/80">
            <Upload size={16} />
            {importando ? "Importando..." : "IPTU Digital"}
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              className="hidden"
              onChange={importarIptuDigital}
              disabled={importando}
            />
          </label>

          <label className="bg-white/95 backdrop-blur px-4 py-2.5 rounded-2xl shadow-sm border border-white/60 text-sm font-semibold flex items-center gap-2 cursor-pointer transition-all hover:shadow-md hover:-translate-y-0.5 active:translate-y-0 focus-within:ring-2 focus-within:ring-yellow-200/80">
            <Upload size={16} />
            {importando ? "Importando..." : "Arrecadação BI"}
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              className="hidden"
              onChange={importarArrecadacao}
              disabled={importando}
            />
          </label>
        </div>
          </div>
    </header>

      <main className="p-6 md:p-8 space-y-12">
  <div className="sticky top-24 z-40">
  <div className="max-w-[1600px] mx-auto space-y-4 bg-gradient-to-b from-slate-50/95 to-transparent backdrop-blur-sm pb-4">
<div className="bg-white/90 backdrop-blur rounded-3xl shadow-sm hover:shadow-md transition p-6 border border-gray-200/70">
          <div className="flex items-center justify-between gap-3 mb-5"><div className="flex items-center gap-3"><Search size={20} />
            <h2 className="text-xl xl:text-2xl font-semibold tracking-tight">Filtros Operacionais</h2></div><button type="button" onClick={() => setMostrarFiltros((v) => !v)} className="inline-flex items-center gap-2 px-4 py-2 rounded-2xl bg-white/80 backdrop-blur border border-gray-200 shadow-sm hover:shadow-md transition text-sm font-semibold">{mostrarFiltros ? (<>Ocultar <ChevronUp size={16} /></>) : (<>Mostrar <ChevronDown size={16} /></>)}</button></div>

          {mostrarFiltros && (
<div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-3">
            <input
              className="w-full border border-gray-200 rounded-2xl px-4 py-2.5 bg-white shadow-sm focus:outline-none focus:ring-2 focus:ring-yellow-400/60 focus:border-yellow-400 placeholder:text-gray-400"
              placeholder="Buscar por nome"
              value={filtroNome}
              onChange={(e) => setFiltroNome(e.target.value)}
            />

            <input
              className="w-full border border-gray-200 rounded-2xl px-4 py-2.5 bg-white shadow-sm focus:outline-none focus:ring-2 focus:ring-yellow-400/60 focus:border-yellow-400 placeholder:text-gray-400"
              placeholder="Buscar por CPF/CNPJ"
              value={filtroCpf}
              onChange={(e) => setFiltroCpf(e.target.value)}
            />

            <input
              className="w-full border border-gray-200 rounded-2xl px-4 py-2.5 bg-white shadow-sm focus:outline-none focus:ring-2 focus:ring-yellow-400/60 focus:border-yellow-400 placeholder:text-gray-400"
              placeholder="Buscar por matrícula"
              value={filtroMatricula}
              onChange={(e) => setFiltroMatricula(e.target.value)}
            />

            <input
              className="w-full border border-gray-200 rounded-2xl px-4 py-2.5 bg-white shadow-sm focus:outline-none focus:ring-2 focus:ring-yellow-400/60 focus:border-yellow-400 placeholder:text-gray-400"
              placeholder="Buscar por código de barras"
              value={filtroCodigoBarras}
              onChange={(e) => setFiltroCodigoBarras(e.target.value)}
            />

            <select
              className="w-full border border-gray-200 rounded-2xl px-4 py-2.5 bg-white shadow-sm focus:outline-none focus:ring-2 focus:ring-yellow-400/60 focus:border-yellow-400 placeholder:text-gray-400"
              value={filtroTributo}
              onChange={(e) => setFiltroTributo(e.target.value)}
            >
              <option value="">Todos os tributos</option>
              {tributosDisponiveis.map((item) => (
                <option key={item} value={item}>
                  {item}
                </option>
              ))}
            </select>

            <select
              className="w-full border border-gray-200 rounded-2xl px-4 py-2.5 bg-white shadow-sm focus:outline-none focus:ring-2 focus:ring-yellow-400/60 focus:border-yellow-400 placeholder:text-gray-400"
              value={filtroBanco}
              onChange={(e) => setFiltroBanco(e.target.value)}
            >
              <option value="">Todos os bancos</option>
              {bancosDisponiveis.map((item) => (
                <option key={item} value={item}>
                  {item}
                </option>
              ))}
            </select>

            <input
              type="date"
              className="w-full border border-gray-200 rounded-2xl px-4 py-2.5 bg-white shadow-sm focus:outline-none focus:ring-2 focus:ring-yellow-400/60 focus:border-yellow-400 placeholder:text-gray-400"
              value={dataInicial}
              onChange={(e) => setDataInicial(e.target.value)}
            />

            <input
              type="date"
              className="w-full border border-gray-200 rounded-2xl px-4 py-2.5 bg-white shadow-sm focus:outline-none focus:ring-2 focus:ring-yellow-400/60 focus:border-yellow-400 placeholder:text-gray-400"
              value={dataFinal}
              onChange={(e) => setDataFinal(e.target.value)}
            />

            <button
              onClick={limparFiltros}
              className="bg-gradient-to-r from-gray-900 to-gray-700 text-white rounded-2xl px-4 py-2.5 shadow-sm hover:shadow-md transition-all flex items-center justify-center gap-2"
            >
              <X size={16} />
              Limpar
            </button>

            <div className="border border-gray-200 rounded-2xl px-4 py-2.5 bg-white/80 backdrop-blur shadow-sm">
              <span className="text-gray-500 text-sm">Registros: </span>
              <strong>{formatarNumero(dadosFiltrados.length)}</strong>
            </div>
          </div>
)}
        </div>
<LinhaPeriodo
          datasDisponiveis={datasDisponiveis}
          indiceAtual={indicePeriodo}
          setIndiceAtual={setIndicePeriodo}
          dataFinalManual={dataFinal}
          setDataInicial={setDataInicial}
        />
  </div>
</div>

  <div className="max-w-[1600px] mx-auto space-y-12">

    <div className="grid grid-cols-2 xl:grid-cols-4 gap-7">
              <CardKpi
                titulo="Liquidação"
                valor={formatarNumero(totalGuias)}
                subtitulo="Guias Liquidadas"
                icon={CreditCard}
                cor="bg-blue-600"
              />

              <CardKpi
                titulo="Valor Arrecadado"
                valor={formatarMoeda(valorTotal)}
                subtitulo="Total Arrecadado"
                icon={Landmark}
                cor="bg-emerald-500"
              />

              <CardKpi
                titulo="Aplicativo Baixado"
                valor={formatarNumero(municipes.length)}
                subtitulo="Total de Downloads"
                icon={Smartphone}
                cor="bg-pink-500"
              />

              <CardKpi
                titulo="IPTU Digital"
                valor={formatarNumero(iptuDigital.length)}
                subtitulo="Total de Cadastros"
                icon={Building2}
                cor="bg-purple-500"
              />
            </div>

    <div className="grid grid-cols-1 xl:grid-cols-2 gap-7">
      <GraficoBarraVertical
                titulo="Tipo de Carnê"
                data={tipoCarne}
                cor="#2563eb"
              />
      <GraficoBarraVertical
                titulo="Tributo"
                data={tributosTelao}
                cor="#2563eb"
                eixoVertical={true}
              />
    </div>

    <div className="grid grid-cols-1 xl:grid-cols-2 gap-7">
      <GraficoBanco data={bancos} />
      <GraficoParcela
                data={parcelas}
                parcelaSelecionada={parcelaSelecionada}
                setParcelaSelecionada={setParcelaSelecionada}
              />
    </div>

    <GraficoBairros data={bairros} />

    <GraficoPerformance data={performance} />

    <div className="bg-white/90 backdrop-blur rounded-3xl shadow-sm hover:shadow-md transition p-6 border border-gray-200/70">
              <h2 className="text-2xl xl:text-3xl font-semibold tracking-tight mb-5">Tabela Operacional</h2>

              <div className="overflow-auto max-h-[560px] rounded-2xl border border-gray-100">
                <table className="w-full text-sm">
                  <thead className="sticky top-0 bg-gradient-to-r from-slate-50 to-gray-100 backdrop-blur border-b border-gray-200 z-10">
                    <tr>
                      <th className="p-3 text-left text-xs uppercase tracking-wider font-extrabold text-gray-600">Liquidação</th>
                      <th className="p-3 text-left text-xs uppercase tracking-wider font-extrabold text-gray-600">Nome</th>
                      <th className="p-3 text-left text-xs uppercase tracking-wider font-extrabold text-gray-600">Documento</th>
                      <th className="p-3 text-left text-xs uppercase tracking-wider font-extrabold text-gray-600">Matrícula</th>
                      <th className="p-3 text-left text-xs uppercase tracking-wider font-extrabold text-gray-600">Tributo</th>
                      <th className="p-3 text-left text-xs uppercase tracking-wider font-extrabold text-gray-600">Banco</th>
                      <th className="p-3 text-left text-xs uppercase tracking-wider font-extrabold text-gray-600">Código Barras</th>
                      <th className="p-3 text-left text-xs uppercase tracking-wider font-extrabold text-gray-600">Valor</th>
                    </tr>
                  </thead>

                  <tbody>
                    {dadosTabela.map((item) => (
                      <tr key={item.id} className="border-b border-gray-100 odd:bg-white even:bg-slate-50/70 hover:bg-amber-50/60 transition-colors">
                        <td className="p-3">{formatarDataBr(item.liquidacao)}</td>
                        <td className="p-3">{item.sacado_nome}</td>
                        <td className="p-3">{item.sacado_documento}</td>
                        <td className="p-3">{item.matricula}</td>
                        <td className="p-3">{item.tributo}</td>
                        <td className="p-3">
                          {item.nome_banco_pagamento || item.banco_pagamento}
                        </td>
                        <td className="p-3 max-w-[260px] truncate font-mono text-xs text-gray-700">
                          {item.codigo_barra_prefeitura}
                        </td>
                        <td className="p-3 font-bold text-emerald-700">
                          {formatarMoeda(item.valor_pago)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div className="flex items-center justify-between mt-4 text-sm">
                <div>
                  Mostrando {formatarNumero(dadosFiltrados.length ? inicio + 1 : 0)}{" "}
                  a {formatarNumero(Math.min(fim, dadosFiltrados.length))} de{" "}
                  {formatarNumero(dadosFiltrados.length)} registros
                </div>

                <div className="flex items-center gap-2">
                  <button
                    className="border border-gray-200 bg-white/80 backdrop-blur rounded-2xl px-3 py-2 shadow-sm hover:shadow-md transition disabled:opacity-40"
                    disabled={paginaAtual === 1}
                    onClick={() => setPaginaAtual((p) => Math.max(1, p - 1))}
                  >
                    <ChevronLeft size={16} />
                  </button>

                  <span className="bg-gradient-to-r from-blue-600 to-indigo-600 text-white px-4 py-2 rounded-2xl shadow-sm">
                    {paginaAtual}
                  </span>

                  <span>de {totalPaginas}</span>

                  <button
                    className="border border-gray-200 bg-white/80 backdrop-blur rounded-2xl px-3 py-2 shadow-sm hover:shadow-md transition disabled:opacity-40"
                    disabled={paginaAtual === totalPaginas}
                    onClick={() =>
                      setPaginaAtual((p) => Math.min(totalPaginas, p + 1))
                    }
                  >
                    <ChevronRight size={16} />
                  </button>
                </div>
              </div>
            </div>
  </div>
</main>
    </div>
  );
}

export default App;
