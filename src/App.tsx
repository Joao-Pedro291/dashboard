import { useState } from "react";
import * as XLSX from "xlsx";
import "./App.css";

type MesData = {
  mes: string;
  receita: number;
  despesa: number;
  lucro: number;
};

type ExcelRow = {
  data_p?: string | number;
  data_v?: string | number;
  valor: number | string;
};

export default function App() {
  const [data, setData] = useState<MesData[]>([]);

  const getMes = (data: string): string | null => {
    // precisa ser YYYYMMDD
    if (!data || data.length !== 8) return null;

    const ano = data.slice(0, 4);
    const mes = data.slice(4, 6);

    // 🔥 FILTRO: só 2026
    if (ano !== "2026") return null;

    return `${ano}-${mes}`;
  };

  const parseValor = (valor: string | number): number => {
    if (typeof valor === "number") return valor;

    const limpo = valor
      .replace(/\./g, "") // remove TODOS os pontos
      .replace(",", "."); // troca vírgula

    return Number(limpo);
  };

  const formatMes = (mes: string): string => {
    const [ano, mesNum] = mes.split("-");
    return `${mesNum}/${ano}`;
  };

  const handleFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();

    reader.onload = (evt) => {
      const result = evt.target?.result;
      if (!result) return;

      const wb = XLSX.read(result, { type: "binary" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json<ExcelRow>(ws);

      const porMes: Record<string, MesData> = {};

      json.forEach((item) => {
        // 💰 conversão correta BR
        const valor = parseValor(item.valor);
        if (isNaN(valor)) return;

        // 📅 pega data completa
        const raw = item.data_p;
        const data = String(raw ?? "").replace(/\D/g, "");

        const mes = getMes(data);
        if (!mes) return;

        if (!porMes[mes]) {
          porMes[mes] = {
            mes,
            receita: 0,
            despesa: 0,
            lucro: 0,
          };
        }

        if (valor > 0) {
          porMes[mes].receita += valor;
        } else {
          porMes[mes].despesa += Math.abs(valor);
        }

        porMes[mes].lucro = porMes[mes].receita - porMes[mes].despesa;
      });
      const mesesCompletos: MesData[] = [];

      for (let i = 1; i <= 12; i++) {
        const mesFormatado = `2026-${String(i).padStart(2, "0")}`;

        if (!porMes[mesFormatado]) {
          mesesCompletos.push({
            mes: mesFormatado,
            receita: 0,
            despesa: 0,
            lucro: 0,
          });
        } else {
          mesesCompletos.push(porMes[mesFormatado]);
        }
      }

      setData(mesesCompletos);
    };

    reader.readAsBinaryString(file);
  };

  return (
    <div className="container">
      <h1 className="title">📊 Fluxo de Caixa 2026</h1>

      <input type="file" onChange={handleFile} className="file-input" />

      <div className="card">
        <table className="table">
          <thead>
            <tr>
              <th>Mês</th>
              <th className="right">Receita</th>
              <th className="right">Despesa</th>
              <th className="right">Líquido</th>
            </tr>
          </thead>

          <tbody>
            {data.map((item) => (
              <tr key={item.mes}>
                <td>{formatMes(item.mes)}</td>

                <td className="right green">
                  {item.receita.toLocaleString("pt-BR", {
                    style: "currency",
                    currency: "BRL",
                  })}
                </td>

                <td className="right red">
                  {item.despesa.toLocaleString("pt-BR", {
                    style: "currency",
                    currency: "BRL",
                  })}
                </td>

                <td
                  className={`right ${item.lucro >= 0 ? "blue" : "negative"}`}
                >
                  {item.lucro.toLocaleString("pt-BR", {
                    style: "currency",
                    currency: "BRL",
                  })}
                </td>
              </tr>
            ))}
          </tbody>
        </table>

        {data.length === 0 && (
          <p className="empty">Nenhum dado de 2026 encontrado</p>
        )}
      </div>
    </div>
  );
}
