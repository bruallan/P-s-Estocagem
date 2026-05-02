import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { UploadCloud, FileSpreadsheet, Download, Search, CheckSquare, Square } from 'lucide-react';

interface Product {
  codigo: string;
  nome: string;
  precoCusto: number;
}

interface Market {
  name: string;
  productCodes: Set<string>;
}

interface MappingInfo {
  prodHeaderRow: number;
  prodNomeCol: string;
  prodCodCol: string;
  prodCustoCol: string;
  planHeaderRow: number;
  planCodCol: string;
}

const getColLetter = (index: number) => {
  if (index < 0) return 'N/A';
  let letter = '';
  let temp = index;
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
};

interface MissingProductRow {
  mercado: string;
  codigo: string;
  produto: string;
  precoCusto: number;
  precoSugerido: number;
  margem20: number;
  margem27: number;
}

const App: React.FC = () => {
  const [produtosFile, setProdutosFile] = useState<File | null>(null);
  const [planogramasFile, setPlanogramasFile] = useState<File | null>(null);
  const [productsList, setProductsList] = useState<Product[]>([]);
  const [marketsList, setMarketsList] = useState<Market[]>([]);
  const [selectedMarkets, setSelectedMarkets] = useState<Set<string>>(new Set());
  const [isProcessing, setIsProcessing] = useState(false);
  const [dropdownOpen, setDropdownOpen] = useState(false);
  const [mappingInfo, setMappingInfo] = useState<MappingInfo | null>(null);

  const handleProdutosUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      setProdutosFile(e.target.files[0]);
    }
  };

  const handlePlanogramasUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      setPlanogramasFile(e.target.files[0]);
    }
  };

  const parseExcelFiles = async () => {
    if (!produtosFile || !planogramasFile) return;

    setIsProcessing(true);
    try {
      // Parse Produtos
      const prodData = await produtosFile.arrayBuffer();
      const prodWorkbook = XLSX.read(prodData, { type: 'array' });
      const prodSheet = prodWorkbook.Sheets[prodWorkbook.SheetNames[0]];
      const prodJson: any[][] = XLSX.utils.sheet_to_json(prodSheet, { header: 1 });

      let pHeaderRow = -1;
      let pCodCol = -1;
      let pNomeCol = -1;
      let pCustoCol = -1;

      for (let i = 0; i < Math.min(prodJson.length, 50); i++) {
        const row = prodJson[i];
        if (!row) continue;
        const rowStr = row.map(c => String(c).toLowerCase()).join('|');
        if (rowStr.includes('código') && rowStr.normalize("NFD").replace(/[\u0300-\u036f]/g, "").includes('custo')) {
          pHeaderRow = i;
          pCodCol = row.findIndex(c => c && String(c).toLowerCase().trim() === 'código');
          pNomeCol = row.findIndex(c => c && String(c).toLowerCase().trim() === 'nome');
          pCustoCol = row.findIndex(c => c && String(c).toLowerCase().includes('custo atual'));
          if (pCustoCol === -1) {
            pCustoCol = row.findIndex(c => c && String(c).toLowerCase().includes('preço de custo'));
          }
          if (pCustoCol === -1) {
            pCustoCol = row.findIndex(c => c && String(c).toLowerCase().includes('custo'));
          }
          break;
        }
      }

      const products: Product[] = [];
      if (pHeaderRow !== -1) {
        for (let i = pHeaderRow + 1; i < prodJson.length; i++) {
          const row = prodJson[i];
          if (!row || row.length === 0) continue;

          const codigo = String(row[pCodCol] ?? '').trim();
          const nome = String(row[pNomeCol] ?? '').trim();
          let custoRaw = row[pCustoCol];

          if (codigo && nome) {
            let precoCusto = 0;
            if (typeof custoRaw === 'number') {
              precoCusto = custoRaw;
            } else if (typeof custoRaw === 'string') {
              let clean = custoRaw.replace('R$', '').trim();
              if (clean.includes(',') && clean.includes('.')) {
                clean = clean.replace(/\./g, '').replace(',', '.');
              } else if (clean.includes(',')) {
                clean = clean.replace(',', '.');
              }
              precoCusto = parseFloat(clean) || 0;
            }
            products.push({ codigo, nome, precoCusto });
          }
        }
      }

      // Parse Planogramas
      const planData = await planogramasFile.arrayBuffer();
      const planWorkbook = XLSX.read(planData, { type: 'array' });
      const markets: Market[] = [];
      let globalPlanHeader = -1;
      let globalPlanCod = -1;

      for (const sheetName of planWorkbook.SheetNames) {
        if (sheetName.toLowerCase() === 'resumo') continue;

        const sheet = planWorkbook.Sheets[sheetName];
        const planJson: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        let planHeaderRow = -1;
        let planCodCol = -1;

        for (let i = 0; i < Math.min(planJson.length, 30); i++) {
          const row = planJson[i];
          if (!row) continue;
          
          const rowTitles = row.map(c => String(c).toLowerCase().trim());
          planCodCol = rowTitles.findIndex(c => c === 'código');
          if (planCodCol !== -1) {
            planHeaderRow = i;
            if (globalPlanHeader === -1) {
              globalPlanHeader = i;
              globalPlanCod = planCodCol;
            }
            break;
          }
        }

        const codes = new Set<string>();
        if (planHeaderRow !== -1) {
          for (let i = planHeaderRow + 1; i < planJson.length; i++) {
            const row = planJson[i];
            const rawCodigo = row?.[planCodCol];
            if (rawCodigo !== undefined && rawCodigo !== null && String(rawCodigo).trim() !== '') {
              // Extract the base code (sometimes formats might differ)
              codes.add(String(rawCodigo).trim());
            }
          }
        }

        markets.push({ name: sheetName, productCodes: codes });
      }

      setProductsList(products);
      setMarketsList(markets);
      setSelectedMarkets(new Set(markets.map(m => m.name)));

      setMappingInfo({
        prodHeaderRow: pHeaderRow + 1,
        prodNomeCol: getColLetter(pNomeCol),
        prodCodCol: getColLetter(pCodCol),
        prodCustoCol: getColLetter(pCustoCol),
        planHeaderRow: globalPlanHeader + 1,
        planCodCol: getColLetter(globalPlanCod)
      });

    } catch (error) {
      console.error("Erro ao processar as planilhas:", error);
      alert("Houve um erro ao processar as planilhas. Verifique se o formato está correto.");
    } finally {
      setIsProcessing(false);
    }
  };

  const missingProductsData: MissingProductRow[] = useMemo(() => {
    const data: MissingProductRow[] = [];
    
    // Sort selected markets conceptually or iteration is fine
    const activeMarkets = marketsList.filter(m => selectedMarkets.has(m.name));

    for (const market of activeMarkets) {
      for (const prod of productsList) {
        if (!market.productCodes.has(prod.codigo)) {
          
          const cost = prod.precoCusto;
          const exactSuggestedPrice = cost / 0.58;
          // Arredondar para o número que termine em 9 seguinte
          // Math.ceil(exactSuggestedPrice * 10) / 10 - 0.01
          const suggestedPrice = Math.max(0, parseFloat((Math.ceil(exactSuggestedPrice * 10) / 10 - 0.01).toFixed(2)));

          let margem20 = 0;
          let margem27 = 0;

          if (suggestedPrice > 0) {
            margem20 = 1 - (cost / suggestedPrice) - 0.20;
            margem27 = 1 - (cost / suggestedPrice) - 0.27;
          }

          data.push({
            mercado: market.name,
            codigo: prod.codigo,
            produto: prod.nome,
            precoCusto: cost,
            precoSugerido: suggestedPrice,
            margem20: margem20,
            margem27: margem27
          });
        }
      }
    }

    return data;
  }, [marketsList, productsList, selectedMarkets]);

  const toggleMarket = (marketName: string) => {
    const next = new Set(selectedMarkets);
    if (next.has(marketName)) {
      next.delete(marketName);
    } else {
      next.add(marketName);
    }
    setSelectedMarkets(next);
  };

  const toggleAllMarkets = () => {
    if (selectedMarkets.size === marketsList.length) {
      setSelectedMarkets(new Set());
    } else {
      setSelectedMarkets(new Set(marketsList.map(m => m.name)));
    }
  };

  const formatCurrency = (value: number) => {
    return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(value);
  };

  const formatPercent = (value: number) => {
    return new Intl.NumberFormat('pt-BR', { style: 'percent', minimumFractionDigits: 1, maximumFractionDigits: 2 }).format(value);
  };

  const exportToExcel = () => {
    if (missingProductsData.length === 0) return;

    const exportData = missingProductsData.map(row => ({
      'Mercado': row.mercado,
      'Código': row.codigo,
      'Produto': row.produto,
      'Preço de Custo': row.precoCusto,
      'Preço Sugerido (R$)': row.precoSugerido,
      'Margem c/ 20% Custo Op': row.margem20,
      'Margem c/ 27% Custo Op': row.margem27,
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Faltantes");

    XLSX.writeFile(workbook, "produtos_faltantes.xlsx");
  };

  const isAllMarketsSelected = selectedMarkets.size === marketsList.length && marketsList.length > 0;
  const isSomeMarketsSelected = selectedMarkets.size > 0 && selectedMarkets.size < marketsList.length;

  return (
    <div className="min-h-screen bg-gray-50 text-gray-900 font-sans flex flex-col items-center">
      <header className="w-full bg-blue-700 text-white shadow-md p-6">
        <div className="max-w-7xl mx-auto">
          <h1 className="text-3xl font-bold tracking-tight">Help4U Analisador de Planogramas</h1>
          <p className="text-blue-100 mt-2">Identifique produtos faltantes em seus mini mercados</p>
        </div>
      </header>

      <main className="w-full max-w-7xl px-6 py-8 flex-1 grid grid-cols-1 gap-8">
        
        {/* Upload Section */}
        <section className="bg-white rounded-xl shadow-sm border border-gray-200 p-6">
          <h2 className="text-xl font-semibold mb-6 flex items-center gap-2">
            <UploadCloud className="text-blue-600" /> Importar Dados
          </h2>
          
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="flex flex-col gap-2">
              <label className="font-medium text-sm text-gray-700">Planilha de Planogramas</label>
              <div className="relative group">
                <input 
                  type="file" 
                  accept=".xls,.xlsx" 
                  onChange={handlePlanogramasUpload}
                  className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10" 
                />
                <div className={`p-4 border-2 border-dashed rounded-lg flex items-center gap-3 transition-colors ${planogramasFile ? 'border-green-500 bg-green-50' : 'border-gray-300 group-hover:border-blue-400 bg-gray-50'}`}>
                  <FileSpreadsheet className={planogramasFile ? 'text-green-600' : 'text-gray-400'} size={24} />
                  <div className="flex-1 truncate">
                    <p className="text-sm font-medium text-gray-700 truncate">
                      {planogramasFile ? planogramasFile.name : "Clique ou arraste Planogramas.xls"}
                    </p>
                  </div>
                </div>
              </div>
            </div>

            <div className="flex flex-col gap-2">
              <label className="font-medium text-sm text-gray-700">Planilha de Produtos</label>
              <div className="relative group">
                <input 
                  type="file" 
                  accept=".xls,.xlsx" 
                  onChange={handleProdutosUpload}
                  className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10" 
                />
                <div className={`p-4 border-2 border-dashed rounded-lg flex items-center gap-3 transition-colors ${produtosFile ? 'border-green-500 bg-green-50' : 'border-gray-300 group-hover:border-blue-400 bg-gray-50'}`}>
                  <FileSpreadsheet className={produtosFile ? 'text-green-600' : 'text-gray-400'} size={24} />
                  <div className="flex-1 truncate">
                    <p className="text-sm font-medium text-gray-700 truncate">
                      {produtosFile ? produtosFile.name : "Clique ou arraste Produtos.xlsx"}
                    </p>
                  </div>
                </div>
              </div>
            </div>
          </div>

          <div className="mt-8 flex justify-end">
            <button 
              onClick={parseExcelFiles}
              disabled={!produtosFile || !planogramasFile || isProcessing}
              className="bg-blue-600 hover:bg-blue-700 disabled:bg-gray-300 disabled:cursor-not-allowed text-white px-6 py-2 rounded-lg font-medium shadow-sm transition-all flex items-center gap-2"
            >
              {isProcessing ? (
                <>Processando...</>
              ) : (
                <>
                  <Search size={18} />
                  Analisar Faltantes
                </>
              )}
            </button>
          </div>
        </section>

        {/* Results Section */}
        {marketsList.length > 0 && (
          <section className="bg-white rounded-xl shadow-sm border border-gray-200 flex flex-col h-[800px]">
            {mappingInfo && (
              <div className="mx-6 mt-6 p-4 bg-blue-50 text-blue-800 border border-blue-200 rounded-lg text-sm flex gap-2 flex-col">
                <h3 className="font-semibold text-blue-900 border-b border-blue-200 pb-2 mb-1 flex items-center gap-2">
                  <Search size={16} /> Diagnóstico de Leitura (Como Busquei os Dados)
                </h3>
                <p>Caso algum valor não bata, verifique se a coluna identificada está correta:</p>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mt-1">
                  <div>
                    <strong className="text-blue-900">Planilha: Produtos</strong>
                    <ul className="list-disc pl-5 opacity-90 mt-1 space-y-1">
                      <li>Cabeçalho: <strong>Linha {mappingInfo.prodHeaderRow}</strong></li>
                      <li>Código: <strong>Coluna {mappingInfo.prodCodCol}</strong></li>
                      <li>Nome: <strong>Coluna {mappingInfo.prodNomeCol}</strong></li>
                      <li>Preço de Custo: <strong className="text-blue-900 bg-blue-200/50 px-1 py-0.5 rounded">Coluna {mappingInfo.prodCustoCol}</strong></li>
                    </ul>
                  </div>
                  <div>
                    <strong className="text-blue-900">Planilha: Planogramas</strong>
                    <ul className="list-disc pl-5 opacity-90 mt-1 space-y-1">
                      <li>Cabeçalho: <strong>Linha {mappingInfo.planHeaderRow}</strong></li>
                      <li>Código: <strong>Coluna {mappingInfo.planCodCol}</strong></li>
                    </ul>
                  </div>
                </div>
              </div>
            )}

            <div className="p-6 border-b border-gray-200 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
              <h2 className="text-xl font-semibold flex items-center gap-2">
                Resultados ({missingProductsData.length})
              </h2>

              <div className="flex items-center gap-3">
                {/* Custom Multi-Select Dropdown */}
                <div className="relative">
                  <button 
                    onClick={() => setDropdownOpen(!dropdownOpen)}
                    className="flex justify-between items-center gap-2 border border-gray-300 bg-white px-4 py-2 rounded-lg text-sm font-medium hover:bg-gray-50 min-w-[200px]"
                  >
                    <span>
                      {isAllMarketsSelected ? 'Todos os mercados' : `${selectedMarkets.size} mercado(s) selecionado(s)`}
                    </span>
                    <span className="text-xs">▼</span>
                  </button>

                  {dropdownOpen && (
                     <>
                      <div className="fixed inset-0 z-20" onClick={() => setDropdownOpen(false)} />
                      <div className="absolute right-0 mt-2 w-72 bg-white border border-gray-200 shadow-xl rounded-lg z-30 max-h-[400px] overflow-hidden flex flex-col">
                        <div className="p-2 border-b border-gray-100 flex items-center justify-between bg-gray-50">
                          <button 
                            onClick={toggleAllMarkets}
                            className="flex items-center gap-2 text-sm font-medium text-gray-700 hover:text-blue-600 p-2 rounded-md hover:bg-blue-50 transition w-full"
                          >
                            {isAllMarketsSelected ? (
                              <CheckSquare className="text-blue-600" size={18} />
                            ) : (
                              <Square className="text-gray-400" size={18} />
                            )}
                            Selecionar Todos
                          </button>
                        </div>
                        <div className="overflow-y-auto p-2 flex-col gap-1">
                          {marketsList.map(market => (
                            <button
                               key={market.name}
                               onClick={() => toggleMarket(market.name)}
                               className="flex items-center gap-2 text-sm text-gray-700 hover:bg-gray-100 p-2 w-full text-left rounded-md transition"
                            >
                              {selectedMarkets.has(market.name) ? (
                                <CheckSquare className="text-blue-600" size={18} />
                              ) : (
                                <Square className="text-gray-400" size={18} />
                              )}
                              <span className="truncate">{market.name}</span>
                            </button>
                          ))}
                        </div>
                      </div>
                    </>
                  )}
                </div>

                <button 
                  onClick={exportToExcel}
                  className="flex items-center gap-2 bg-green-600 hover:bg-green-700 text-white px-4 py-2 rounded-lg font-medium text-sm transition-colors shadow-sm"
                >
                  <Download size={18} />
                  Exportar
                </button>
              </div>
            </div>

            <div className="flex-1 overflow-auto bg-gray-50 border-t border-gray-200">
              <table className="w-full text-left border-collapse text-sm">
                <thead className="bg-white sticky top-0 z-10 shadow-sm ring-1 ring-gray-200">
                  <tr>
                    <th className="px-4 py-3 font-semibold text-gray-600">Mercado</th>
                    <th className="px-4 py-3 font-semibold text-gray-600">Código</th>
                    <th className="px-4 py-3 font-semibold text-gray-600 w-1/4">Produto</th>
                    <th className="px-4 py-3 font-semibold text-gray-600 text-right">Preço Custo</th>
                    <th className="px-4 py-3 font-semibold text-blue-700 text-right bg-blue-50/50">Preço Sugerido</th>
                    <th className="px-4 py-3 font-semibold text-gray-600 text-right">Mgm (20%)</th>
                    <th className="px-4 py-3 font-semibold text-gray-600 text-right">Mgm (27%)</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-200">
                  {missingProductsData.length === 0 ? (
                    <tr>
                      <td colSpan={7} className="px-4 py-12 text-center text-gray-500 bg-white">
                        Nenhum produto faltante encontrado para a seleção atual.
                      </td>
                    </tr>
                  ) : (
                    missingProductsData.map((row, idx) => (
                      <tr key={idx} className="bg-white hover:bg-gray-50 transition-colors">
                        <td className="px-4 py-2 font-medium text-gray-800 border-x border-gray-100">{row.mercado}</td>
                        <td className="px-4 py-2 text-gray-500 font-mono text-xs border-r border-gray-100">{row.codigo}</td>
                        <td className="px-4 py-2 text-gray-700 truncate border-r border-gray-100" title={row.produto}>{row.produto}</td>
                        <td className="px-4 py-2 text-right text-gray-600 border-r border-gray-100">{formatCurrency(row.precoCusto)}</td>
                        <td className="px-4 py-2 text-right font-medium text-blue-700 bg-blue-50/30 border-r border-gray-100">{formatCurrency(row.precoSugerido)}</td>
                        <td className="px-4 py-2 text-right text-gray-600 border-r border-gray-100">
                          {row.precoSugerido > 0 ? (
                            <span className={row.margem20 < 0 ? 'text-red-500' : 'text-green-600'}>
                              {formatPercent(row.margem20)}
                            </span>
                          ) : '-'}
                        </td>
                        <td className="px-4 py-2 text-right text-gray-600 border-r border-gray-100">
                          {row.precoSugerido > 0 ? (
                            <span className={row.margem27 < 0 ? 'text-red-500' : 'text-green-600'}>
                              {formatPercent(row.margem27)}
                            </span>
                          ) : '-'}
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
            
            <div className="p-4 border-t border-gray-200 text-xs text-gray-500 flex justify-between bg-white text-center">
              <span>* <strong>Preço Sugerido</strong> = Custo / (1 - (27% + 15%)) arredondado para próximo .x9</span>
              <span><strong>Margem</strong> = 1 - (Custo / Sugerido) - % Op</span>
            </div>
          </section>
        )}
      </main>
    </div>
  );
};

export default App;
