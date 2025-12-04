/**
 * Extrai o mês de referência do nome do arquivo
 * Exemplos:
 * - "Faturamento_Novembro_2025.xlsx" -> "2025-11"
 * - "Empenho_1411_Novembro_2025_BR.xlsx" -> "2025-11"
 * - "planilha-123456.xlsx" -> mês atual
 */
export function extrairMesReferencia(nomeArquivo: string): string {
  const mesesMap: { [key: string]: string } = {
    'janeiro': '01',
    'fevereiro': '02',
    'março': '03',
    'marco': '03',
    'abril': '04',
    'maio': '05',
    'junho': '06',
    'julho': '07',
    'agosto': '08',
    'setembro': '09',
    'outubro': '10',
    'novembro': '11',
    'dezembro': '12'
  };

  const nomeMinusculo = nomeArquivo.toLowerCase();
  
  // Procurar padrão: mes_ano ou mes ano (ex: Novembro_2025)
  for (const [mes, numero] of Object.entries(mesesMap)) {
    const regexComAno = new RegExp(`${mes}[_\\s]+(\\d{4})`, 'i');
    const match = nomeMinusculo.match(regexComAno);
    
    if (match) {
      const ano = match[1];
      return `${ano}-${numero}`;
    }
  }
  
  // Procurar padrão: ano_mes (ex: 2025_Novembro)
  for (const [mes, numero] of Object.entries(mesesMap)) {
    const regexAnoMes = new RegExp(`(\\d{4})[_\\s]+${mes}`, 'i');
    const match = nomeMinusculo.match(regexAnoMes);
    
    if (match) {
      const ano = match[1];
      return `${ano}-${numero}`;
    }
  }
  
  // Procurar apenas o mês com ano em qualquer lugar do nome
  for (const [mes, numero] of Object.entries(mesesMap)) {
    if (nomeMinusculo.includes(mes)) {
      // Procurar por 4 dígitos (ano) em qualquer lugar
      const anoMatch = nomeMinusculo.match(/(\d{4})/);
      if (anoMatch) {
        return `${anoMatch[1]}-${numero}`;
      }
    }
  }
  
  // Se não encontrar, usar mês/ano atual
  const agora = new Date();
  const ano = agora.getFullYear();
  const mes = String(agora.getMonth() + 1).padStart(2, '0');
  
  return `${ano}-${mes}`;
}

/**
 * Formata o mês de referência para exibição
 * "2025-11" -> "Novembro/2025"
 */
export function formatarMesReferencia(mesRef: string): string {
  const mesesNomes = [
    'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
    'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
  ];
  
  const [ano, mes] = mesRef.split('-');
  const mesIndex = parseInt(mes) - 1;
  
  return `${mesesNomes[mesIndex]}/${ano}`;
}
