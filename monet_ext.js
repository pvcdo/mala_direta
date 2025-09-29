/**
 * Converte um valor numérico para sua representação monetária por extenso (em Reais).
 * Esta função aceita um único valor ou um intervalo (para uso com ARRAYFORMULA).
 * Exemplo: =EXTENSO(A1:A10)
 *
 * @param {Array<Array<number>> | number} input O valor numérico ou um intervalo (Array 2D) de valores.
 * @return {Array<Array<string>> | string} O valor monetário por extenso ou um Array 2D de resultados.
 * @customfunction
 */
function EXTENSO(input) {
  // 1. Verifica se a entrada é um array (intervalo)
  if (Array.isArray(input)) {
    const results = [];
    
    // Itera sobre as linhas do array
    input.forEach(row => {
      const rowResults = [];
      // Itera sobre as células de cada linha
      row.forEach(cellValue => {
        // Aplica a lógica de conversão a cada célula
        rowResults.push(_numeroParaMoedaPorExtenso(cellValue));
      });
      results.push(rowResults);
    });
    
    return results; // Retorna o array 2D de resultados
  } else {
    // 2. Se for um valor único, apenas o converte (mantém a compatibilidade)
    return _numeroParaMoedaPorExtenso(input);
  }
}


// --- LÓGICA CENTRAL DE CONVERSÃO (Auxiliar) ---
// Esta função contém toda a lógica de escrita por extenso do código anterior.
// Ela não deve ser chamada diretamente na planilha.

function _numeroParaMoedaPorExtenso(valor) {
  // *** COLE AQUI TODO O CÓDIGO DA FUNÇÃO EXTENSO ANTERIOR ***
  // Certifique-se de que a primeira linha da sua função auxiliar comece com:
  // if (typeof valor !== 'number' || isNaN(valor)) { return ""; }
  
  if (typeof valor !== 'number' || isNaN(valor) || valor === "") {
    // Retorna vazio para células vazias ou não numéricas (importante para ARRAYFORMULA)
    return ""; 
  }
  
  // Define os blocos de leitura
  const unidades = ['', 'um', 'dois', 'três', 'quatro', 'cinco', 'seis', 'sete', 'oito', 'nove'];
  const dezenas = ['', 'dez', 'vinte', 'trinta', 'quarenta', 'cinquenta', 'sessenta', 'setenta', 'oitenta', 'noventa'];
  const centenas = ['', 'cento', 'duzentos', 'trezentos', 'quatrocentos', 'quinhentos', 'seiscentos', 'setecentos', 'oitocentos', 'novecentos'];
  const especiais = ['dez', 'onze', 'doze', 'treze', 'quatorze', 'quinze', 'dezesseis', 'dezessete', 'dezoito', 'dezenove'];
  const milhares = ['', 'mil', 'milhão', 'bilhão', 'trilhão'];

  // Função auxiliar para ler números de 0 a 999
  function lerBloco(num) {
    let str = '';
    let n = parseInt(num);

    if (n === 100) return 'cem';
    if (n === 0) return '';

    // Centenas
    if (n >= 100) {
      str += centenas[Math.floor(n / 100)];
      n %= 100;
      if (n > 0) str += ' e ';
    }

    // Dezenas e Unidades
    if (n > 0) {
      if (n < 10) {
        str += unidades[n];
      } else if (n < 20) {
        str += especiais[n - 10];
      } else {
        str += dezenas[Math.floor(n / 10)];
        n %= 10;
        if (n > 0) str += ' e ' + unidades[n];
      }
    }
    return str;
  }

  // Separa a parte inteira (reais) e a decimal (centavos)
  let inteiro = Math.floor(valor).toString();
  let centavos = Math.round((valor - Math.floor(valor)) * 100).toString();

  // Garante que os centavos tenham 2 dígitos
  if (centavos.length === 1) centavos = '0' + centavos;
  centavos = parseInt(centavos);

  let textoReais = '';
  let blocos = [];

  // Divide o número inteiro em blocos de 3 dígitos (milhares)
  while (inteiro.length > 0) {
    const bloco = inteiro.slice(-3);
    blocos.unshift(bloco);
    inteiro = inteiro.slice(0, -3);
  }

  // Processa cada bloco
  for (let i = 0; i < blocos.length; i++) {
    const num = parseInt(blocos[i]);
    if (num > 0) {
      const idxMilhar = blocos.length - 1 - i;
      let textoBloco = lerBloco(num);

      if (idxMilhar > 0) {
        // Adiciona o nome do milhar (mil, milhão, bilhão, etc.)
        if (idxMilhar === 1 && num === 1) { // Caso especial para 'mil'
          textoBloco = 'mil';
        } else if (idxMilhar === 1) {
          textoBloco += ' mil';
        } else {
          const nomeMilhar = milhares[idxMilhar];
          textoBloco += ' ' + nomeMilhar;
          // Adiciona plural se o bloco for maior que 1
          if (num > 1) {
            textoBloco += nomeMilhar.endsWith('l') ? 'hões' : 'ões';
          }
        }
      }
      // Conecta os blocos com ' e ' ou ' '
      if (textoReais.length > 0) {
        const ultimoChar = textoReais.slice(-1);
        if (ultimoChar === 'o' || ultimoChar === 'a') {
             textoReais += ' e '; // Se o último for ão, ões, ...
        } else {
             textoReais += ' ';
        }
      }
      textoReais += textoBloco;
    }
  }

  // Lógica do Real/Reais
  if (textoReais.length === 0) {
    textoReais = 'zero';
  } else if (textoReais.includes('um mil')) {
    textoReais = textoReais.replace('um mil', 'mil');
  }

  // Adiciona a moeda
  let parteInteira = textoReais;
  const realPlural = (Math.floor(valor) === 1 && valor < 2) ? 'real' : 'reais';
  
  parteInteira += ' ' + realPlural;

  // Lógica dos centavos
  let parteDecimal = '';
  if (centavos > 0) {
    const centavoPlural = (centavos === 1) ? 'centavo' : 'centavos';
    parteDecimal = lerBloco(centavos) + ' ' + centavoPlural;
  }

  // Conecta Reais e Centavos
  let resultado = parteInteira;
  if (centavos > 0) {
    resultado += ' e ' + parteDecimal;
  }
  
  // Formato: Capitaliza a primeira letra e remove excesso de espaços.
  resultado = resultado.trim().toLowerCase();
  resultado = resultado.charAt(0).toUpperCase() + resultado.slice(1);
  
  return resultado;
}