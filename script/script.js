function gerarExcel() {
    const dadosDdd = {
        "estadoPorDdd": {
            "11": "SP", "12": "SP", "13": "SP", "14": "SP", "15": "SP", "16": "SP", "17": "SP", "18": "SP", "19": "SP",
            "21": "RJ", "22": "RJ", "24": "RJ", "27": "ES", "28": "ES", "31": "MG", "32": "MG", "33": "MG", "34": "MG",
            "35": "MG", "37": "MG", "38": "MG", "41": "PR", "42": "PR", "43": "PR", "44": "PR", "45": "PR", "46": "PR",
            "47": "SC", "48": "SC", "49": "SC", "51": "RS", "53": "RS", "54": "RS", "55": "RS", "61": "DF", "62": "GO",
            "63": "TO", "64": "GO", "65": "MT", "66": "MT", "67": "MS", "68": "AC", "69": "RO", "71": "BA", "73": "BA",
            "74": "BA", "75": "BA", "77": "BA", "79": "SE", "81": "PE", "82": "AL", "83": "PB", "84": "RN", "85": "CE",
            "86": "PI", "87": "PE", "88": "CE", "89": "PI", "91": "PA", "92": "AM", "93": "PA", "94": "PA", "95": "RR",
            "96": "AP", "97": "AM", "98": "MA", "99": "MA"
        }
    };

    function formatarNumero(numero) {
        numero = numero.replace(/\D/g, '');

        if (numero.startsWith('55')) {
            numero = numero.substring(2);
        }

        if (numero.length === 8 || numero.length === 9) {
            numero = '84' + numero;
        }

        const ddd = numero.substring(0, 2);
        let resto = numero.substring(2);

        if (!['SP', 'RJ', 'ES'].includes(dadosDdd.estadoPorDdd[ddd])) {
            if (resto.length === 8) {
                resto = '9' + resto;
            }
        } else {
            if (resto.startsWith('9') && resto.length === 9) {
                resto = resto.substring(1);
            }
        }

        numero = '55' + ddd + resto;

        return numero;
    }

    const entradaContatos = document.getElementById('entradaContatos').value;

    const linhas = entradaContatos.trim().split('\n');
    const contatos = linhas.map(linha => {
        const [nome, numeroBruto] = linha.split(',');
        const numeroFormatado = formatarNumero(numeroBruto.trim());

        return { nome: nome.trim(), contato: numeroFormatado };
    });

    const planilha = XLSX.utils.json_to_sheet(contatos);
    const pastaTrabalho = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(pastaTrabalho, planilha, 'Contatos');

    XLSX.writeFile(pastaTrabalho, 'contatos.xlsx');
}
