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

    if (numero.length > 11 && numero.startsWith('55')) {
        numero = numero.substring(2);
    }

    const ddd = numero.substring(0, 2);
    let resto = numero.substring(2);

    if (resto.length === 8) {
        resto = '9' + resto;
    } else if (resto.length === 9 && resto[0] !== '9') {
        resto = '9' + resto;
    }

    if (resto.length === 9 && resto[0] !== '9') {
        return null; 
    }

    numero = '55' + ddd + resto;

    return numero;
}

    const entradaContatos = document.getElementById('entradaContatos').value;

    const linhas = entradaContatos.trim().split('\n');
    const numerosInvalidos = [];

    const contatos = linhas.map((linha, index) => {
        const partes = linha.split(/[;,]+/);

        if (partes.length !== 2) {
            exibirModal(`A linha ${index + 1} está incompleta. Por favor, adicione o nome e o número (com DDD).`);
            marcarNumeroInvalido(index + 1);
            return null;
        }

        const [nome, numeroBruto] = partes;
        const numeroLimpo = numeroBruto.replace(/\D/g, '');

        if (numeroLimpo.length < 10 || numeroLimpo.length > 13) {
            numerosInvalidos.push(numeroBruto);
            marcarNumeroInvalido(index + 1);
            return null;
        }

        const numeroFormatado = formatarNumero(numeroBruto.trim());

        if (!numeroFormatado) {
            numerosInvalidos.push(numeroBruto);
            marcarNumeroInvalido(index + 1);
            return null;
        }

        return { nome: nome.trim(), contato: numeroFormatado };
    }).filter(Boolean);

    contatos.forEach((contato, index) => {
        const numeroLimpo = contato.contato.replace(/\D/g, '');
        const ddd = numeroLimpo.substring(2, 4);

        if (!(ddd in dadosDdd.estadoPorDdd)) {
            numerosInvalidos.push(contato.contato);
            marcarNumeroInvalido(index + 1);
        } else {
            if (numeroLimpo.length !== 13) {
                numerosInvalidos.push(contato.contato);
                marcarNumeroInvalido(index + 1);
            } else {
                if (numeroLimpo[4] !== '9') {
                    numerosInvalidos.push(contato.contato);
                    marcarNumeroInvalido(index + 1);
                }
            }
        }
    });

    if (numerosInvalidos.length > 0) {
        const numerosInvalidosMsg = numerosInvalidos.map(numero => `"${numero}"`).join(', ');
        exibirModal(`Corrija os números antes de gerar a planilha, siga o modelo "nome, contato", lembre-se do DDD.`);
        return;
    }

    if (contatos.length === 0) {
        exibirModal('Nenhum contato válido foi inserido, certifique-se de usar o modelo informado (nome, telefone).');
        return;
    }

    const planilha = XLSX.utils.json_to_sheet(contatos);
    const pastaTrabalho = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(pastaTrabalho, planilha, 'Contatos');

    XLSX.writeFile(pastaTrabalho, 'contatos.xlsx');
}

function exibirModal(mensagem) {
    const modal = document.getElementById('modalAviso');
    const modalMessage = document.getElementById('modalMessage');
    modal.style.display = 'block';
    modalMessage.textContent = mensagem;
}

function marcarNumeroInvalido(numeroLinha) {
    const inputContatos = document.getElementById('entradaContatos');
    let linhas = inputContatos.value.split('\n');
    const linhaInvalida = linhas[numeroLinha - 1];

    if (!linhaInvalida.startsWith('-->')) {
        linhas[numeroLinha - 1] = `--> ${linhaInvalida}`;
        inputContatos.value = linhas.join('\n');
    }
}

const spanClose = document.getElementsByClassName('close')[0];
spanClose.onclick = function() {
    const modal = document.getElementById('modalAviso');
    modal.style.display = 'none';
};

window.onclick = function(event) {
    const modal = document.getElementById('modalAviso');
    if (event.target === modal) {
        modal.style.display = 'none';
    }
};
