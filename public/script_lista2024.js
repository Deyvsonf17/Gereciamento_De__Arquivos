document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('searchForm2024');
    const resultsContainer = document.getElementById('results');
    let alertMessage = document.getElementById('alertMessage');

    form.addEventListener('submit', async (event) => {
        event.preventDefault();
        resultsContainer.innerHTML = '';
        alertMessage = document.getElementById('alertMessage'); // Atualiza referência ao elemento

        if (!alertMessage) {
            console.error('Element alertMessage not found.');
            return;
        }

        alertMessage.innerHTML = ''; // Limpa mensagens de alerta anteriores

        let name = document.getElementById('name').value.trim();
        const dateOfBirth = document.getElementById('dateOfBirth').value.trim();
        const sigem = document.getElementById('sigem').value.trim();

        // Normaliza o nome
        name = normalizeName(name);

        // Verifica quantos campos foram preenchidos
        const filledFields = [name, dateOfBirth, sigem].filter(field => field !== '');

        // Caso mais de um campo seja preenchido, exibe mensagem de alerta
        if (filledFields.length > 1) {
            alertMessage.innerHTML = '<div class="alert alert-warning" role="alert">Por favor, escolha apenas um campo para preencher.</div>';
            return;
        }

        // Caso nenhum campo seja preenchido, exibe mensagem de alerta
        if (filledFields.length === 0) {
            alertMessage.innerHTML = '<div class="alert alert-danger" role="alert">Por favor, preencha pelo menos um campo.</div>';
            return;
        }

        // Só um campo foi preenchido, continua com a busca
        try {
            const response = await fetch('/search2024', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ name, dateOfBirth, sigem })
            });

            const results = await response.json();

            // Verificar duplicidade em todas as planilhas
            const duplicates = checkForDuplicatesInAllSheets(results, name, dateOfBirth);

            if (duplicates.length > 1) {
                alertMessage.innerHTML = '<div class="alert alert-warning" role="alert">Este aluno está presente em mais de uma planilha. Recomenda-se verificar.</div>';
            }

            if (results.length === 0) {
                resultsContainer.innerHTML = '<p class="text-center">Nenhum resultado encontrado.</p>';
                return;
            }

            // Limpar os campos do formulário após a busca
            form.reset();

            results.forEach(result => {
                const div = document.createElement('div');
                div.classList.add('result-item');

                const sigemCode = result.sigem && /^\d+$/.test(result.sigem) ? result.sigem : 'O aluno está sem código';
                const observation = result.observation ? result.observation : 'O aluno está sem observação';

                div.innerHTML = `
                    <div class="result-header">
                        <h4><i class="fas fa-user"></i> <strong>${result.name}</strong></h4>
                        <button class="btn btn-outline-primary sigem-btn" data-sigem="${sigemCode}"><strong>Acessar o Sigem</strong></button>
                    </div>
                    <p><i class="fas fa-calendar-alt"></i> Data de Nascimento: <strong>${result.dateOfBirth}</strong></p>
                    <p><i class="fas fa-id-badge"></i> Código SIGEM: <strong>${sigemCode}</strong></p>
                    <p><i class="fas fa-info-circle"></i> Observação: <strong>${observation}</strong></p>
                    <p>Para abrir o nome no Excel, clique no link abaixo.</p>
                `;

                const ul = document.createElement('ul');
                const li = document.createElement('li');
                li.innerHTML = `<a href="#" class="open-excel" data-sheet="${result.sheet || ''}" data-numeracao="${result.numeracao || ''}"><i class="fas fa-file-excel"></i> <strong>${result.sheet || 'N/A'}</strong></a>`;
                ul.appendChild(li);
                div.appendChild(ul);

                resultsContainer.appendChild(div);

                div.querySelectorAll('.open-excel').forEach(link => {
                    link.addEventListener('click', async (event) => {
                        event.preventDefault();
                        const sheet = link.getAttribute('data-sheet');
                        const numeracao = link.getAttribute('data-numeracao');

                        if (!sheet || !numeracao) {
                            console.error('Sheet or Numeracao is missing!');
                            return;
                        }

                        // Verificar se o Excel já está aberto
                        const excelOpen = await checkIfExcelOpen();

                        if (excelOpen) {
                            $('#excelOpenModal').modal('show');

                            document.getElementById('forceCloseExcelBtn').addEventListener('click', async () => {
                                $('#excelOpenModal').modal('hide');
                                await openExcel(sheet, numeracao);
                            });
                        } else {
                            await openExcel(sheet, numeracao);
                        }
                    });
                });

                div.querySelectorAll('.sigem-btn').forEach(button => {
                    button.addEventListener('click', async () => {
                        const sigemCode = button.getAttribute('data-sigem');
                        $('#loadingModal').modal('show');

                        setTimeout(() => {
                            window.open(`https://educacao.moreno.pe.gov.br/`, '_blank');
                            $('#loadingModal').modal('hide');
                        }, 2000);
                    });
                });
            });
        } catch (error) {
            console.error('Erro ao buscar alunos:', error);
            alert('Erro ao buscar alunos. Por favor, tente novamente mais tarde.');
        }
    });

    function normalizeName(name) {
        // Remove acentos
        name = name.normalize('NFD').replace(/[\u0300-\u036f]/g, '');

        // Remove espaços extras
        name = name.replace(/\s+/g, ' ').trim();

        // Converte para minúsculas
        name = name.toLowerCase();

        // Remove duplicações (exemplo: "da da", "do do", "dos dos")
        name = name.split(' ').filter((word, index, array) => {
            return index === 0 || word !== array[index - 1];
        }).join(' ');

        return name;
    }

    function checkForDuplicatesInAllSheets(results, name, dateOfBirth) {
        const duplicates = [];

        for (let result of results) {
            if (normalizeName(result.name) === normalizeName(name) && result.dateOfBirth === dateOfBirth) {
                duplicates.push(result);
            }
        }

        return duplicates;
    }

    async function checkIfExcelOpen() {
        // Simulando a verificação se o Excel está aberto
        return false; // Altere para true se desejar simular o Excel aberto
    }

    async function openExcel(sheet, numeracao) {
        $('#excelModal').modal('show');

        document.getElementById('openExcelBtn').addEventListener('click', async () => {
            $('#excelModal').modal('hide');
            await fetch(`/open-excel-2024?sheet=${sheet}&numeracao=${numeracao}`)
        });
    }
});
