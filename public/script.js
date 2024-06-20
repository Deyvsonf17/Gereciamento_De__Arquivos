document.getElementById('searchForm').addEventListener('submit', function(event) {
    event.preventDefault();

    const name = normalizeText(document.getElementById('name').value);
    const dateOfBirth = document.getElementById('dateOfBirth').value.trim();

    // Verifica se ambos os campos estão vazios
    if (!name && !dateOfBirth) {
        $('#alertModal').modal('show');
        return;
    }

    const searchData = {
        name: name || null,
        dateOfBirth: dateOfBirth || null
    };

    fetch('/search', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(searchData)
    })
    .then(response => response.json())
    .then(data => {
        console.log('Received Data:', data);
        displayResults(data);
        // Limpar os campos do formulário após a busca
        document.getElementById('searchForm').reset();
    })
    .catch(error => console.error('Error:', error));
});

function displayResults(results) {
    const resultsDiv = document.getElementById('results');
    resultsDiv.innerHTML = '';

    if (!Array.isArray(results) || results.length === 0) {
        resultsDiv.textContent = 'Nenhum resultado encontrado.';
        return;
    }

    const nameGroups = groupByNameAndDate(results);

    nameGroups.forEach((group, key) => {
        const [name, dateOfBirth] = key.split('-');

        const resultDiv = document.createElement('div');
        resultDiv.className = 'result';

        const resultHeader = document.createElement('div');
        resultHeader.className = 'result-header';

        resultHeader.innerHTML = `
            <h4><i class="fas fa-user"></i> <strong>${name.toUpperCase()}</strong></h4>
            <button class="btn btn-outline-primary" onclick="redirectToStudentGrades('${name}', '${dateOfBirth}')"><strong>Notas do aluno</strong></button>
        `;
        resultDiv.appendChild(resultHeader);

        const dateParagraph = document.createElement('p');
        dateParagraph.innerHTML = `<i class="fas fa-calendar-alt"></i> Data de Nascimento: <strong>${dateOfBirth}</strong>`;
        resultDiv.appendChild(dateParagraph);

        if (group.length > 1) {
            const warningDiv = document.createElement('div');
            warningDiv.className = 'result-warning';
            warningDiv.innerHTML = `
                <p><i class="fas fa-exclamation-triangle"></i> Aviso:<br> Esse nome foi encontrado em várias planilhas. <br> Para manter os dados organizados, é recomendado consolidar as informações em uma única planilha.</p>
            `;
            resultDiv.appendChild(warningDiv);
        }

        const linksParagraph = document.createElement('p');
        linksParagraph.textContent = 'Para abrir o nome no Excel, clique no link abaixo.';
        resultDiv.appendChild(linksParagraph);

        const ul = document.createElement('ul');
        group.forEach(item => {
            const li = document.createElement('li');
            li.innerHTML = `<a class=".nav-link-open-excel" href="#" onclick="openExcelModal('${item.sheet}', '${item.numeracao}')"><i class="fas fa-file-excel"></i> <strong>${item.sheet}</strong></a>`;
            ul.appendChild(li);
        });
        resultDiv.appendChild(ul);

        resultsDiv.appendChild(resultDiv);
    });
}

function groupByNameAndDate(results) {
    const groups = new Map();
    
    results.forEach(result => {
        const key = `${normalizeText(result.name)}-${result.dateOfBirth}`;
        if (!groups.has(key)) {
            groups.set(key, []);
        }
        groups.get(key).push(result);
    });
    return groups;
}

function openExcelModal(sheet, numeracao) {
    document.getElementById('openExcelBtn').onclick = function() {
        openExcel(sheet, numeracao);
    };
    $('#excelModal').modal('show');
}

function openExcel(sheet, numeracao) {
    fetch(`/open-excel?sheet=${encodeURIComponent(sheet)}&numeracao=${encodeURIComponent(numeracao)}`)
        .then(response => {
            if (response.ok) {
                console.log('Excel opened');
            } else {
                console.error('Error opening Excel');
            }
        })
        .catch(error => console.error('Error:', error));
}

function normalizeText(name) {
    // Remove acentos
    name = name.normalize('NFD').replace(/[\u0300-\u036f]/g, '');

    // Remove espaços extras
    name = name.replace(/\s+/g, ' ').trim();

    // Converte para maiúsculas
    name = name.toUpperCase();

    // Remove duplicações (exemplo: "da da", "do do", "dos dos")
    name = name.split(' ').filter((word, index, array) => {
        return index === 0 || word !== array[index - 1];
    }).join(' ');

    return name;
}

function redirectToStudentGrades(name, dateOfBirth) {
    const encodedName = encodeURIComponent(name);
    const encodedDateOfBirth = encodeURIComponent(dateOfBirth);
    window.location.href = `/student-grades.html`;
}
