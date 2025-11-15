// Função para atualizar o nome do arquivo com animação
function updateFileName(input, displayElement) {
    if (input.files && input.files.length > 0) {
        const fileName = input.files[0].name;
        const fileSize = (input.files[0].size / 1024 / 1024).toFixed(2);
        displayElement.querySelector('span').innerHTML = `${fileName} <small style="opacity: 0.7;">(${fileSize} MB)</small>`;
        displayElement.style.display = 'flex';
        displayElement.style.animation = 'fadeInUp 0.3s ease-out';
    } else {
        displayElement.style.display = 'none';
    }
}

// Função para destacar área de upload quando arquivo é selecionado
function highlightUploadArea(uploadArea) {
    uploadArea.style.borderColor = 'var(--success)';
    setTimeout(() => {
        uploadArea.style.borderColor = '';
    }, 1000);
}

// Função para mostrar erro
function showError(message) {
    const errorAlert = document.getElementById('errorAlert');
    if (errorAlert) {
        errorAlert.textContent = message;
        errorAlert.style.display = 'block';
        setTimeout(() => {
            errorAlert.style.display = 'none';
        }, 5000);
    }
}

// Configurar eventos para CapaSimples
function setupCapaSimples() {
    const capaFileInput = document.getElementById('capaFileInput');
    const capaFileName = document.getElementById('capaFileName');
    const capaUploadArea = document.getElementById('capaUploadArea');

    if (!capaFileInput || !capaFileName || !capaUploadArea) return;

    capaFileInput.addEventListener('change', function() {
        updateFileName(this, capaFileName);
        if (this.files.length > 0) {
            highlightUploadArea(capaUploadArea);
        }
    });

    capaUploadArea.addEventListener('click', function(e) {
        if (e.target !== capaFileInput && !e.target.closest('.file-name')) {
            capaFileInput.click();
        }
    });

    capaUploadArea.addEventListener('dragover', function(e) {
        e.preventDefault();
        e.stopPropagation();
        this.classList.add('dragover');
    });

    capaUploadArea.addEventListener('dragleave', function(e) {
        e.preventDefault();
        e.stopPropagation();
        this.classList.remove('dragover');
    });

    capaUploadArea.addEventListener('drop', function(e) {
        e.preventDefault();
        e.stopPropagation();
        this.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) {
            const dataTransfer = new DataTransfer();
            dataTransfer.items.add(e.dataTransfer.files[0]);
            capaFileInput.files = dataTransfer.files;
            updateFileName(capaFileInput, capaFileName);
            highlightUploadArea(capaUploadArea);
        }
    });
}

// Configurar eventos para Extra_Movimentações
function setupExtraMovimentacoes() {
    const extraFileInput = document.getElementById('extraFileInput');
    const extraFileName = document.getElementById('extraFileName');
    const extraUploadArea = document.getElementById('extraUploadArea');

    if (!extraFileInput || !extraFileName || !extraUploadArea) return;

    extraFileInput.addEventListener('change', function() {
        updateFileName(this, extraFileName);
        if (this.files.length > 0) {
            highlightUploadArea(extraUploadArea);
        }
    });

    extraUploadArea.addEventListener('click', function(e) {
        if (e.target !== extraFileInput && !e.target.closest('.file-name')) {
            extraFileInput.click();
        }
    });

    extraUploadArea.addEventListener('dragover', function(e) {
        e.preventDefault();
        e.stopPropagation();
        this.classList.add('dragover');
    });

    extraUploadArea.addEventListener('dragleave', function(e) {
        e.preventDefault();
        e.stopPropagation();
        this.classList.remove('dragover');
    });

    extraUploadArea.addEventListener('drop', function(e) {
        e.preventDefault();
        e.stopPropagation();
        this.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) {
            const dataTransfer = new DataTransfer();
            dataTransfer.items.add(e.dataTransfer.files[0]);
            extraFileInput.files = dataTransfer.files;
            updateFileName(extraFileInput, extraFileName);
            highlightUploadArea(extraUploadArea);
        }
    });
}

// Validação e envio do formulário
function setupFormValidation() {
    const processForm = document.getElementById('processForm');
    if (!processForm) return;

    processForm.addEventListener('submit', function(e) {
        e.preventDefault();
        
        const capaFileInput = document.getElementById('capaFileInput');
        const extraFileInput = document.getElementById('extraFileInput');
        const submitBtn = document.getElementById('submitBtn');
        const btnText = document.getElementById('btnText');
        const loadingSpinner = document.getElementById('loadingSpinner');

        const capaFile = capaFileInput && capaFileInput.files.length > 0;
        const extraFile = extraFileInput && extraFileInput.files.length > 0;

        if (!capaFile || !extraFile) {
            showError('Por favor, selecione ambos os arquivos antes de processar.');
            processForm.classList.add('shake');
            setTimeout(() => {
                processForm.classList.remove('shake');
            }, 500);
            return false;
        }

        // Desabilitar botão e mostrar loading
        if (submitBtn) {
            submitBtn.disabled = true;
        }
        if (btnText) {
            btnText.textContent = 'Processando...';
        }
        if (loadingSpinner) {
            loadingSpinner.style.display = 'inline-block';
        }

        // Criar FormData e enviar
        const formData = new FormData();
        formData.append('CapaSimplesFile', capaFileInput.files[0]);
        formData.append('ExtraMovimentacoesFile', extraFileInput.files[0]);

        fetch('./Home/Processar', {
            method: 'POST',
            body: formData
        })
        .then(response => {
            if (response.ok) {
                return response.blob();
            } else {
                return response.json().then(data => {
                    throw new Error(data.error || 'Erro ao processar');
                });
            }
        })
        .then(blob => {
            // Criar link de download
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `PlanilhaTratada_${new Date().toISOString().replace(/[:.]/g, '-')}.xlsx`;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);

            // Resetar formulário
            processForm.reset();
            document.getElementById('capaFileName').style.display = 'none';
            document.getElementById('extraFileName').style.display = 'none';
            
            // Reabilitar botão
            if (submitBtn) submitBtn.disabled = false;
            if (btnText) btnText.textContent = 'Processar e Baixar Planilha Tratada';
            if (loadingSpinner) loadingSpinner.style.display = 'none';
        })
        .catch(error => {
            showError(error.message);
            if (submitBtn) submitBtn.disabled = false;
            if (btnText) btnText.textContent = 'Processar e Baixar Planilha Tratada';
            if (loadingSpinner) loadingSpinner.style.display = 'none';
        });
    });
}

// Inicializar quando o DOM estiver pronto
document.addEventListener('DOMContentLoaded', function() {
    setupCapaSimples();
    setupExtraMovimentacoes();
    setupFormValidation();
});