class XLSBConverter {
    constructor() {
        this.taskId = null;
        this.startTime = null;
        this.progressInterval = null;
        this.initializeEventListeners();
    }

    initializeEventListeners() {
        const fileInput = document.getElementById('fileInput');
        const uploadArea = document.getElementById('uploadArea');

        fileInput.addEventListener('change', (e) => {
            this.handleFileSelect(e.target.files[0]);
        });

        // Drag and drop
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });

        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragover');
        });

        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            if (e.dataTransfer.files.length) {
                this.handleFileSelect(e.dataTransfer.files[0]);
            }
        });
    }

    handleFileSelect(file) {
        if (!file) return;

        const fileInfo = document.getElementById('fileInfo');
        const errorMessage = document.getElementById('errorMessage');
        
        // Validar tipo de arquivo
        if (!file.name.toLowerCase().endsWith('.xlsb')) {
            this.showError('Por favor, selecione um arquivo .xlsb');
            return;
        }

        errorMessage.style.display = 'none';
        fileInfo.innerHTML = `📄 Arquivo selecionado: <strong>${file.name}</strong> (${this.formatFileSize(file.size)})`;

        this.uploadAndConvert(file);
    }

    async uploadAndConvert(file) {
        const formData = new FormData();
        formData.append('file', file);

        try {
            const response = await fetch('/upload', {
                method: 'POST',
                body: formData
            });

            const data = await response.json();

            if (data.error) {
                this.showError(data.error);
                return;
            }

            this.taskId = data.task_id;
            this.startConversionMonitoring();
            
        } catch (error) {
            this.showError('Erro ao enviar arquivo: ' + error.message);
        }
    }

    startConversionMonitoring() {
        this.startTime = Date.now();
        this.updateElapsedTime();
        
        const progressContainer = document.getElementById('progressContainer');
        progressContainer.style.display = 'block';

        this.progressInterval = setInterval(async () => {
            await this.checkProgress();
        }, 1000);
    }

    async checkProgress() {
        if (!this.taskId) return;

        try {
            const response = await fetch(`/progress/${this.taskId}`);
            const progress = await response.json();

            this.updateProgressUI(progress);

            if (progress.status === 'completo') {
                this.onConversionComplete(progress);
            } else if (progress.status === 'erro') {
                this.onConversionError(progress.error);
            }
        } catch (error) {
            console.error('Erro ao verificar progresso:', error);
        }
    }

    updateProgressUI(progress) {
        const progressFill = document.getElementById('progressFill');
        const progressPercent = document.getElementById('progressPercent');
        const progressMessage = document.getElementById('progressMessage');
        const timeInfo = document.getElementById('timeInfo');

        progressFill.style.width = `${progress.progress}%`;
        progressPercent.textContent = `${Math.round(progress.progress)}%`;
        progressMessage.textContent = progress.message;

        // Mostrar tempo decorrido quando a conversão iniciar
        if (progress.progress > 0) {
            timeInfo.style.display = 'block';
            this.updateElapsedTime();
        }
    }

    updateElapsedTime() {
        if (!this.startTime) return;

        const elapsedTimeElement = document.getElementById('elapsedTime');
        const elapsedSeconds = Math.floor((Date.now() - this.startTime) / 1000);
        
        const minutes = Math.floor(elapsedSeconds / 60);
        const seconds = elapsedSeconds % 60;
        
        elapsedTimeElement.textContent = `${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
    }

    onConversionComplete(progress) {
        clearInterval(this.progressInterval);
        
        const downloadSection = document.getElementById('downloadSection');
        const downloadLink = document.getElementById('downloadLink');
        
        downloadLink.href = `/download/${progress.filename}`;
        downloadLink.download = progress.filename;
        downloadSection.style.display = 'block';

        // Rolar para a seção de download
        downloadSection.scrollIntoView({ behavior: 'smooth' });
    }

    onConversionError(error) {
        clearInterval(this.progressInterval);
        this.showError(error || 'Erro desconhecido na conversão');
    }

    showError(message) {
        const errorMessage = document.getElementById('errorMessage');
        errorMessage.textContent = message;
        errorMessage.style.display = 'block';
    }

    formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }
}

// Inicializar quando a página carregar
document.addEventListener('DOMContentLoaded', () => {
    new XLSBConverter();
});