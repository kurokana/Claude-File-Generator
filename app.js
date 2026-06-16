/* ================================================================
   DOCX Generator — Application Logic
   Sandbox execution of user JS scripts that use the docx library.
   ================================================================ */

document.addEventListener('DOMContentLoaded', () => {
    // ─── DOM Elements ─────────────────────────────────────────────
    const codeEditor = document.getElementById('code-editor');
    const btnClearCode = document.getElementById('btn-clear-code');
    const codeLineCount = document.getElementById('code-line-count');
    const btnGenerate = document.getElementById('btn-generate');
    const btnText = document.getElementById('btn-text');
    const btnLoading = document.getElementById('btn-loading');
    const statusBadge = document.getElementById('status-badge');
    const statusIcon = document.getElementById('status-icon');
    const statusText = document.getElementById('status-text');
    const consoleOutput = document.getElementById('console-output');
    const btnClearConsole = document.getElementById('btn-clear-console');

    // ─── Input Code Handlers ──────────────────────────────────────
    function updateCodeMetrics() {
        const code = codeEditor.value;
        const lines = code ? code.split('\n').length : 0;
        codeLineCount.textContent = `${lines} baris`;
        
        // Enable generate button if there is text content
        const hasContent = code.trim().length > 0;
        btnGenerate.disabled = !hasContent;
    }

    codeEditor.addEventListener('input', () => {
        updateCodeMetrics();
        hideStatus();
    });

    btnClearCode.addEventListener('click', () => {
        codeEditor.value = '';
        updateCodeMetrics();
        hideStatus();
        logConsole('Input kode JS dibersihkan.', 'system');
    });

    // ─── Generate DOCX ───────────────────────────────────────────
    btnGenerate.addEventListener('click', async () => {
        const code = codeEditor.value;
        if (!code.trim()) return;

        // Check docx library
        if (!window.docx) {
            logConsole('Library docx belum dimuat. Periksa koneksi internet dan refresh halaman.', 'error');
            showStatus('error', '⚠️ Library docx belum tersedia');
            return;
        }

        // UI: loading state
        btnGenerate.disabled = true;
        btnText.classList.add('hidden');
        btnLoading.classList.remove('hidden');
        hideStatus();

        logConsole('Mulai memproses script...', 'system');

        try {
            await executeDocxScript(code, 'editor.js');
        } catch (err) {
            logConsole(`Error: ${err.message}`, 'error');
            showStatus('error', `❌ Gagal: ${err.message}`);
        } finally {
            // UI: restore button
            btnText.classList.remove('hidden');
            btnLoading.classList.add('hidden');
            btnGenerate.disabled = false;
        }
    });

    // ─── Script Executor (Sandbox) ───────────────────────────────
    async function executeDocxScript(code, sourceFileName) {
        // Track download completion
        let downloadTriggered = false;
        let downloadFileName = 'output.docx';
        let downloadResolve;
        const downloadPromise = new Promise((resolve) => {
            downloadResolve = resolve;
        });

        // Timeout (30 seconds)
        const timeoutPromise = new Promise((_, reject) => {
            setTimeout(() => reject(new Error('Script timeout (30 detik). Pastikan script memanggil Packer.toBuffer() dan fs.writeFileSync().')), 30000);
        });

        // ─── Mock: console ───────────────────────────────────────
        const mockConsole = {
            log: (...args) => logConsole(args.map(String).join(' '), 'info'),
            error: (...args) => logConsole('ERROR: ' + args.map(String).join(' '), 'error'),
            warn: (...args) => logConsole('WARN: ' + args.map(String).join(' '), 'warn'),
            info: (...args) => logConsole(args.map(String).join(' '), 'info'),
            debug: (...args) => logConsole('DEBUG: ' + args.map(String).join(' '), 'info'),
            table: (...args) => logConsole('[table] ' + args.map(a => JSON.stringify(a)).join(' '), 'info'),
            dir: (...args) => logConsole(args.map(a => JSON.stringify(a, null, 2)).join(' '), 'info'),
            time: () => {},
            timeEnd: () => {},
            trace: () => {},
            assert: () => {},
            clear: () => {},
            count: () => {},
            countReset: () => {},
            group: () => {},
            groupEnd: () => {},
            groupCollapsed: () => {},
        };

        // ─── Mock: fs ────────────────────────────────────────────
        const mockFs = {
            writeFileSync: (name, data) => {
                downloadFileName = name;
                downloadTriggered = true;
                logConsole(`fs.writeFileSync("${name}") dipanggil — memulai download...`, 'system');
                triggerDownload(name, data);
                downloadResolve(name);
            },
            readFileSync: () => {
                throw new Error('fs.readFileSync tidak tersedia di browser');
            },
            existsSync: () => false,
            mkdirSync: () => {},
            readdirSync: () => [],
        };

        // ─── Mock: require ───────────────────────────────────────
        const mockRequire = (moduleName) => {
            if (moduleName === 'docx') {
                logConsole('require("docx") → library docx dimuat', 'system');
                
                // Create a patched version of docx where Packer.toBuffer works in browser
                // Packer.toBuffer is Node.js only; we redirect to Packer.toBlob + conversion
                const patchedDocx = Object.create(null);
                for (const key of Object.keys(window.docx)) {
                    patchedDocx[key] = window.docx[key];
                }
                
                // Patch Packer with browser-compatible toBuffer
                patchedDocx.Packer = {
                    toBlob: (doc, options) => window.docx.Packer.toBlob(doc, options),
                    toBase64String: (doc, options) => window.docx.Packer.toBase64String(doc, options),
                    toBuffer: async (doc, options) => {
                        logConsole('Packer.toBuffer() → menggunakan Packer.toBlob() (browser mode)', 'system');
                        const blob = await window.docx.Packer.toBlob(doc, options);
                        const arrayBuffer = await blob.arrayBuffer();
                        return new Uint8Array(arrayBuffer);
                    }
                };
                
                return patchedDocx;
            }
            if (moduleName === 'fs') {
                return mockFs;
            }
            if (moduleName === 'path') {
                return {
                    join: (...args) => args.join('/'),
                    resolve: (...args) => args.join('/'),
                    basename: (p) => p.split('/').pop() || p.split('\\').pop(),
                    dirname: (p) => p.substring(0, p.lastIndexOf('/')),
                    extname: (p) => { const m = p.match(/\.[^.]+$/); return m ? m[0] : ''; },
                };
            }
            throw new Error(`Module "${moduleName}" tidak tersedia di browser. Hanya "docx", "fs", dan "path" yang didukung.`);
        };

        // ─── Execute ─────────────────────────────────────────────
        logConsole('Mengeksekusi script...', 'system');

        try {
            // Use AsyncFunction to support top-level await
            const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
            const fn = new AsyncFunction(
                'require', 'module', 'exports', 'console',
                '__filename', '__dirname', 'process', 'Buffer', 'setTimeout', 'setInterval', 'clearTimeout', 'clearInterval',
                code
            );

            const mockModule = { exports: {} };
            const mockProcess = {
                env: {},
                cwd: () => '.',
                exit: () => { logConsole('process.exit() dipanggil (diabaikan di browser)', 'warn'); },
                argv: ['node', sourceFileName],
                platform: 'browser',
                version: 'v18.0.0',
            };

            await fn(
                mockRequire,
                mockModule,
                mockModule.exports,
                mockConsole,
                sourceFileName,
                '.',
                mockProcess,
                {
                    from: (data, encoding) => {
                        if (typeof data === 'string') {
                            return new TextEncoder().encode(data);
                        }
                        return new Uint8Array(data);
                    },
                    isBuffer: (obj) => obj instanceof Uint8Array,
                    alloc: (size) => new Uint8Array(size),
                },
                window.setTimeout.bind(window),
                window.setInterval.bind(window),
                window.clearTimeout.bind(window),
                window.clearInterval.bind(window)
            );

            // Wait for async operations to complete (Packer.toBuffer().then(...))
            logConsole('Menunggu operasi async selesai...', 'system');

            await Promise.race([downloadPromise, timeoutPromise]);

            if (downloadTriggered) {
                logConsole(`✅ File "${downloadFileName}" berhasil di-generate dan didownload!`, 'success');
                showStatus('success', `✅ File "${downloadFileName}" berhasil didownload!`);

                // Add success pulse animation to generate card
                const genCard = document.getElementById('generate-card');
                genCard.classList.add('success-pulse');
                setTimeout(() => genCard.classList.remove('success-pulse'), 1500);
            }
        } catch (err) {
            if (err.message.includes('timeout')) {
                throw err;
            }
            throw new Error(`Script execution error: ${err.message}`);
        }
    }

    // ─── Download Trigger ─────────────────────────────────────────
    function triggerDownload(filename, data) {
        try {
            const blob = new Blob([data], {
                type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            a.style.display = 'none';
            document.body.appendChild(a);
            a.click();

            // Cleanup
            setTimeout(() => {
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
            }, 1000);
        } catch (err) {
            logConsole(`Download error: ${err.message}`, 'error');
        }
    }

    // ─── Console Logger ───────────────────────────────────────────
    function logConsole(message, type = 'info') {
        const line = document.createElement('div');
        line.className = `console-line console-${type}`;

        const time = document.createElement('span');
        time.className = 'console-time';
        const now = new Date();
        time.textContent = `${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;

        const text = document.createElement('span');
        text.textContent = message;

        line.appendChild(time);
        line.appendChild(text);
        consoleOutput.appendChild(line);

        // Auto-scroll to bottom
        consoleOutput.scrollTop = consoleOutput.scrollHeight;
    }

    // ─── Status Badge ─────────────────────────────────────────────
    function showStatus(type, message) {
        statusBadge.className = `status-badge ${type}`;
        statusBadge.classList.remove('hidden');

        const icons = { success: '✓', error: '✕', processing: '⟳' };
        statusIcon.textContent = '';
        statusText.textContent = message;
    }

    function hideStatus() {
        statusBadge.classList.add('hidden');
    }

    // ─── Clear Console ────────────────────────────────────────────
    btnClearConsole.addEventListener('click', () => {
        consoleOutput.innerHTML = '';
        logConsole('Console dibersihkan.', 'system');
    });

    // ─── Utilities ────────────────────────────────────────────────
    function formatFileSize(bytes) {
        if (bytes < 1024) return bytes + ' B';
        if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
        return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
    }

    function pad(n) {
        return n.toString().padStart(2, '0');
    }
});
