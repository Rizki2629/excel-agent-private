// ==========================================================================
// UI TAB v1.7.0 "Chat-first" (Okt 2026)
// Chat jadi elemen utama; fitur dikelompokkan: Chat | Cepat | Gambar | Pengaturan.
// File terpisah — tidak mengubah satu baris pun taskpane.js.
// Sintaks ES5 (kompatibel IE11).
// ==========================================================================

var EA_TABS = ['chat', 'cepat', 'gambar', 'pengaturan'];

// Pindah tab aktif. Dipanggil dari klik tab maupun aksi (kartu cepat, lampir gambar).
function eaSwitchTab(name) {
    var i, panel, btn;
    for (i = 0; i < EA_TABS.length; i++) {
        panel = document.getElementById('tab-' + EA_TABS[i]);
        if (panel) {
            panel.className = 'tab-panel' + (EA_TABS[i] === name ? ' active' : '');
        }
        btn = document.getElementById('tabbtn-' + EA_TABS[i]);
        if (btn) {
            btn.className = 'tab-btn' + (EA_TABS[i] === name ? ' active' : '');
        }
    }
    if (name === 'pengaturan') {
        eaRefreshSettingsTab();
    }
}

// Isi ringkasan provider/model di tab Pengaturan dari objek settings global.
function eaRefreshSettingsTab() {
    var provEl = document.getElementById('setProviderName');
    var modelEl = document.getElementById('setModelName');
    try {
        if (provEl) {
            provEl.textContent = (typeof settings !== 'undefined' && settings.activeProvider)
                ? settings.activeProvider
                : '\u2014 Manual \u2014';
        }
        if (modelEl) {
            modelEl.textContent = (typeof settings !== 'undefined' && settings.model)
                ? settings.model
                : '\u2014 belum diatur \u2014';
        }
    } catch (e) { /* tab boleh kosong bila settings belum termuat */ }
}

function eaSetupTabs() {
    var i;

    // Klik pada setiap tombol tab
    for (i = 0; i < EA_TABS.length; i++) {
        (function (tabName) {
            var btn = document.getElementById('tabbtn-' + tabName);
            if (btn) {
                btn.onclick = function () { eaSwitchTab(tabName); };
            }
        })(EA_TABS[i]);
    }

    // Tab Gambar: tombol dropzone memicu input file yang sama dengan tab Chat,
    // lalu kembali ke Chat supaya thumbnail pratinjau terlihat di input.
    var gambarBtn = document.getElementById('gambarAttachBtn');
    if (gambarBtn) {
        gambarBtn.onclick = function () {
            if (typeof bindingBlocked !== 'undefined' && bindingBlocked) return;
            if (typeof workbookBinding === 'undefined' || !workbookBinding) return;
            var fileInput = document.getElementById('imageInput');
            if (fileInput) fileInput.click();
            eaSwitchTab('chat');
        };
    }

    // Tab Pengaturan: tombol "Buka Pengaturan Lengkap" memicu tombol settings
    // asli yang disembunyikan, sehingga modal & binding lama tetap dipakai.
    var openSettingsBtn = document.getElementById('openSettingsFullBtn');
    if (openSettingsBtn) {
        openSettingsBtn.onclick = function () {
            var hiddenBtn = document.getElementById('settingsBtn');
            if (hiddenBtn) hiddenBtn.click();
        };
    }

    // Kartu aksi di tab Cepat: setelah prompt dikirim (handler lama),
    // pindahkan tampilan ke tab Chat agar user melihat prosesnya.
    // Dipasang di fase bubble pada panel, jadi jalan SETELAH onclick kartu.
    var cepatPanel = document.getElementById('tab-cepat');
    if (cepatPanel && cepatPanel.addEventListener) {
        cepatPanel.addEventListener('click', function (e) {
            var t = e.target || e.srcElement;
            while (t && t !== cepatPanel) {
                var cls = t.className || '';
                if (typeof cls === 'string' && cls.indexOf('example-card') !== -1) {
                    eaSwitchTab('chat');
                    break;
                }
                t = t.parentNode;
            }
        });
    }
}

// Script dimuat di akhir <body> setelah taskpane.js, sehingga seluruh elemen tab sudah ada di DOM.
eaSetupTabs();
