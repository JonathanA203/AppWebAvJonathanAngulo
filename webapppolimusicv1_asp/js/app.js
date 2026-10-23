/* =====================================================================
   PoliMusic · interacciones de la interfaz
   1) Tema claro / oscuro      2) Menú lateral móvil
   3) Confirmaciones            4) Filtro de canciones
   5) Reproductor de audio
   ===================================================================== */
(function () {
    'use strict';

    var DATA = window.POLIMUSIC || { playlist: [], hasPlayer: false, base: './' };
    var playlist = Array.isArray(DATA.playlist) ? DATA.playlist : [];
    var $ = function (sel, root) { return (root || document).querySelector(sel); };
    var $$ = function (sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); };

    /* ---------- 1. Tema ---------- */
    var THEME_KEY = 'polimusic-theme';

    function applyTheme(theme) {
        document.documentElement.setAttribute('data-theme', theme);
        try { localStorage.setItem(THEME_KEY, theme); } catch (e) {}
    }

    var btnTheme = $('#btnTheme');
    if (btnTheme) {
        btnTheme.addEventListener('click', function () {
            var current = document.documentElement.getAttribute('data-theme') || 'dark';
            applyTheme(current === 'dark' ? 'light' : 'dark');
        });
    }

    /* ---------- 2. Menú lateral móvil ---------- */
    var sidebar = $('#sidebar');
    var overlay = $('#overlay');
    var btnMenu = $('#btnMenu');

    function closeMenu() {
        if (sidebar) sidebar.classList.remove('is-open');
        if (overlay) overlay.hidden = true;
    }

    if (btnMenu && sidebar && overlay) {
        btnMenu.addEventListener('click', function () {
            sidebar.classList.add('is-open');
            overlay.hidden = false;
        });
        overlay.addEventListener('click', closeMenu);
        $$('.nav-item', sidebar).forEach(function (a) { a.addEventListener('click', closeMenu); });
        document.addEventListener('keydown', function (e) {
            if (e.key === 'Escape') closeMenu();
        });
    }

    /* ---------- 3. Confirmaciones ---------- */
    $$('form[data-confirm]').forEach(function (form) {
        form.addEventListener('submit', function (e) {
            var msg = form.getAttribute('data-confirm') || '¿Confirmas esta acción?';
            if (!window.confirm(msg)) {
                e.preventDefault();
            }
        });
    });

    /* ---------- 4. Filtro de canciones ---------- */
    var searchInput = $('#songSearch');
    var grid = $('#songGrid');
    var countPill = $('#songCount');
    var gridNote = $('#gridNote');

    function updateCount(visible) {
        if (countPill) {
            countPill.textContent = visible + (visible === 1 ? ' canción' : ' canciones');
        }
    }

    if (grid) {
        var cards = $$('.song-card', grid);

        if (searchInput) {
            searchInput.addEventListener('input', function () {
                var term = searchInput.value.trim().toLowerCase();
                var visible = 0;
                cards.forEach(function (card) {
                    var hay = (card.getAttribute('data-search') || '').toLowerCase();
                    var match = term === '' || hay.indexOf(term) !== -1;
                    card.style.display = match ? '' : 'none';
                    if (match) visible++;
                });
                updateCount(visible);
                if (gridNote) gridNote.hidden = visible !== 0;
            });
            updateCount(cards.length);
        }
    }

    /* ---------- 5. Reproductor ---------- */
    var audio = $('#audio');
    if (!audio || !DATA.hasPlayer) {
        // Sin reproductor en esta vista: solo se aplican las utilidades.
        autoDismissFlash();
        return;
    }

    var el = {
        cover: $('#plCover'),
        title: $('#plTitle'),
        artist: $('#plArtist'),
        play: $('#btnPlay'),
        prev: $('#btnPrev'),
        next: $('#btnNext'),
        seek: $('#plSeek'),
        current: $('#plCurrent'),
        duration: $('#plDuration'),
        repeat: $('#btnRepeat'),
        modeLabel: $('#plMode'),
        volume: $('#plVolume'),
        mute: $('#btnMute'),
        volumeWrap: $('.volume')
    };

    var MODES = ['list', 'all', 'one'];
    var MODE_LABEL = { list: 'Lista', all: 'Bucle', one: 'Una' };
    var MODE_TITLE = {
        list: 'Modo de reproducción: lista consecutiva (se detiene al final)',
        all: 'Modo de reproducción: bucle (repite toda la lista)',
        one: 'Modo de reproducción: repite la canción actual'
    };

    var state = {
        index: -1,
        mode: 'list',
        seeking: false,
        countedFor: -1
    };

    try {
        var savedMode = localStorage.getItem('polimusic-mode');
        if (savedMode && MODES.indexOf(savedMode) !== -1) state.mode = savedMode;
        var savedVol = localStorage.getItem('polimusic-volume');
        if (savedVol !== null) audio.volume = Math.min(1, Math.max(0, parseFloat(savedVol)));
        else audio.volume = 0.8;
    } catch (e) {
        audio.volume = 0.8;
    }

    if (el.volume) el.volume.value = String(audio.volume);

    function fmtTime(sec) {
        if (!isFinite(sec) || sec < 0) return '0:00';
        var m = Math.floor(sec / 60);
        var s = Math.floor(sec % 60);
        return m + ':' + (s < 10 ? '0' + s : s);
    }

    function currentTrack() {
        return (state.index >= 0 && state.index < playlist.length) ? playlist[state.index] : null;
    }

    function updateModeUI() {
        if (!el.repeat) return;
        el.repeat.setAttribute('data-mode', state.mode);
        el.repeat.title = MODE_TITLE[state.mode];
        if (el.modeLabel) el.modeLabel.textContent = MODE_LABEL[state.mode];
        try { localStorage.setItem('polimusic-mode', state.mode); } catch (e) {}
    }

    function highlightCards() {
        $$('.song-card').forEach(function (card) {
            var idx = parseInt(card.getAttribute('data-index'), 10);
            var isCurrent = idx === state.index;
            card.classList.toggle('is-active', isCurrent);
            card.classList.toggle('is-playing', isCurrent && !audio.paused);
        });
    }

    function syncPlayerUI() {
        var track = currentTrack();
        if (!track) {
            if (el.title) el.title.textContent = 'Sin reproducción';
            if (el.artist) el.artist.textContent = 'Elige una canción del listado';
            if (el.current) el.current.textContent = '0:00';
            if (el.duration) el.duration.textContent = '0:00';
            if (el.seek) el.seek.value = '0';
            return;
        }
        if (el.cover) el.cover.src = track.cover;
        if (el.title) el.title.textContent = track.title;
        if (el.artist) el.artist.textContent = track.artist;
        if (el.duration && isFinite(audio.duration)) el.duration.textContent = fmtTime(audio.duration);
    }

    function updatePlayState() {
        document.body.classList.toggle('is-playing', !audio.paused);
        if (el.play) {
            el.play.title = audio.paused ? 'Reproducir' : 'Pausar';
            el.play.setAttribute('aria-label', audio.paused ? 'Reproducir' : 'Pausar');
        }
        var t = currentTrack();
        if (t) document.title = (audio.paused ? '' : '▶ ') + t.title + ' · PoliMusic';
        highlightCards();
    }

    function reportPlay(track) {
        if (!track || state.countedFor === state.index) return;
        state.countedFor = state.index;
        fetch(DATA.base + 'ajax_play.asp?id=' + encodeURIComponent(track.id), { credentials: 'same-origin' })
            .then(function (r) { return r.json(); })
            .then(function (res) {
                if (!res || !res.ok) return;
                track.plays = res.plays;
                var card = $('.song-card[data-index="' + state.index + '"] .plays-pill');
                if (card) {
                    var label = res.plays >= 1000
                        ? (Math.round(res.plays / 100) / 10).toString().replace('.', ',') + 'K'
                        : String(res.plays);
                    card.innerHTML = '<svg><use href="#i-play"></use></svg>' + label;
                }
            })
            .catch(function () { /* el contador es opcional */ });
    }

    function playAt(index, autoplay) {
        if (index < 0 || index >= playlist.length) return;
        state.index = index;
        var track = playlist[index];
        audio.src = track.src;
        syncPlayerUI();

        if (autoplay !== false) {
            var p = audio.play();
            if (p && p.catch) p.catch(function () { updatePlayState(); });
        }
        if ('mediaSession' in navigator && window.MediaMetadata) {
            try {
                navigator.mediaSession.metadata = new MediaMetadata({
                    title: track.title,
                    artist: track.artist,
                    album: track.album || 'PoliMusic',
                    artwork: [{ src: track.cover, sizes: '512x512', type: 'image/jpeg' }]
                });
            } catch (e) {}
        }
    }

    function togglePlay() {
        if (state.index < 0) {
            if (playlist.length > 0) playAt(0, true);
            return;
        }
        if (audio.paused) {
            var p = audio.play();
            if (p && p.catch) p.catch(function () {});
        } else {
            audio.pause();
        }
    }

    function nextTrack(auto) {
        if (playlist.length === 0) return;
        if (state.index < 0) { playAt(0, true); return; }

        if (state.index + 1 < playlist.length) {
            playAt(state.index + 1, true);
            return;
        }
        if (state.mode === 'all') {
            playAt(0, true);
        } else if (!auto) {
            playAt(0, true);
        } else {
            audio.pause();
            audio.currentTime = 0;
            updatePlayState();
        }
    }

    function prevTrack() {
        if (playlist.length === 0) return;
        if (audio.currentTime > 3 || state.index <= 0) {
            audio.currentTime = 0;
            if (audio.paused) togglePlay();
            return;
        }
        playAt(state.index - 1, true);
    }

    /* Controles principales */
    if (el.play) el.play.addEventListener('click', togglePlay);
    if (el.next) el.next.addEventListener('click', function () { nextTrack(false); });
    if (el.prev) el.prev.addEventListener('click', prevTrack);

    if (el.repeat) {
        el.repeat.addEventListener('click', function () {
            var i = MODES.indexOf(state.mode);
            state.mode = MODES[(i + 1) % MODES.length];
            updateModeUI();
        });
    }

    /* Progreso */
    audio.addEventListener('timeupdate', function () {
        if (state.seeking) return;
        if (el.current) el.current.textContent = fmtTime(audio.currentTime);
        if (el.seek && isFinite(audio.duration) && audio.duration > 0) {
            el.seek.value = String((audio.currentTime / audio.duration) * 100);
        }
    });

    audio.addEventListener('loadedmetadata', function () {
        if (el.duration) el.duration.textContent = fmtTime(audio.duration);
    });

    audio.addEventListener('play', function () {
        updatePlayState();
        reportPlay(currentTrack());
    });

    audio.addEventListener('pause', updatePlayState);

    audio.addEventListener('ended', function () {
        if (state.mode === 'one') {
            audio.currentTime = 0;
            audio.play();
            return;
        }
        nextTrack(true);
    });

    audio.addEventListener('error', function () {
        if (el.artist) el.artist.textContent = 'No se pudo cargar el archivo de audio';
    });

    if (el.seek) {
        el.seek.addEventListener('input', function () {
            state.seeking = true;
            if (isFinite(audio.duration)) {
                if (el.current) el.current.textContent = fmtTime((parseFloat(el.seek.value) / 100) * audio.duration);
            }
        });
        el.seek.addEventListener('change', function () {
            if (isFinite(audio.duration)) {
                audio.currentTime = (parseFloat(el.seek.value) / 100) * audio.duration;
            }
            state.seeking = false;
        });
    }

    /* Volumen */
    if (el.volume) {
        el.volume.addEventListener('input', function () {
            audio.volume = Math.min(1, Math.max(0, parseFloat(el.volume.value)));
            audio.muted = audio.volume === 0;
            if (el.volumeWrap) el.volumeWrap.classList.toggle('is-muted', audio.muted);
            try { localStorage.setItem('polimusic-volume', String(audio.volume)); } catch (e) {}
        });
    }

    if (el.mute) {
        el.mute.addEventListener('click', function () {
            audio.muted = !audio.muted;
            if (el.volumeWrap) el.volumeWrap.classList.toggle('is-muted', audio.muted);
            if (el.volume) el.volume.value = audio.muted ? '0' : String(audio.volume || 0.8);
        });
    }

    /* Tarjetas de canción */
    $$('.song-card').forEach(function (card) {
        var idx = parseInt(card.getAttribute('data-index'), 10);

        card.addEventListener('click', function (e) {
            if (e.target.closest('.song-admin') || e.target.closest('a') || e.target.closest('form')) return;
            if (idx === state.index) { togglePlay(); return; }
            playAt(idx, true);
        });

        card.addEventListener('keydown', function (e) {
            if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                if (idx === state.index) { togglePlay(); } else { playAt(idx, true); }
            }
        });
    });

    /* Reproducir todo */
    ['#btnPlayAll', '#btnPlayAllTop'].forEach(function (sel) {
        var btn = $(sel);
        if (btn) {
            btn.addEventListener('click', function () {
                if (playlist.length === 0) return;
                if (state.index === 0 && !audio.paused) { togglePlay(); return; }
                playAt(0, true);
            });
        }
    });

    /* Atajos de teclado */
    document.addEventListener('keydown', function (e) {
        var tag = (e.target.tagName || '').toLowerCase();
        if (tag === 'input' || tag === 'textarea' || tag === 'select' || e.target.isContentEditable) return;

        if (e.code === 'Space') {
            e.preventDefault();
            togglePlay();
        } else if (e.code === 'ArrowRight' && e.shiftKey) {
            nextTrack(false);
        } else if (e.code === 'ArrowLeft' && e.shiftKey) {
            prevTrack();
        } else if (e.code === 'KeyM') {
            if (el.mute) el.mute.click();
        }
    });

    if ('mediaSession' in navigator) {
        try {
            navigator.mediaSession.setActionHandler('play', function () { audio.play(); });
            navigator.mediaSession.setActionHandler('pause', function () { audio.pause(); });
            navigator.mediaSession.setActionHandler('nexttrack', function () { nextTrack(false); });
            navigator.mediaSession.setActionHandler('previoustrack', prevTrack);
        } catch (e) {}
    }

    updateModeUI();
    updatePlayState();
    syncPlayerUI();
    autoDismissFlash();

    /* La barra arranca con la primera canción cargada (sin reproducir) */
    if (playlist.length > 0 && state.index < 0) {
        state.index = 0;
        audio.src = playlist[0].src;
        syncPlayerUI();
        highlightCards();
    }

    /* ---------- Utilidades ---------- */
    function autoDismissFlash() {
        $$('.flash').forEach(function (flash) {
            setTimeout(function () {
                flash.style.transition = 'opacity .4s ease, transform .4s ease';
                flash.style.opacity = '0';
                flash.style.transform = 'translateY(-6px)';
                setTimeout(function () { flash.remove(); }, 420);
            }, 6500);
        });
    }
})();
