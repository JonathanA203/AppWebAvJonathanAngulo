<%
'=====================================================================
'  PoliMusic :: pie comun + barra de reproduccion
'=====================================================================
%>
        </main>
    </div>
</div>

<div class="overlay" id="overlay" hidden></div>

<% If pageId = "home" Then %>
<footer class="player" id="player" aria-label="Reproductor de música">
    <div class="player-track">
        <div class="player-cover" id="plCoverWrap">
            <img id="plCover" src="<%=strBase%>SongFilesAndImages_resources/PlaylistImages/playlist1.png" alt="Portada de la canción">
            <span class="player-cover-glow" aria-hidden="true"></span>
        </div>
        <div class="player-meta">
            <strong id="plTitle">Sin reproducción</strong>
            <span id="plArtist">Elige una canción del listado</span>
        </div>
        <div class="eq" id="plEq" aria-hidden="true"><i></i><i></i><i></i><i></i></div>
    </div>

    <div class="player-main">
        <div class="player-controls">
            <button type="button" class="icon-btn" id="btnPrev" title="Canción anterior" aria-label="Canción anterior">
                <svg><use href="#i-prev"></use></svg>
            </button>
            <button type="button" class="play-btn" id="btnPlay" title="Reproducir" aria-label="Reproducir">
                <svg class="ico-play"><use href="#i-play"></use></svg>
                <svg class="ico-pause"><use href="#i-pause"></use></svg>
            </button>
            <button type="button" class="icon-btn" id="btnNext" title="Canción siguiente" aria-label="Canción siguiente">
                <svg><use href="#i-next"></use></svg>
            </button>
        </div>
        <div class="player-progress">
            <span id="plCurrent">0:00</span>
            <input type="range" id="plSeek" class="range" min="0" max="100" value="0" step="0.1"
                   aria-label="Avanzar en la canción">
            <span id="plDuration">0:00</span>
        </div>
    </div>

    <div class="player-right">
        <button type="button" class="mode-btn" id="btnRepeat" data-mode="list"
                title="Modo de reproducción: lista consecutiva">
            <svg class="ico-repeat"><use href="#i-repeat"></use></svg>
            <svg class="ico-repeat-one"><use href="#i-repeat-one"></use></svg>
            <span class="mode-label" id="plMode">Lista</span>
        </button>
        <div class="volume">
            <button type="button" class="icon-btn" id="btnMute" title="Silenciar" aria-label="Silenciar">
                <svg class="ico-vol"><use href="#i-volume"></use></svg>
                <svg class="ico-mute"><use href="#i-mute"></use></svg>
            </button>
            <input type="range" id="plVolume" class="range range-vol" min="0" max="1" step="0.01" value="0.8"
                   aria-label="Volumen">
        </div>
    </div>

    <audio id="audio" preload="metadata"></audio>
</footer>
<% End If %>

<script>
window.POLIMUSIC = {
    base: "<%=strBase%>",
    playlist: <%=strPlaylistJson%>,
    hasPlayer: <%=IIf(pageId = "home", "true", "false")%>
};
</script>
<script src="<%=strBase%>js/app.js"></script>
</body>
</html>
<% Call CloseConn() %>
