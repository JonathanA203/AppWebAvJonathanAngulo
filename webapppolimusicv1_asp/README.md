# PoliMusic · WebApp en ASP clásico (VBScript)

Aplicación web de música inspirada en la disposición de Spotify, con identidad visual
propia en tonos azules, dos roles de usuario y un reproductor completo.

Todo el proyecto está hecho con **ASP clásico + VBScript** (sin ASP.NET, sin C#):
todas las vistas son archivos `.asp`.

---

## 1. Requisitos

| Componente | Detalle |
|---|---|
| Servidor web | IIS o IIS Express con la característica **ASP** (ASP clásico) habilitada |
| Base de datos | SQL Server con la base `BDD_PoliMusic` creada desde `BDD_PoliMusic.sql` |
| Navegador | Chrome, Edge o Firefox actualizados |

La base de datos **no se modifica**: se usan las tablas y relaciones ya existentes
(`TBL_USER`, `TBL_SONG`, `TBL_ARTIST`, `TBL_ALBUM`, `TBL_TRACKLIST`, `TBL_PLAYER`,
`TBL_PLAYLIST`, `TBL_PLAYLIST_SONG`, `TBL_CATALOG`).

## 2. Puesta en marcha

1. Ejecutar `BDD_PoliMusic.sql` en SQL Server (crea la base, el login `usr_polimusic` y los datos de ejemplo).
2. Publicar la carpeta `webapppolimusicv1_asp` como sitio o aplicación virtual.
   - **IIS Express (rápido):**
     ```bat
     "C:\Program Files\IIS Express\iisexpress.exe" /path:"C:\ruta\webapppolimusicv1_asp" /port:8099
     ```
   - **IIS:** crear un sitio/apartado con esa carpeta como ruta física y un grupo de
     aplicaciones con .NET CLR habilitado (el ASP clásico no requiere código administrado).
3. Abrir `http://localhost:8099/` → redirige a `login.asp`.
4. Si tu servidor SQL no es `localhost` o cambian las credenciales, edita
   **`includes/config.asp`**:

   ```vbscript
   Const DB_SERVER = "localhost"
   Const DB_NAME   = "BDD_PoliMusic"
   Const DB_USER   = "usr_polimusic"
   Const DB_PASS   = "Politecnica1"
   ```

   La cadena usa `Provider=SQLOLEDB`. Si el equipo solo tiene el driver moderno,
   cambia el proveedor por `MSOLEDBSQL` en la función `ConnString()`.

### Cuentas de prueba

| Usuario | Contraseña | Rol |
|---|---|---|
| `userAdmin` | `12345678` | Administrador (CRUD habilitado) |
| `userNormal` | `12345678` | Usuario normal (solo escucha) |
| `username1` | `12345678` | Usuario normal, contraseña guardada en SHA-1 |

El login acepta contraseñas en texto plano **y** en SHA-1 (la tabla mezcla ambos formatos).
Las contraseñas nuevas creadas desde el CRUD se guardan cifradas con SHA-1.

## 3. Vistas

| Archivo | Descripción |
|---|---|
| `login.asp` | Inicio de sesión. Guarda `Session("UserID")`, `Session("Username")`, `Session("UserType")`, `Session("UserPhoto")`. |
| `home.asp` | Vista principal: mensaje de bienvenida con el usuario de la sesión, **Song List** y playlists. |
| `admin.asp` | Panel del administrador con métricas y accesos rápidos. |
| `admin_songs.asp` / `admin_song_form.asp` | CRUD de canciones (incluye álbum y artistas). |
| `admin_artists.asp` / `admin_artist_form.asp` | CRUD de artistas. |
| `admin_albums.asp` / `admin_album_form.asp` | CRUD de álbumes. |
| `admin_playlists.asp` / `admin_playlist_form.asp` | CRUD de playlists. |
| `admin_users.asp` / `admin_user_form.asp` | CRUD de usuarios y roles. |
| `logout.asp` | Cierre de sesión (`Session.Abandon`). |
| `ajax_play.asp` | Suma una reproducción a `TBL_SONG.PLAYS` (lo llama el reproductor). |
| `index.asp` | Punto de entrada: redirige a `home.asp` o `login.asp`. |

Carpetas de apoyo:

```
includes/   config.asp (constantes y utilidades), db.asp (ADO), sha1.asp,
            auth.asp (sesión y roles), header.asp, footer.asp, sprite.asp (iconos SVG)
css/        styles.css  (diseño azul, modo claro y oscuro)
js/         app.js      (tema, filtro, reproductor)
SongFilesAndImages_resources/   audios e imágenes originales del proyecto
```

## 4. Funcionalidades

**Roles.** `USER_TYPE = 0` es usuario normal y `USER_TYPE = 1` es administrador.
`includes/auth.asp` expone `RequireLogin()` y `RequireAdmin()`; las páginas de
administración redirigen a `home.asp` si la sesión no es de tipo admin. La barra
lateral y los botones de edición/eliminación solo se muestran al administrador.

**Song List.** La sección incluye el texto introductorio, el contador de canciones y
tarjetas con portada, título, artistas, álbum, año y reproducciones. El buscador
superior filtra en el servidor (`LIKE`) y en el cliente mientras se escribe.

**Reproductor (barra inferior).**

- Play / pausa, canción anterior y siguiente.
- Modo de reproducción conmutables con el botón de la derecha:
  - **Lista** → lista consecutiva: al terminar la última canción se detiene.
  - **Bucle** → repite toda la lista.
  - **Una** → repite la canción actual.
- Barra de progreso con búsqueda, volumen, silencio y atajos de teclado
  (`Espacio`, `Shift + ←/→`, `M`).
- Cada reproducción incrementa `PLAYS` en la base de datos vía `ajax_play.asp`.
- El modo, el volumen y el tema se recuerdan en `localStorage`.

**Tema claro / oscuro.** El botón de la barra superior alterna `data-theme` en `<html>`;
los colores viven en variables CSS, así que ambos modos comparten la misma estructura,
tipografía y acentos azules.

**Diseño.** Paleta azul (`#3b82f6` / `#22d3ee` sobre `#060d1c`), superficies con
bordes suaves, degradados, tarjetas con elevación al pasar el cursor y barra inferior
con efecto glass. Totalmente responsive (barra lateral convertida en menú deslizante
en pantallas pequeñas).

## 5. Notas sobre los archivos multimedia

La base de datos referencia rutas como `~/songFiles/bensound-dreams.mp3`, pero la
carpeta de recursos solo contiene tres audios. `includes/config.asp` resuelve cada
canción en este orden:

1. Coincidencia exacta del nombre del archivo.
2. Coincidencia por nombre normalizado (ignora espacios y guiones), lo que permite
   que `Loser-TameImpala.mp3` encuentre a `Loser-Tame Impala.mp3`.
3. Asignación cíclica entre los audios disponibles, para que ninguna canción quede muda.

Las portadas se resuelven igual: si la imagen del álbum no existe, se usan
`play-back-in-black.jpg`, `play-loser.jpg` o las portadas `playlist1..5.png`.

## 6. Seguridad y buenas prácticas

- Todas las consultas usan **parámetros ADO** (`?`), no concatenación de SQL.
- Las salidas de base de datos pasan por `Server.HTMLEncode` (`H()`), evitando XSS.
- Las eliminaciones se envían por **POST** y piden confirmación.
- Las operaciones con varias tablas (canción + artistas + álbum) usan transacción.
- Los borrados que violan claves foráneas se informan con un mensaje claro en lugar
  de romper la página.
- Un administrador no puede eliminarse a sí mismo ni quitarse el rol.

## 7. Capturas de referencia

En `docs/screenshots/` están las vistas del resultado final:

| Archivo | Vista |
|---|---|
| `login.png` | Pantalla de inicio de sesión |
| `home-oscuro.png` | Song List + reproductor en modo oscuro |
| `home-claro.png` | La misma vista en modo claro |
| `admin-canciones.png` | CRUD de canciones (rol administrador) |

