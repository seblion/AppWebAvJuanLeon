# PoliMusic — Clon tipo Spotify (ASP clásico + SQL Server)

Sitio web tipo Spotify desarrollado con **ASP clásico (VBScript)**, **HTML5**, **CSS3** y
**JavaScript puro**, conectado a **SQL Server en localhost** usando exclusivamente la
estructura de la base de datos `BDD_PoliMusic` definida en `BDD_PoliMusic.sql`
(no se crearon tablas adicionales).

## 1. Requisitos

- Windows con **IIS** instalado y el rol **ASP** habilitado
  (Características de Windows → Internet Information Services → World Wide Web → ASP).
- **SQL Server** (cualquier edición, Express sirve) en `localhost` con modo de
  **autenticación mixta** (SQL Server + Windows).
- Proveedor OLE DB **SQL Server Native Client 11.0** (SQLNCLI11). Viene con SQL Server
  2012/2014/2016/2017/2019/2022. Si no está disponible, cambiar en
  `includes/conexion.asp` la constante `DB_PROVEEDOR` a `SQLOLEDB` (MDAC) o
  `MSOLEDBSQL` (cliente más reciente).

## 2. Crear la base de datos

1. Abrir **SQL Server Management Studio (SSMS)** y conectarse a `localhost`.
2. Abrir y **ejecutar** el script `BDD_PoliMusic.sql` (crea el login `usr_polimusic`
   con contraseña `Politecnica1`, la base `BDD_PoliMusic`, las 9 tablas, los datos
   iniciales y las llaves foráneas).
   - Nota: el archivo está codificado en UTF-16; SSMS lo abre sin problemas.
3. Credenciales usadas por el sitio (configuradas en `includes/conexion.asp`):

   | Parámetro     | Valor            |
   |---------------|------------------|
   | Servidor      | `localhost`      |
   | Base de datos | `BDD_PoliMusic`  |
   | Usuario       | `usr_polimusic`  |
   | Contraseña    | `Politecnica1`   |

4. Verificar que el protocolo **TCP/IP** esté habilitado en *SQL Server Configuration
   Manager → Configuración de red de SQL Server → Protocolos*.

## 3. Publicar el sitio en IIS

1. Copiar la carpeta del proyecto a `C:\inetpub\wwwroot\polimusic` (opcional).
2. En el **Administrador de IIS**, crear un nuevo **sitio** (o agregar una
   **aplicación** a Default Web Site) cuya ruta física apunte a la carpeta del
   proyecto, de modo que quede como **raíz de la aplicación** (el código usa rutas
   absolutas desde la raíz: `/css/styles.css`, `/index.asp`, etc.).
3. Grupo de aplicaciones: dejar el pipeline clásico o integrado, con .NET CLR en
   *Sin código administrado*; en **Comportamiento → Habilitar rutas principales**
   dejar activo.
4. Asegurar que la **identidad del grupo de aplicaciones** tenga permisos de lectura
   sobre la carpeta del proyecto (IIS_IUSRS / IUSR).
5. Abrir el navegador en `http://localhost/` (o el puerto configurado).

> `RUTA_BASE` se detecta automaticamente: si el sitio esta en la raiz de IIS usa
> `/`, y si esta un nivel debajo (p. ej. `http://localhost/webapppolimusicv1_asp/`)
> usa esa subcarpeta. Si se publica a mayor profundidad, escribir la ruta manualmente
> en `includes/conexion.asp`.

## 4. Uso del sitio

### Cuentas de prueba (incluidas en el script SQL)

| Usuario      | Contraseña  | Rol   | Puede hacer CRUD |
|--------------|-------------|-------|------------------|
| `userNormal` | `12345678`  | USER  | No               |
| `userAdmin`  | `12345678`  | ADMIN | Sí               |
| `username1`  | `12345678`  | USER  | No               |

El login acepta contraseñas guardadas en texto plano **o** en hash SHA1
(`HASHBYTES('SHA1', @pwd)` se evalúa en SQL Server), compatible con los usuarios
sembrados por el script.

### Funcionalidad

- **Inicio de sesión** (`login.asp`) validado contra `TBL_USER`.
- **Bienvenida personalizada** en la página principal usando la variable de sesión
  `Session("USERNAME")`.
- **Cierre de sesión** (`logout.asp`).
- **Listado de canciones ("Song List")** con portada, artista, álbum y contador de
  reproducciones; las canciones se reproducen desde el reproductor inferior
  (play/pausa, anterior/siguiente, barra de progreso, volumen, atajos de teclado).
  Cada reproducción incrementa `PLAYS` vía `api_reproducir.asp`.
- **Álbumes, artistas y listas de reproducción** del usuario.
- **Panel de administración** (solo rol admin): CRUD completo de
  **canciones**, **artistas** y **álbumes**, más gestión del tracklist de cada álbum
  (`admin/canciones.asp`, `admin/artistas.asp`, `admin/albumes.asp`).
- Restricción estricta: cualquier página de `admin/` verifica `USER_TYPE = 1` y
  redirige a `index.asp?error=permiso` si no corresponde.

## 5. Estructura del proyecto

```
webapppolimusicv1_asp/
├── BDD_PoliMusic.sql              # Script de base de datos (dado)
├── SongFilesAndImages_resources/  # MP3 e imágenes (dado)
│   ├── SongFiles/bensound-adventure.mp3
│   ├── PlaylistImages/…
│   └── UserImages/…
├── README.md
├── index.asp                      # Principal: bienvenida + Song List + reproductor
├── login.asp                      # Inicio de sesión
├── logout.asp                     # Cierre de sesión
├── api_reproducir.asp             # Contador de reproducciones (JSON)
├── css/styles.css                 # Paleta negro/azul, diseño propio
├── js/app.js                      # Utilidades generales
├── js/player.js                   # Reproductor (JS puro)
├── includes/
│   ├── conexion.asp               # Conexión ADO parametrizada + utilidades
│   ├── header.asp                 # Barra superior (consciente del rol)
│   ├── footer.asp                 # Pie + cierre de conexión
│   └── reproductor.asp            # Barra del reproductor
└── admin/
    ├── canciones.asp              # CRUD TBL_SONG
    ├── artistas.asp               # CRUD TBL_ARTIST
    └── albumes.asp                # CRUD TBL_ALBUM + tracklist
```

## 6. Notas de implementación

- **Consultas parametrizadas** (`ADODB.Command` + `CreateParameter`) en todos los
  accesos a datos; salida escapada con `EscaparHTML` para evitar XSS.
- **Rutas de recursos**: la BD guarda rutas como `~/songFiles/bensound-adventure.mp3`
  o `~/albumImages/album1.png`. La función `RutaRecurso` (en `includes/conexion.asp`)
  las traduce a la carpeta real `SongFilesAndImages_resources` y, si el archivo no
  existe, usa un reemplazo de la misma carpeta (p. ej. canciones sin MP3 disponible
  se reproducen con `bensound-adventure.mp3`, y álbumes/artistas usan las portadas
  `playlist1..5.png`).
- **Eliminaciones en cascada manuales**: al borrar una canción/artista/álbum se
  eliminan primero sus filas dependientes (`TBL_PLAYLIST_SONG`, `TBL_TRACKLIST`,
  `TBL_PLAYER`) dentro de una **transacción**, respetando las llaves foráneas sin
  crear tablas ni restricciones nuevas.
- **Sesión**: `Session.Abandon` al iniciar sesión para mitigar fijación de sesión.
- **Sin tablas adicionales**: toda la información proviene de las 9 tablas del script.

## 7. Solución de problemas

| Síntoma | Causa probable |
|---|---|
| Página de error de conexión | SQL Server detenido, TCP/IP deshabilitado, autenticación solo de Windows, o proveedor SQLNCLI11 ausente |
| HTTP 404.3 en IIS | El rol ASP no está instalado o no se registró (`dism /online /enable-feature ...` o reinstalar el rol) |
| Imágenes/audio no cargan | Modificar `RUTA_BASE` si el sitio no está en la raíz de IIS |
| Los cambios no se ven | Verificar que el sitio apunte a la carpeta correcta y recargar con Ctrl+F5 |

### Diagnostico rapido

1. Abrir **http://localhost/webapppolimusicv1_asp/hola.asp** (o la ruta correspondiente):
   pagina minima SIN includes ni base de datos que imprime por etapas sin buffer; la
   ultima linea visible indica donde falla el motor de ASP.
2. Abrir **http://localhost/webapppolimusicv1_asp/test_diagnostico.asp**: verifica ASP,
   sesion, FSO, proveedores ADO (SQLNCLI11 / MSOLEDBSQL / SQLOLEDB) y la conexion a
   `BDD_PoliMusic`, con el detalle de cada falla. Esta pagina no requiere sesion.
3. **paso2.asp**: intermedio entre hola.asp y test_diagnostico.asp; carga
   `includes/conexion.asp` y prueba la conexion directa a SQL Server.
2. Si el error persiste, revisar el **Visor de sucesos** (`eventvwr.msc`) → Registros de
   Windows → Aplicacion → fuente **Active Server Pages**: guarda el codigo exacto
   (ASP 0x800xxxxx) y la linea del error.
3. El archivo **web.config** incluido activa: errores detallados de ASP en el navegador
   (`scriptErrorSentToBrowser`), rutas principales para los `<!--#include file="../...">`
   (`enableParentPaths`), estado de sesion y documento por defecto `index.asp`.
   **En produccion** cambiar `scriptErrorSentToBrowser` a `false` y
   `errorMode="Detailed"` por `"DetailedLocalOnly"`.
