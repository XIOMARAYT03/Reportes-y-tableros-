# Kardex Fimas - Guia de despliegue (Windows Server)

Aplicacion Python (solo libreria estandar + `openpyxl`) que sirve el panel Kardex
y se conecta al ERP (Odoo) por XML-RPC. Se instala como **tarea programada** que
arranca al iniciar el servidor, sin ventanas y con reinicio automatico ante fallos.
No requiere software de terceros.

---

## 1. Contenido del paquete

```
Kardex-ERP\
  app\       Aplicacion: kardex_server.py, *.py, *.html, requirements.txt, .env.example
  deploy\    install.ps1, uninstall.ps1, actualizar.ps1, README_DEPLOY.md
```

> Las bases de datos `permisos.db` y `usuarios.db` **no** van en el paquete por
> defecto: el servidor las crea vacias en el primer arranque. Si quieres conservar
> usuarios y permisos, genera el paquete con `pack.ps1 -ConDatos`.

---

## 2. Requisitos en el servidor

- Windows Server 2016 o superior.
- **Python 3.10+** (<https://www.python.org/downloads/windows/>), marcando
  *"Add python.exe to PATH"*.
- Permisos de **Administrador** para ejecutar el instalador.
- Conectividad de red saliente hacia el ERP (Odoo) por HTTP/HTTPS.

---

## 3. Generar el paquete (en el equipo de DESARROLLO)

Desde la carpeta `deploy\` del proyecto:

```powershell
.\pack.ps1                 # genera deploy\dist\Kardex-ERP_<version>.zip
.\pack.ps1 -ConDatos       # ademas incluye las bases de datos actuales
.\pack.ps1 -Version 1.0.0  # version fija para el nombre del ZIP
```

Sube el ZIP resultante al servidor.

---

## 4. Instalar en el servidor

1. Copia y descomprime el ZIP, por ejemplo en `C:\Temp\Kardex-ERP`.
2. Abre **PowerShell como Administrador**.
3. Ejecuta:

```powershell
cd C:\Temp\Kardex-ERP\deploy
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

Por defecto instala en `C:\Kardex-ERP` y publica el puerto `8765`.

### Opciones utiles

| Parametro        | Descripcion                                                | Ejemplo                       |
|------------------|------------------------------------------------------------|-------------------------------|
| `-Destino`       | Carpeta de instalacion                                     | `-Destino "D:\Apps\Kardex"`   |
| `-Puerto`        | Puerto de escucha                                          | `-Puerto 8080`                |
| `-EnvPath`       | Ruta del `.env` con credenciales (fuera de la app)         | `-EnvPath "D:\Config\kardex.env"` |
| `-Tarea`         | Nombre de la tarea programada                              | `-Tarea "KardexERP"`          |
| `-SinFirewall`   | No crear la regla de firewall                              | `-SinFirewall`                |
| `-SinIniciar`    | No arrancar el servicio al terminar                        | `-SinIniciar`                 |

Ejemplo completo:

```powershell
.\install.ps1 -Destino "D:\Apps\Kardex" -Puerto 8080 -EnvPath "D:\Config\kardex.env"
```

---

## 5. Configurar las credenciales (obligatorio)

En la primera instalacion se crea un `.env` con valores de **ejemplo**. Editelo:

```
C:\Kardex-ERP\.env
```

```ini
odoo_url=https://erp.suempresa.com
odoo_db=nombre_de_la_base
odoo_username=usuario.servicio@suempresa.com
odoo_password=clave_de_servicio
```

Tras editar, reinicie el servicio:

```powershell
schtasks /End    /TN KardexERP
schtasks /Run    /TN KardexERP
```

> El `.env` contiene credenciales. Restrinja el acceso NTFS a ese archivo.
> Si prefiere guardarlo fuera de la carpeta de la app, use `-EnvPath` al instalar.

---

## 6. Operacion diaria

Acceso desde el navegador:

```
http://localhost:8765/            (en el propio servidor)
http://<IP-del-servidor>:8765/    (desde la red)
```

- **Log:** `C:\Kardex-ERP\logs\kardex_server.log`
- **Estado del servicio:** `schtasks /Query /TN KardexERP /V /FO LIST`
- **Reiniciar:**  `schtasks /End /TN KardexERP` y luego `schtasks /Run /TN KardexERP`
- El servicio arranca solo al encender el servidor y se reinicia ante fallos.

---

## 7. Actualizar una version

1. Genera un nuevo ZIP con `pack.ps1` y subelo al servidor.
2. Ejecuta (los datos y el `.env` se conservan):

```powershell
cd <nueva_carpeta>\deploy
powershell -ExecutionPolicy Bypass -File .\actualizar.ps1 -Destino "C:\Kardex-ERP"
```

`actualizar.ps1` detiene el servicio, copia la version nueva (sin sobrescribir
bases de datos ni `.env`) y lo vuelve a iniciar.

---

## 8. Desinstalar

```powershell
cd C:\Kardex-ERP\deploy
powershell -ExecutionPolicy Bypass -File .\uninstall.ps1
```

- Quita la tarea programada y la regla de firewall.
- **No** borra archivos ni datos salvo que se indique `-BorrarTodo`.

---

## 9. Variables de entorno reconocidas

| Variable           | Descripcion                                | Valor por defecto            |
|--------------------|--------------------------------------------|------------------------------|
| `KARDEX_HOST`      | Interfaz de escucha                        | `0.0.0.0`                    |
| `KARDEX_PORT`      | Puerto                                     | `8765`                       |
| `KARDEX_ENV_PATH`  | Ruta del `.env` con credenciales de Odoo   | `.env` junto a la app        |

---

## 10. Solucion de problemas

| Sintoma                                   | Causa / Solucion                                                       |
|-------------------------------------------|------------------------------------------------------------------------|
| El puerto no responde tras instalar       | Revise `logs\kardex_server.log`; suele faltar completar el `.env`.     |
| *"[Kardex] No se encontro ... .env"*      | El `.env` no existe o la ruta de `KARDEX_ENV_PATH` es incorrecta.      |
| *No se pudo autenticar contra el ERP*     | Credenciales de Odoo incorrectas o sin salida de red al ERP.           |
| `python` no se reconoce al instalar       | Reinstale Python marcando *"Add python.exe to PATH"* o use `py`.       |
| No hay acceso desde otras PCs             | Habilite la regla de firewall del puerto o reejecute el instalador.    |

---

## 11. Alternativa: ejecutar con PM2 (opcional)

En vez de la tarea programada de Windows, puede gestionar la app con **PM2**
(gestor de procesos de Node.js). Instala y registra todo automaticamente con:

```cmd
instalar_pm2.bat
```

Ese archivo, ejecutado **como Administrador**:

1. Instala **Node.js LTS** (si falta, via `winget`).
2. Instala **PM2** y **pm2-windows-startup** (arranque automatico al iniciar el servidor).
3. Registra la app: `pm2 start kardex_server.py --name kardex --interpreter python`
   con `KARDEX_HOST=0.0.0.0`, `KARDEX_PORT=8765` y `KARDEX_ENV_PATH=C:\Kardex-ERP\.env`.
4. Guarda el estado con `pm2 save`.

Ajuste al inicio del `.bat` las variables `APP_DIR`, `APP_NOMBRE`, `PUERTO` y
`ENV_PATH` si su instalacion es distinta.

Comandos utiles:

```cmd
pm2 status                 :: estado de los procesos
pm2 logs kardex            :: ver registros
pm2 restart kardex         :: reiniciar la app
pm2 save                   :: guardar la lista de procesos
```

> **Importante:** use **una sola** forma de arranque. Si instala PM2, desactive la
> tarea programada para no ocupar el mismo puerto dos veces:
> `schtasks /End /TN KardexERP` y `schtasks /Delete /TN KardexERP /F`.
