# Panel ANDE1723-EO — actualización diaria automática

## Qué hace esto
- `template.html` es el diseño del panel con placeholders (`{{TOTAL_ALIM}}`, etc.)
- `build.py` descarga la planilla `.xlsm` desde SharePoint, calcula los mismos
  números que vimos en el chat, y genera `index.html` reemplazando los placeholders.
- El módulo de lotes (dentro del HTML) sigue leyendo el CSV publicado de Google Sheets
  directo en el navegador — no necesita el script, ya se actualiza solo en cada visita.
- El workflow de GitHub Actions corre `build.py` todos los días a las 09:00 UTC,
  y si `index.html` cambió, hace commit y push. Netlify (conectado al repo) detecta
  ese push y republica el sitio solo.

## Pasos para dejarlo andando (una sola vez)

### 1. Subir este repo a GitHub
```bash
cd panel-ande1723
git init
git add .
git commit -m "Panel inicial con actualización automática"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/panel-ande1723.git
git push -u origin main
```

### 2. Cargar el secret con el link de descarga
En GitHub: **Settings → Secrets and variables → Actions → New repository secret**
- Nombre: `XLSM_URL`
- Valor: el link de SharePoint con `&download=1` al final, por ejemplo:
  ```
  https://cgisre-my.sharepoint.com/:x:/g/personal/amado_arrieta_cgis_com_py/IQDqMmbeT86UQKvLYAOxEniVAQ9d72bPySU7yjDjJeKzUqw?rtime=RcN51R3y3kg&download=1
  ```

Nunca pongas este link directo en el código — por eso va como secret.

### 3. Probar el workflow a mano (sin esperar al cron)
En GitHub: pestaña **Actions → Actualizar panel diario → Run workflow**.
Si falla, el log te va a decir por qué (por ejemplo, si el link de SharePoint
dejó de ser público — mirá la sección "Qué puede romperse" abajo).

### 4. Conectar Netlify al repo
**Netlify → Add new site → Import an existing project → GitHub** → elegís
`panel-ande1723` → Deploy site. A partir de acá cada push a `main` (incluido
el commit automático diario) republica el sitio solo.

## Qué puede romperse (y cómo lo vas a notar)

- **El link de SharePoint deja de ser público** (alguien lo revoca, cambia
  permisos, o Amado deja de tener acceso a esa carpeta): `build.py` corta
  con error y el workflow queda en rojo en la pestaña Actions — no rompe el
  sitio, `index.html` simplemente se queda con los datos del último día bueno.
  GitHub manda un email automático cuando un workflow programado falla.
- **Cambia el nombre de alguna hoja del xlsm** (ej. renombran "2. AVANCES"):
  el script también corta con error claro, mismo comportamiento de arriba.
- **El Sheet de lotes deja de estar publicado**: no afecta al build diario,
  pero el módulo 4 del panel va a mostrar el aviso rojo "sin conexión" en
  el navegador de quien lo abra.

## Estructura
```
panel-ande1723/
├── build.py                          # descarga + calcula + regenera index.html
├── template.html                     # diseño con placeholders
├── requirements.txt
├── .github/workflows/actualizar.yml  # cron diario
└── index.html                        # se genera solo, no editar a mano
```
