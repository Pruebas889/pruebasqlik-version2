# Qlik Sense Login Diagnostics

Este pequeño proyecto contiene utilidades para diagnosticar por qué no
puedes encontrar el formulario de login en una app de Qlik Sense (iframe,
Shadow DOM, SSO, modal dinámico, etc.).

Archivos:
- `qlik.py`: script principal que abre la URL, lista iframes, vuelca HTML y busca inputs comunes.
- `requirements.txt`: dependencias para instalar.

Instalación:
```powershell
python -m pip install -r requirements.txt
```

Uso básico:
```powershell
python qlik.py "https://qlik.copservir.com/sense/app/d39c40fb-a304-4eaf-9a30-50b7279d33f1/sheet/4f191cdb-aa40-409d-86b2-497a427a8b6a/state/analysis"
```

Recomendaciones rápidas:
- Si el login está en un `iframe`, el script volcará HTML de cada iframe en `debug_html/`.
- Si usa Shadow DOM, usa la función `try_shadow_query` como ejemplo para acceder mediante `execute_script`.
- Para evitar automatizar SSO repetidamente, inicia Chrome manualmente con un perfil y usa `--user-data-dir` o arranca Chrome con `--remote-debugging-port` y usa `--debugger 127.0.0.1:9222`.

Siguientes pasos sugeridos:
- Ejecuta el script y comparte los HTML volcados o la salida JSON si quieres que te ayude a identificar selectores.
- Si prefieres, puedo añadir intentos automáticos para loguear en proveedores comunes (MS, Okta), pero necesitaré más información.