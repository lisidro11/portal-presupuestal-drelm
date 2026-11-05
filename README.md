# Portal de Ejecución Presupuestal DRELM

Aplicación Streamlit con panel ejecutivo, análisis por dimensiones, ejecución mensual, descarga Excel, informe PDF y cuentas por instituto.

## Uso local (con paquete privado)
1. Ejecutar `INSTALAR_REQUISITOS.bat` y `INICIAR_APP.bat`.
2. Ingresar con la cuenta administradora o crearla en la primera ejecución.
3. El corte inicial y las cuentas institucionales se encuentran **solo en el paquete privado**, no en el repositorio público.

## GitHub + Streamlit
1. Crear **repositorio privado** en GitHub y subir **solo el contenido del paquete Git**.
2. Crear aplicación en [Streamlit Community Cloud](https://share.streamlit.io/) y seleccionar `app.py` como archivo principal.
3. Crear cuenta administradora y cargar desde el módulo inicial un Excel presupuestal enriquecido con columna `NOMBRE`.
4. Desde Administración, importar la matriz validada `NOMBRE`, `CODIGO_MODULAR` para crear cuentas de instituto.

### Advertencia crítica para producción
La aplicación usa SQLite y guarda los Excel en disco. **Streamlit Community Cloud no garantiza persistencia del almacenamiento local entre reinicios/reimplementaciones**. Por tanto, este despliegue es **solo de demostración**, no es adecuado para distribuir credenciales reales a institutos. Para producción, migrar autenticación, cuentas, historial y fuentes a una base y almacenamiento persistentes con copias de seguridad, y verificar autorización institucional mediante pruebas.

No subir al repositorio `datos/`, `fuentes/`, archivos Excel, contraseñas ni `secrets.toml`. GitHub Pages no ejecuta Streamlit.


VERSIÓN 9: 25 claves institucionales de 6 letras, credenciales en CSV CONFIDENCIAL separado del ZIP. Las claves anteriores ya no funcionan. El administrador conserva clave robusta. No publicar el ZIP local en GitHub.
